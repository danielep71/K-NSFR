"""Tests of check_workflows.py, and of pr-milestones.yml's script against a fake API."""

from __future__ import annotations

import contextlib
import copy
import io
import os
import re
import shutil
import tempfile
import unittest
from collections.abc import Callable
from pathlib import Path
from typing import Any
from unittest import mock

import check
import check_workflows as cw

ROOT = Path(__file__).resolve().parents[1]
YAML: Any = cw.load_yaml()


def setUpModule() -> None:
    if YAML is None:
        if os.environ.get("GITHUB_ACTIONS"):
            raise RuntimeError("PyYAML is required on GitHub Actions "
                               "(tools/requirements-quality-ci.txt)")
        raise unittest.SkipTest("PyYAML is not installed; skipped locally only")


class Repository:
    """A mutable copy of the repository's workflows, Dependabot file and documents."""

    def __init__(self) -> None:
        workflows, dependabot, governance, settings, tooling = cw.load(ROOT, YAML)
        self.workflows: dict[str, Any] = copy.deepcopy(workflows)
        self.dependabot: Any = copy.deepcopy(dependabot)
        self.governance = governance
        self.settings = settings
        self.tooling = tooling

    def job(self, workflow: str, job: str) -> Any:
        return self.workflows[workflow]["jobs"][job]

    def on(self, workflow: str) -> Any:
        return cw.triggers(self.workflows[workflow])

    def findings(self) -> list[str]:
        return cw.findings_for(self.workflows, self.dependabot, self.governance, self.settings,
                               self.tooling)


class WorkflowInvariantTests(unittest.TestCase):
    def setUp(self) -> None:
        self.repo = Repository()

    def assertFinding(self, fragment: str) -> None:
        findings = self.repo.findings()
        self.assertTrue(any(fragment in finding for finding in findings),
                        f"no finding containing {fragment!r} in {findings}")

    def test_repository_passes(self) -> None:
        self.assertEqual(self.repo.findings(), [])

    def test_every_workflow_is_read(self) -> None:
        names = {path.name for path in cw.workflow_files(ROOT)}
        self.assertEqual(set(self.repo.workflows), names)
        self.assertIn(cw.MILESTONES_FILE, names)

    # Actions and checkout
    def test_unpinned_action(self) -> None:
        step = self.repo.job("repository-integrity.yml", "integrity")["steps"][0]
        step["uses"] = "actions/checkout@v7"
        self.assertFinding("`actions/checkout@v7` is not pinned")

    def test_short_sha(self) -> None:
        self.repo.job("labels-drift.yml", "detect")["steps"][0]["uses"] = "actions/checkout@3d3c42e"
        self.assertFinding("not pinned to a 40-character SHA")

    def test_reusable_workflow_must_be_pinned(self) -> None:
        self.repo.workflows["labels-drift.yml"]["jobs"]["reuse"] = {
            "uses": "owner/repo/.github/workflows/x.yml@main", "permissions": {},
            "timeout-minutes": 5}
        self.assertFinding("job reuse: `owner/repo/.github/workflows/x.yml@main`")

    def test_checkout_persisting_credentials(self) -> None:
        step = self.repo.job("labels-sync.yml", "validate")["steps"][0]
        del step["with"]["persist-credentials"]
        self.assertFinding("actions/checkout without persist-credentials: false")
        step["with"]["persist-credentials"] = "false"
        self.assertFinding("actions/checkout without persist-credentials: false")

    # Permissions and timeouts
    def test_missing_workflow_permissions(self) -> None:
        del self.repo.workflows["issue-triage.yml"]["permissions"]
        self.assertFinding("issue-triage.yml: no workflow-level permissions")

    def test_workflow_level_write(self) -> None:
        self.repo.workflows["labels-sync.yml"]["permissions"] = {"contents": "write"}
        self.assertFinding("labels-sync.yml: workflow-level write permission ['contents']")

    def test_missing_timeout(self) -> None:
        del self.repo.job("labels-sync.yml", "validate")["timeout-minutes"]
        self.assertFinding("labels-sync.yml: job validate has no timeout-minutes")

    def test_extra_write_permission(self) -> None:
        job = self.repo.job("repository-integrity.yml", "integrity")
        job["permissions"]["issues"] = "write"
        self.assertFinding("job integrity has write permission ['issues']")

    def test_write_outside_the_allowed_scope(self) -> None:
        self.repo.job("labels-sync.yml", "reconcile")["permissions"]["contents"] = "write"
        self.assertFinding("job reconcile has write permission ['contents']")

    def test_traffic_alert_cannot_write_contents(self) -> None:
        self.repo.job("traffic-history.yml", "alert")["permissions"]["contents"] = "write"
        self.assertFinding("traffic-history.yml: job alert has write permission ['contents']")

    def traffic_step(self, job: str, predicate: Callable[[Any], bool]) -> Any:
        return next(step for step in self.repo.job("traffic-history.yml", job)["steps"]
                    if predicate(step))

    def token_step(self) -> Any:
        return self.traffic_step("collect", lambda s: str(s.get("uses", "")).startswith(cw.TOKEN_ACTION))

    def test_the_repository_traffic_workflow_passes(self) -> None:
        self.assertEqual(cw.traffic_findings(self.repo.workflows), [])

    def test_the_app_token_has_only_administration_read_on_this_repository(self) -> None:
        for key, value in (("permission-administration", "write"), ("permission-contents", "read"),
                           ("repositories", "K-NSFR,other")):
            with self.subTest(key=key):
                self.repo = Repository()
                self.token_step()["with"][key] = value
                self.assertFinding("the App token step must be `app-token` with exactly")

    def test_the_private_key_reaches_only_the_token_step(self) -> None:
        fetch = self.traffic_step("collect", lambda s: s.get("id") == "fetch")
        fetch["env"]["KEY"] = "${{ secrets.TRAFFIC_APP_PRIVATE_KEY }}"
        self.assertFinding("secrets used outside the App token step: ['secrets.TRAFFIC_APP_PRIVATE_KEY']")

    def test_the_preflight_sees_only_whether_the_key_is_set(self) -> None:
        preflight = self.traffic_step("collect", lambda s: "HAS_KEY" in (s.get("env") or {}))
        preflight["env"]["HAS_KEY"] = "${{ secrets.TRAFFIC_APP_PRIVATE_KEY }}"
        self.assertFinding("a preflight before the token step must test only whether")
        self.assertFinding("secrets used outside the App token step")

    def test_only_the_fetch_step_receives_the_token(self) -> None:
        merge = self.traffic_step("collect", lambda s: s.get("name") == "Merge into the history")
        merge["env"]["LEAK"] = "${{ steps.app-token.outputs.token }}"
        self.assertFinding("only the fetch step may receive the App token")

    def test_the_alert_reads_the_published_commit(self) -> None:
        checkout = self.traffic_step("alert", lambda s: (s.get("with") or {}).get("path") == "history")
        checkout["with"]["ref"] = "traffic-history"
        self.assertFinding("the alert job must check out the history at")

    def test_the_alert_validates_the_sha_before_its_checkout(self) -> None:
        steps = self.repo.job("traffic-history.yml", "alert")["steps"]
        steps.append(steps.pop(0))
        self.assertFinding("the alert job must validate the published SHA before its checkout")

    def test_the_alert_has_no_environment_or_secret(self) -> None:
        self.repo.job("traffic-history.yml", "alert")["environment"] = "analytics"
        self.assertFinding("the alert job must have no environment and no secret")

    def test_collect_publishes_the_history_sha(self) -> None:
        del self.repo.job("traffic-history.yml", "collect")["outputs"]
        self.assertFinding("collect must output history-sha from the publish step")

    def test_traffic_collect_cannot_write_issues(self) -> None:
        self.repo.job("traffic-history.yml", "collect")["permissions"]["issues"] = "write"
        self.assertFinding("traffic-history.yml: job collect has write permission ['issues']")

    def test_write_all(self) -> None:
        self.repo.job("labels-sync.yml", "reconcile")["permissions"] = "write-all"
        self.assertFinding("job reconcile has write permission ['*']")

    def test_allowed_writes_pass(self) -> None:
        for (workflow, job), scopes in cw.ALLOWED_WRITES.items():
            granted = cw.writes(self.repo.job(workflow, job)["permissions"])
            self.assertEqual(granted, set(scopes), f"{workflow}/{job}")

    # The required check
    def test_required_job_with_if(self) -> None:
        self.repo.job("repository-integrity.yml", "integrity")["if"] = "true"
        self.assertFinding("required job integrity has `if`")

    def test_required_job_with_strategy(self) -> None:
        job = self.repo.job("repository-integrity.yml", "integrity")
        job["strategy"] = {"matrix": {"python": ["3.13"]}}
        self.assertFinding("required job integrity has `strategy`")

    def test_path_filter_on_pull_request(self) -> None:
        self.repo.on("repository-integrity.yml")["pull_request"]["paths"] = ["src/**"]
        self.assertFinding("`pull_request` filters are")

    def test_paths_ignore_on_push(self) -> None:
        self.repo.on("repository-integrity.yml")["push"]["paths-ignore"] = ["docs/**"]
        self.assertFinding("`push` filters are")

    def test_branch_filter_changed(self) -> None:
        self.repo.on("repository-integrity.yml")["pull_request"]["branches"] = ["main"]
        self.assertFinding("`pull_request` filters are")

    def test_required_check_renamed(self) -> None:
        self.repo.job("repository-integrity.yml", "integrity")["name"] = "Integrity"
        self.assertFinding("0 jobs named `Repository integrity`")

    def test_rulesets_require_another_name(self) -> None:
        self.repo.settings = self.repo.settings.replace(
            "| Required status context | `Repository integrity` |",
            "| Required status context | `Integrity` |")
        self.assertFinding("the rulesets require ['Integrity', 'Repository integrity']")

    def test_governance_must_name_the_required_check_and_branches(self) -> None:
        self.repo.governance = self.repo.governance.replace("| Required status check |", "| Status check |")
        self.assertFinding("no 'Required status check' row naming the job")
        self.repo.governance = self.repo.governance.replace("Applies to `main`", "For `main`")
        self.assertFinding("no 'Applies to ... branch' line for the rulesets")

    # labels-sync and pr-milestones
    def test_labels_sync_condition(self) -> None:
        job = self.repo.job("labels-sync.yml", "reconcile")
        condition = job["if"]
        for changed in ("github.event_name != 'pull_request'", condition + " || true",
                        f"!({condition})"):
            with self.subTest(condition=changed):
                job["if"] = changed
                self.assertFinding("labels-sync.yml: reconcile's condition must be")

    def test_yaml_extension_workflows_are_read(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            shutil.copytree(ROOT / cw.WORKFLOWS, root / cw.WORKFLOWS)
            (root / cw.WORKFLOWS / "extra.yaml").write_text(
                "on: push\npermissions: {}\njobs:\n  x:\n    runs-on: ubuntu-24.04\n"
                "    steps:\n      - uses: actions/checkout@v7\n", encoding="utf-8")
            workflows = {path.name for path in cw.workflow_files(root)}
            self.assertIn("extra.yaml", workflows)
            loaded = {path.name: YAML.load(path.read_text(encoding="utf-8"), Loader=cw.strict_loader(YAML))
                      for path in cw.workflow_files(root)}
            findings = cw.action_findings("extra.yaml", loaded["extra.yaml"])
            findings += cw.permission_findings("extra.yaml", loaded["extra.yaml"])
            self.assertIn("extra.yaml: job x step 1: `actions/checkout@v7` is not pinned to a 40-character SHA",
                          findings)
            self.assertIn("extra.yaml: job x has no timeout-minutes", findings)

    def test_runtimes_are_exact_and_documented(self) -> None:
        setup = next(step for step in self.repo.job("repository-integrity.yml", "integrity")["steps"]
                     if "python-version" in (step.get("with") or {}))
        setup["with"]["python-version"] = "3.13"
        self.assertFinding("python-version `3.13` must be the exact version")
        self.repo = Repository()
        self.repo.tooling = self.repo.tooling.replace("| Node.js | 22.23.4 |", "| Node.js | 22 |")
        self.assertFinding("the Node.js row must give an exact X.Y.Z version")

    def test_triage_runs_on_milestone_events(self) -> None:
        self.repo.on("issue-triage.yml")["milestone"]["types"].remove("edited")
        self.assertFinding("issue-triage.yml: milestone events must be")
        del self.repo.on("issue-triage.yml")["milestone"]
        self.assertFinding("issue-triage.yml: milestone events must be")

    def test_milestones_queue_per_pull_request(self) -> None:
        self.repo.workflows["pr-milestones.yml"]["concurrency"]["group"] = "pr-milestones"
        self.assertFinding("pr-milestones.yml: concurrency must be")
        self.repo = Repository()
        self.repo.workflows["pr-milestones.yml"]["concurrency"]["cancel-in-progress"] = True
        self.assertFinding("pr-milestones.yml: concurrency must be")

    def test_integrity_cancels_only_superseded_pull_request_runs(self) -> None:
        self.repo.workflows["repository-integrity.yml"]["concurrency"] = {
            "group": "integrity-${{ github.workflow }}-${{ github.ref }}", "cancel-in-progress": True}
        self.assertFinding("repository-integrity.yml: concurrency must be")

    # CodeQL
    def test_codeql_analyzes_every_supported_language(self) -> None:
        self.repo.job("codeql.yml", "analyze")["strategy"]["matrix"]["language"].remove(
            "javascript-typescript")
        self.assertFinding("codeql.yml: the matrix must be exactly language:")
        self.repo = Repository()
        self.repo.job("codeql.yml", "analyze")["strategy"]["matrix"]["include"] = [
            {"language": "csharp"}]
        self.assertFinding("codeql.yml: the matrix must be exactly language:")

    def test_codeql_uses_the_extended_queries(self) -> None:
        init = self.repo.job("codeql.yml", "analyze")["steps"][1]
        init["with"]["queries"] = "security-and-quality"
        self.assertFinding("codeql.yml: exactly one init step")
        del init["with"]["queries"]
        self.assertFinding("codeql.yml: exactly one init step")

    def test_codeql_uploads_one_category_per_language(self) -> None:
        del self.repo.job("codeql.yml", "analyze")["steps"][2]["with"]
        self.assertFinding("codeql.yml: exactly one analyze step")
        self.repo = Repository()
        del self.repo.job("codeql.yml", "analyze")["steps"][2]
        self.assertFinding("codeql.yml: exactly one analyze step")

    def test_codeql_actions_share_one_commit(self) -> None:
        self.repo.job("codeql.yml", "analyze")["steps"][2]["uses"] = (
            "github/codeql-action/analyze@" + "0" * 40)
        self.assertFinding("codeql.yml: init and analyze must be pinned to the same commit")

    def test_codeql_triggers_and_concurrency(self) -> None:
        self.repo.on("codeql.yml")["pull_request"]["paths"] = ["tools/**"]
        self.assertFinding("codeql.yml: must run on push and pull_request")
        self.repo = Repository()
        del self.repo.on("codeql.yml")["schedule"]
        self.assertFinding("codeql.yml: must run on push and pull_request")
        self.repo = Repository()
        self.repo.workflows["codeql.yml"]["concurrency"]["cancel-in-progress"] = True
        self.assertFinding("codeql.yml: concurrency must be")

    def test_codeql_writes_only_security_events(self) -> None:
        self.repo.job("codeql.yml", "analyze")["permissions"]["actions"] = "read"
        self.assertFinding("codeql.yml: analyze's permissions must be")
        self.repo = Repository()
        self.repo.job("codeql.yml", "analyze")["permissions"]["contents"] = "write"
        self.assertFinding("job analyze has write permission ['contents'] outside the allowed list")

    def test_codeql_pr_is_read_only(self) -> None:
        self.repo.job("codeql.yml", "analyze-pr")["permissions"]["security-events"] = "write"
        self.assertFinding("codeql.yml: analyze-pr's permissions must be")

    def test_codeql_event_paths_are_separate(self) -> None:
        for name in ("analyze", "analyze-pr"):
            self.repo.job("codeql.yml", name)["if"] = "always()"
            self.assertFinding(f"codeql.yml: {name}'s condition must be")

    def test_codeql_pr_never_uploads_results_or_databases(self) -> None:
        inputs = self.repo.job("codeql.yml", "analyze-pr")["steps"][2]["with"]
        for key, value in (("upload", "always"), ("upload-database", True)):
            original = inputs[key]
            inputs[key] = value
            self.assertFinding("codeql.yml: exactly one analyze step")
            inputs[key] = original

    def test_codeql_pr_retains_sarif(self) -> None:
        artifact = self.repo.job("codeql.yml", "analyze-pr")["steps"][3]
        artifact["with"]["retention-days"] = 1
        self.assertFinding("codeql.yml: PR SARIF must be retained for 30 days")

    def test_codeql_manual_trigger_is_required(self) -> None:
        del self.repo.on("codeql.yml")["workflow_dispatch"]
        self.assertFinding("codeql.yml: must run on push and pull_request")

    def test_codeql_workflow_is_required(self) -> None:
        del self.repo.workflows["codeql.yml"]
        self.assertFinding("codeql.yml: not found")

    def test_integrity_fetches_the_history_for_host_evidence(self) -> None:
        del self.repo.job("repository-integrity.yml", "integrity")["steps"][0]["with"]["fetch-depth"]
        self.assertFinding("must check out with fetch-depth: 0")

    def test_link_check_never_gates_a_change(self) -> None:
        for event in ("pull_request", "push"):
            with self.subTest(event=event):
                self.repo = Repository()
                self.repo.on("external-links.yml")[event] = {"branches": ["main"]}
                self.assertFinding("external-links.yml: must run only on")
        self.repo = Repository()
        del self.repo.on("external-links.yml")["schedule"]
        self.assertFinding("external-links.yml: must run only on")
        self.repo = Repository()
        del self.repo.workflows["external-links.yml"]
        self.assertFinding("external-links.yml: not found")

    def test_label_reconcile_reads_pull_requests(self) -> None:
        del self.repo.job("labels-sync.yml", "reconcile")["permissions"]["pull-requests"]
        self.assertFinding("labels-sync.yml: reconcile's permissions must be")

    def test_milestones_events(self) -> None:
        self.repo.on("pr-milestones.yml")["pull_request"] = None
        self.assertFinding("pr-milestones.yml: events changed")

    def test_milestones_workflow_permissions(self) -> None:
        self.repo.workflows["pr-milestones.yml"]["permissions"] = {"contents": "read"}
        self.assertFinding("pr-milestones.yml: workflow-level permissions must be {}")

    def test_milestones_job_permissions(self) -> None:
        self.repo.job("pr-milestones.yml", "reconcile")["permissions"]["contents"] = "read"
        self.assertFinding("reconcile's permissions must be pull-requests: write")

    def test_milestones_condition(self) -> None:
        self.repo.job("pr-milestones.yml", "reconcile")["if"] = "always()"
        self.assertFinding("pr-milestones.yml: reconcile's condition must be")

    def test_milestones_second_step(self) -> None:
        steps = self.repo.job("pr-milestones.yml", "reconcile")["steps"]
        steps.insert(0, {"uses": "actions/checkout@" + "0" * 40,
                         "with": {"persist-credentials": False}})
        self.assertFinding("reconcile must have a single inline `run` step")

    def test_milestones_expression_in_code(self) -> None:
        step = self.repo.job("pr-milestones.yml", "reconcile")["steps"][0]
        step["run"] = step["run"].replace("import json", "import json  # ${{ github.head_ref }}")
        self.assertFinding("the inline code contains `${{`")

    # Dependabot
    def test_dependabot_duplicate_ecosystem(self) -> None:
        self.repo.dependabot["updates"].append(copy.deepcopy(self.repo.dependabot["updates"][1]))
        self.assertFinding("`pip` is covered 2 times")

    def test_dependabot_missing_ecosystem(self) -> None:
        del self.repo.dependabot["updates"][0]
        self.assertFinding("`github-actions` is covered 0 times")

    def test_dependabot_milestone(self) -> None:
        self.repo.dependabot["updates"][0]["milestone"] = 4
        self.assertFinding("`github-actions` sets a milestone")

    def test_dependabot_priorities(self) -> None:
        self.repo.dependabot["updates"][1]["labels"].append("P2")
        self.assertFinding("`pip` has priorities ['P2', 'P3']")
        self.repo.dependabot["updates"][1]["labels"] = ["ci"]
        self.assertFinding("`pip` has priorities []")


def quiet_main() -> int:
    with contextlib.redirect_stdout(io.StringIO()):
        return cw.main(["--root", str(ROOT)])


class StrictLoaderTests(unittest.TestCase):
    def test_duplicate_key_is_rejected(self) -> None:
        text = "permissions:\n  issues: write\n  contents: read\n  issues: read\n"
        with self.assertRaisesRegex(YAML.YAMLError, r"duplicate key\(s\) \['issues'\]"):
            YAML.load(text, Loader=cw.strict_loader(YAML))

    def test_on_key_is_read(self) -> None:
        workflow = YAML.load("on:\n  push: {}\n", Loader=cw.strict_loader(YAML))
        self.assertEqual(cw.triggers(workflow), {"push": {}})


class MissingYamlTests(unittest.TestCase):
    def test_fails_on_github_actions(self) -> None:
        with mock.patch.object(cw, "load_yaml", return_value=None), \
                mock.patch.dict(os.environ, {"GITHUB_ACTIONS": "true"}):
            self.assertEqual(quiet_main(), 1)

    def test_skipped_locally(self) -> None:
        env = {key: value for key, value in os.environ.items() if key != "GITHUB_ACTIONS"}
        with mock.patch.object(cw, "load_yaml", return_value=None), \
                mock.patch.dict(os.environ, env, clear=True):
            self.assertEqual(quiet_main(), 2)

    def test_gate_runs_on_github_actions(self) -> None:
        with mock.patch("importlib.util.find_spec", return_value=None):
            with mock.patch.dict(os.environ, {"GITHUB_ACTIONS": "true"}):
                self.assertTrue(check.available("pyyaml"))
            env = {key: value for key, value in os.environ.items() if key != "GITHUB_ACTIONS"}
            with mock.patch.dict(os.environ, env, clear=True):
                self.assertFalse(check.available("pyyaml"))


REPOSITORY = "owner/repo"


class FakeGitHub:
    """The issues and pulls endpoints the milestone script uses, over live state."""

    def __init__(self, prs: dict[int, Any], issues: dict[int, Any]) -> None:
        self.items: dict[int, dict[str, Any]] = {}
        for number, milestone in prs.items():
            self.items[number] = {"number": number, "milestone": milestone, "pull_request": {}}
        for number, milestone in issues.items():
            self.items[number] = {"number": number, "milestone": milestone}
        self.patched: list[int] = []
        self.pages: list[int] = []

    def __call__(self, method: str, path: str, data: Any = None) -> Any:
        item = re.fullmatch(rf"/repos/{REPOSITORY}/issues/(\d+)", path)
        if item:
            number = int(item.group(1))
            if method == "PATCH":
                self.patched.append(number)
                self.items[number].update(data)
            return copy.deepcopy(self.items[number])
        if method == "GET" and path.startswith(f"/repos/{REPOSITORY}/pulls?"):
            # [?&]page= : a bare page= would match per_page= first.
            page = int(re.search(r"[?&]page=(\d+)", path).group(1))  # type: ignore[union-attr]
            size = int(re.search(r"[?&]per_page=(\d+)", path).group(1))  # type: ignore[union-attr]
            self.pages.append(page)
            prs = sorted(n for n, i in self.items.items() if "pull_request" in i)
            return [{"number": n, "milestone": self.items[n]["milestone"]}
                    for n in prs[(page - 1) * size:page * size]]
        raise AssertionError(f"unexpected request {method} {path}")

    def milestones(self, *, pulls: bool) -> dict[int, Any]:
        return {n: i["milestone"] for n, i in self.items.items()
                if ("pull_request" in i) == pulls}


class MilestoneScriptTests(unittest.TestCase):
    def setUp(self) -> None:
        workflow = Repository().workflows[cw.MILESTONES_FILE]
        source = cw.milestone_script(workflow)
        assert source is not None
        namespace: dict[str, Any] = {"__name__": "pr_milestones"}
        exec(compile(source, cw.MILESTONES_FILE, "exec"), namespace)
        self.script: Callable[[Any, str, Any], Any] = namespace["reconcile"]
        self.summary: Callable[[Any], str] = namespace["summary"]

    def reconcile(self, api: FakeGitHub, repository: str, event: Any) -> None:
        with contextlib.redirect_stdout(io.StringIO()):
            self.script(api, repository, event)

    def test_pull_request_event_clears_its_milestone(self) -> None:
        api = FakeGitHub({5: {"number": 2}, 6: {"number": 2}}, {})
        self.reconcile(api, REPOSITORY, {"number": 5, "pull_request": {}})
        self.assertEqual(api.patched, [5])
        self.assertEqual(api.milestones(pulls=True), {5: None, 6: {"number": 2}})

    def test_removal_requires_independent_read_back(self) -> None:
        calls: list[str] = []

        def request(method: str, path: str, data: Any = None) -> Any:
            calls.append(method)
            if method == "PATCH":
                return {"milestone": None}
            # PATCH claims success but the independent GET still sees the milestone.
            return {"pull_request": {}, "milestone": {"number": 2}}

        with self.assertRaisesRegex(ValueError, "after read-back"):
            self.script(request, REPOSITORY, {"number": 5, "pull_request": {}})
        self.assertEqual(calls, ["GET", "PATCH", "GET"])

    def test_patch_response_must_confirm_removal(self) -> None:
        def request(method: str, path: str, data: Any = None) -> Any:
            return {"pull_request": {}, "milestone": {"number": 2}}

        with self.assertRaisesRegex(ValueError, "still has a milestone"):
            self.script(request, REPOSITORY, {"number": 5, "pull_request": {}})

    def test_read_back_failure_is_not_success(self) -> None:
        calls: list[str] = []

        def request(method: str, path: str, data: Any = None) -> Any:
            calls.append(method)
            if method == "PATCH":
                return {"milestone": None}
            if len(calls) == 3:
                raise OSError("GET failed")
            return {"pull_request": {}, "milestone": {"number": 2}}

        with self.assertRaisesRegex(OSError, "GET failed"):
            self.script(request, REPOSITORY, {"number": 5, "pull_request": {}})

    def test_pull_request_without_milestone_is_left_alone(self) -> None:
        api = FakeGitHub({5: None}, {})
        self.reconcile(api, REPOSITORY, {"number": 5, "pull_request": {}})
        self.assertEqual(api.patched, [])

    def test_scan_clears_every_page_and_no_issue(self) -> None:
        milestone = {"number": 4}
        prs = {n: (milestone if n % 40 == 7 else None) for n in range(1, 251)}
        issues = {n: milestone for n in range(251, 261)}
        api = FakeGitHub(prs, issues)
        self.reconcile(api, REPOSITORY, {"schedule": "17 3 * * *"})
        self.assertEqual(api.pages, [1, 2, 3])
        self.assertEqual(api.patched, [n for n in prs if prs[n] is not None])
        self.assertTrue(any(n > 200 for n in api.patched), "the third page was not cleared")
        self.assertTrue(all(value is None for value in api.milestones(pulls=True).values()))
        self.assertEqual(api.milestones(pulls=False), issues)

    def test_removed_milestones_are_returned_for_the_summary(self) -> None:
        api = FakeGitHub({5: {"number": 2, "title": "v0.1.0 - Repo setup"}, 6: None}, {})
        with contextlib.redirect_stdout(io.StringIO()):
            removed = self.script(api, REPOSITORY, {"schedule": "17 3 * * *"})
        self.assertEqual(removed, [(5, "v0.1.0 - Repo setup")])

    def test_summary_lists_each_removal_and_treats_titles_as_data(self) -> None:
        text = self.summary([(5, "a|b"), (7, "back\\slash\r\nnew row")])
        self.assertIn("| #5 | a\\|b |", text)
        self.assertIn("| #7 | back\\\\slash new row |", text)
        self.assertEqual(text.count("\n| #"), 2)
        self.assertIn("No pull request had a milestone.", self.summary([]))

    def test_issue_number_is_refused(self) -> None:
        api = FakeGitHub({5: {"number": 2}}, {9: {"number": 2}})
        with self.assertRaisesRegex(ValueError, "Refusing to change an issue"):
            self.reconcile(api, REPOSITORY, {"number": 9, "pull_request": {}})
        self.assertEqual(api.patched, [])
        self.assertEqual(api.milestones(pulls=False), {9: {"number": 2}})


if __name__ == "__main__":
    unittest.main()
