#!/usr/bin/env python3
"""Check the security and governance invariants of the workflows and Dependabot.

Reads .github/workflows/*.yml and .github/dependabot.yml with PyYAML and
reports every finding (docs/GOVERNANCE.md; tools/README.md, workflow-invariants).
Read-only. Without PyYAML it fails when GITHUB_ACTIONS is set and exits 2
(skipped) otherwise, so that only a local run can skip it.
"""

from __future__ import annotations

import argparse
import importlib
import importlib.util
import os
import re
import sys
from collections import Counter
from collections.abc import Iterator, Mapping, Sequence
from pathlib import Path
from typing import Any

WORKFLOWS = Path(".github/workflows")
DEPENDABOT = Path(".github/dependabot.yml")
GOVERNANCE = Path("docs/GOVERNANCE.md")
SETTINGS = Path("docs/REPOSITORY_SETTINGS.md")
TOOLING = Path("tools/README.md")

SHA_PINNED = re.compile(r"^[A-Za-z0-9_.-]+/[A-Za-z0-9_./-]+@[0-9a-f]{40}$")
DEFAULT_BRANCH_REF = "github.ref == format('refs/heads/{0}', github.event.repository.default_branch)"

# The only write permissions a job may hold: (workflow file, job id) -> scopes.
ALLOWED_WRITES: dict[tuple[str, str], frozenset[str]] = {
    ("labels-sync.yml", "reconcile"): frozenset({"issues"}),
    ("traffic-history.yml", "collect"): frozenset({"contents"}),
    ("traffic-history.yml", "alert"): frozenset({"issues"}),
    ("pr-milestones.yml", "reconcile"): frozenset({"pull-requests"}),
    ("codeql.yml", "analyze"): frozenset({"security-events"}),
}

LABELS_RECONCILE_IF = "github.event_name != 'pull_request' && " + DEFAULT_BRANCH_REF
# The prune check reads the issues API, which also lists pull requests.
LABELS_RECONCILE_PERMISSIONS = {"contents": "read", "issues": "write", "pull-requests": "read"}

# A pull request's superseded run is cancelled; every push or manual run keeps its own.
INTEGRITY_CONCURRENCY = {
    "group": "integrity-${{ github.event_name == 'pull_request' && "
             "format('pr-{0}', github.event.pull_request.number) || "
             "format('run-{0}', github.run_id) }}",
    "cancel-in-progress": "${{ github.event_name == 'pull_request' }}",
}

CODEQL_FILE = "codeql.yml"
CODEQL_ACTION = re.compile(r"^github/codeql-action/(init|analyze)@([0-9a-f]{40})$")
# Every language the repository contains that CodeQL supports; VBA has no extractor.
CODEQL_LANGUAGES = ["actions", "python", "javascript-typescript"]
CODEQL_INIT = {"languages": "${{ matrix.language }}", "build-mode": "none",
               "queries": "security-extended"}
CODEQL_ANALYZE = {"category": "/language:${{ matrix.language }}"}
CODEQL_PERMISSIONS = {"contents": "read", "security-events": "write"}
CODEQL_CONCURRENCY = {
    "group": INTEGRITY_CONCURRENCY["group"].replace("integrity-", "codeql-", 1),
    "cancel-in-progress": INTEGRITY_CONCURRENCY["cancel-in-progress"],
}

LINKS_FILE = "external-links.yml"
# The network is not reproducible: the link check never gates a change.
LINKS_EVENTS = {"schedule", "workflow_dispatch"}

TRIAGE_FILE = "issue-triage.yml"
# A changed due date, or a closed or deleted milestone, changes the audit.
TRIAGE_MILESTONE_EVENTS = {"types": ["created", "edited", "closed", "opened", "deleted"]}

MILESTONES_FILE = "pr-milestones.yml"
MILESTONES_EVENTS: dict[str, Any] = {
    "push": {"branches": ["main"], "paths": [".github/workflows/pr-milestones.yml"]},
    "pull_request_target": {"types": ["opened", "reopened", "edited", "milestoned"]},
    "schedule": [{"cron": "17 3 * * *"}],
    "workflow_dispatch": None,
}
# One queue per pull request for its own events; full scans share another.
MILESTONES_CONCURRENCY = {
    "group": "pr-milestones-${{ github.event_name == 'pull_request_target' && "
             "format('pr-{0}', github.event.pull_request.number) || 'scan' }}",
    "cancel-in-progress": False,
}
MILESTONES_IF = ("github.event_name == 'pull_request_target' || "
                 "github.event_name == 'schedule' || " + DEFAULT_BRANCH_REF)
SCRIPT = re.compile(r"<<'PY'\n(.*)\nPY\s*$", re.S)

# setup-* action input -> the Tool versions row of tools/README.md that records it.
RUNTIME_INPUTS = {"python-version": "Python", "node-version": "Node.js"}
EXACT_VERSION = re.compile(r"^\d+\.\d+\.\d+$")

TRAFFIC_FILE = "traffic-history.yml"
TOKEN_ACTION = "actions/create-github-app-token@"
# The App token step's inputs: the App, its key, this repository only, and only
# the one permission the traffic API needs.
TOKEN_INPUTS = {
    "client-id": "${{ vars.TRAFFIC_APP_CLIENT_ID }}",
    "private-key": "${{ secrets.TRAFFIC_APP_PRIVATE_KEY }}",
    "repositories": "${{ github.event.repository.name }}",
    "permission-administration": "read",
}
KEY_PRESENT = "${{ secrets.TRAFFIC_APP_PRIVATE_KEY != '' }}"
TRAFFIC_TOKEN_ENV = "${{ steps.app-token.outputs.token }}"
HISTORY_SHA = "${{ needs.collect.outputs.history-sha }}"
SECRET = re.compile(r"secrets\.[A-Za-z_]+")

# Ecosystems whose versions are pinned here, and where Dependabot must look.
PINNED_ECOSYSTEMS = {"github-actions": "/", "pip": "/tools"}
PRIORITY = re.compile(r"^P\d$")

Workflow = Mapping[Any, Any]


def load_yaml() -> Any | None:
    """PyYAML's module, or None when it is not installed."""
    if importlib.util.find_spec("yaml") is None:
        return None
    return importlib.import_module("yaml")


def strict_loader(yaml: Any) -> Any:
    """A SafeLoader that rejects a duplicate key, which PyYAML would silently
    resolve to its last value (a hidden `issues: write`, for example)."""

    class StrictLoader(yaml.SafeLoader):  # type: ignore[misc]
        def construct_mapping(self, node: Any, deep: bool = False) -> Any:
            keys = [self.construct_object(key, deep=deep) for key, _ in node.value]
            duplicates = sorted({str(key) for key in keys if keys.count(key) > 1})
            if duplicates:
                raise yaml.constructor.ConstructorError(
                    None, None, f"duplicate key(s) {duplicates}", node.start_mark)
            return super().construct_mapping(node, deep=deep)

    return StrictLoader


def triggers(workflow: Workflow) -> Any:
    """The `on:` mapping; YAML 1.1 reads an unquoted `on` key as True."""
    return workflow.get("on", workflow.get(True))


def squash(text: Any) -> str:
    return " ".join(str(text).split())


def jobs(workflow: Workflow) -> Iterator[tuple[str, Mapping[str, Any]]]:
    for job_id, job in (workflow.get("jobs") or {}).items():
        yield str(job_id), job if isinstance(job, Mapping) else {}


def writes(permissions: Any) -> set[str]:
    """The scopes a permissions value grants write access to; "*" for write-all."""
    if isinstance(permissions, str):
        return {"*"} if permissions == "write-all" else set()
    if isinstance(permissions, Mapping):
        return {str(scope) for scope, level in permissions.items() if level == "write"}
    return set()


def action_findings(name: str, workflow: Workflow) -> list[str]:
    findings = []
    for job_id, job in jobs(workflow):
        steps = job.get("steps") or []
        uses = [(f"job {job_id}", job["uses"])] if "uses" in job else []
        uses += [(f"job {job_id} step {i}", step["uses"]) for i, step in enumerate(steps, 1)
                 if isinstance(step, Mapping) and "uses" in step]
        for where, ref in uses:
            if not SHA_PINNED.match(str(ref)):
                findings.append(f"{name}: {where}: `{ref}` is not pinned to a 40-character SHA")
        for i, step in enumerate(steps, 1):
            if isinstance(step, Mapping) and str(step.get("uses", "")).startswith("actions/checkout@"):
                if (step.get("with") or {}).get("persist-credentials") is not False:
                    findings.append(f"{name}: job {job_id} step {i}: actions/checkout without "
                                    "persist-credentials: false")
    return findings


def permission_findings(name: str, workflow: Workflow) -> list[str]:
    findings = []
    if "permissions" not in workflow:
        findings.append(f"{name}: no workflow-level permissions")
    elif writes(workflow["permissions"]):
        findings.append(f"{name}: workflow-level write permission "
                        f"{sorted(writes(workflow['permissions']))}")
    for job_id, job in jobs(workflow):
        if "timeout-minutes" not in job:
            findings.append(f"{name}: job {job_id} has no timeout-minutes")
        extra = writes(job.get("permissions")) - ALLOWED_WRITES.get((name, job_id), frozenset())
        if extra:
            findings.append(f"{name}: job {job_id} has write permission {sorted(extra)} "
                            "outside the allowed list")
    return findings


def required_check(governance: str, settings: str) -> tuple[str, str, list[str], list[str]]:
    """The required check's name and workflow file, the protected branches, and findings."""
    findings = []
    match = re.search(r"^\| Required status check \| `([^`]+)`, the job of `([^`]+)`",
                      governance, re.M)
    name, path = (match.group(1), match.group(2)) if match else ("", "")
    if not match:
        findings.append(f"{GOVERNANCE}: no 'Required status check' row naming the job")
    contexts = re.findall(r"^\| Required status context \| `([^`]+)` \| `([^`]+)` \|", settings, re.M)
    names = {context for row in contexts for context in row}
    if not contexts:
        findings.append(f"{SETTINGS}: no 'Required status context' row")
    elif names != {name}:
        findings.append(f"{SETTINGS}: the rulesets require {sorted(names)}, "
                        f"{GOVERNANCE} names `{name}`")
    branches = re.search(r"^Applies to `([^`]+)` and every `([^`]+)` branch", governance, re.M)
    protected = [branches.group(1), branches.group(2)] if branches else []
    if not branches:
        findings.append(f"{GOVERNANCE}: no 'Applies to ... branch' line for the rulesets")
    return name, path, protected, findings


def integrity_findings(workflows: Mapping[str, Workflow], governance: str,
                       settings: str) -> list[str]:
    name, path, protected, findings = required_check(governance, settings)
    workflow = workflows.get(Path(path).name) if path else None
    if workflow is None:
        return findings + [f"required check workflow `{path}` not found"]
    file = Path(path).name
    matching = [(job_id, job) for job_id, job in jobs(workflow) if job.get("name") == name]
    if len(matching) != 1:
        findings.append(f"{file}: {len(matching)} jobs named `{name}`, the required check; "
                        "expected exactly 1")
    for job_id, job in matching:
        for key in ("if", "strategy"):
            if key in job:
                findings.append(f"{file}: required job {job_id} has `{key}`")
    on = triggers(workflow) or {}
    expected = {"branches": protected}
    for event in ("pull_request", "push"):
        if not isinstance(on, Mapping) or on.get(event) != expected:
            actual = on.get(event) if isinstance(on, Mapping) else on
            findings.append(f"{file}: `{event}` filters are {actual!r}; {GOVERNANCE} requires "
                            f"{expected!r} and no other filter")
    if workflow.get("concurrency") != INTEGRITY_CONCURRENCY:
        findings.append(f"{file}: concurrency must be {INTEGRITY_CONCURRENCY}")
    for job_id, job in matching:
        checkouts = [step for step in steps_of(job)
                     if str(step.get("uses", "")).startswith("actions/checkout@")]
        if not checkouts or any((step.get("with") or {}).get("fetch-depth") != 0
                                for step in checkouts):
            findings.append(f"{file}: required job {job_id} must check out with fetch-depth: 0, "
                            "so that host evidence can be bound to its commit")
    return findings


def labels_sync_findings(workflows: Mapping[str, Workflow]) -> list[str]:
    job = dict(jobs(workflows.get("labels-sync.yml", {}))).get("reconcile")
    if job is None:
        return ["labels-sync.yml: no reconcile job"]
    # The whole condition, so that a negation or an appended `|| true` cannot pass.
    findings = []
    if squash(job.get("if", "")) != LABELS_RECONCILE_IF:
        findings.append(f"labels-sync.yml: reconcile's condition must be: {LABELS_RECONCILE_IF}")
    if job.get("permissions") != LABELS_RECONCILE_PERMISSIONS:
        findings.append(f"labels-sync.yml: reconcile's permissions must be {LABELS_RECONCILE_PERMISSIONS}")
    return findings


def codeql_findings(workflows: Mapping[str, Workflow]) -> list[str]:
    """One analysis per language, with the extended queries, on the protected branches."""
    workflow = workflows.get(CODEQL_FILE)
    if workflow is None:
        return [f"{CODEQL_FILE}: not found"]
    findings = []
    on = triggers(workflow)
    branches = {"branches": ["main", "release/**"]}
    if (not isinstance(on, Mapping) or on.get("push") != branches
            or on.get("pull_request") != branches or not on.get("schedule")
            or "workflow_dispatch" not in on):
        findings.append(f"{CODEQL_FILE}: must run on push and pull_request with {branches} "
                        "and no other filter, and on a schedule")
    if workflow.get("concurrency") != CODEQL_CONCURRENCY:
        findings.append(f"{CODEQL_FILE}: concurrency must be {CODEQL_CONCURRENCY}")
    for name, condition in (("analyze", "github.event_name != 'pull_request'"),
                            ("analyze-pr", "github.event_name == 'pull_request'")):
        job = dict(jobs(workflow)).get(name)
        if job is None:
            findings.append(f"{CODEQL_FILE}: no {name} job")
            continue
        permissions = CODEQL_PERMISSIONS if name == "analyze" else {"contents": "read"}
        if job.get("permissions") != permissions:
            findings.append(f"{CODEQL_FILE}: {name}'s permissions must be {permissions}")
        if job.get("if") != condition:
            findings.append(f"{CODEQL_FILE}: {name}'s condition must be {condition}")
        analysis_inputs: dict[str, Any] = dict(CODEQL_ANALYZE)
        if name == "analyze-pr":
            analysis_inputs.update({"upload": "never", "upload-database": False,
                                    "output": "codeql-results"})
            artifacts = [step for step in steps_of(job)
                         if str(step.get("uses", "")).startswith("actions/upload-artifact@")]
            expected_artifact = {
                "name": "codeql-pr-${{ matrix.language }}-${{ github.run_id }}-${{ github.run_attempt }}",
                "path": "codeql-results", "if-no-files-found": "error", "retention-days": 30}
            if len(artifacts) != 1 or artifacts[0].get("with") != expected_artifact:
                findings.append(f"{CODEQL_FILE}: PR SARIF must be retained for 30 days")
        findings += codeql_job_findings(job, analysis_inputs)
    return findings


def codeql_job_findings(job: Workflow, analysis_inputs: Mapping[str, Any]) -> list[str]:
    """Check the same languages and queries on both event paths."""
    findings = []
    matrix = (job.get("strategy") or {}).get("matrix") or {}
    if matrix != {"language": CODEQL_LANGUAGES}:
        findings.append(f"{CODEQL_FILE}: the matrix must be exactly language: {CODEQL_LANGUAGES}")
    actions: dict[str, list[tuple[str, Any]]] = {}
    for step in steps_of(job):
        match = CODEQL_ACTION.match(str(step.get("uses", "")))
        if match:
            actions.setdefault(match.group(1), []).append((match.group(2), step.get("with")))
    expected = {"init": CODEQL_INIT, "analyze": analysis_inputs}
    for action, inputs in expected.items():
        if [given for _, given in actions.get(action, [])] != [inputs]:
            findings.append(f"{CODEQL_FILE}: exactly one {action} step, with {inputs}")
    if len({sha for steps in actions.values() for sha, _ in steps}) > 1:
        findings.append(f"{CODEQL_FILE}: init and analyze must be pinned to the same commit")
    return findings


def links_findings(workflows: Mapping[str, Workflow]) -> list[str]:
    workflow = workflows.get(LINKS_FILE)
    if workflow is None:
        return [f"{LINKS_FILE}: not found"]
    on = triggers(workflow)
    events = set(on) if isinstance(on, Mapping) else {str(on)}
    if events != LINKS_EVENTS:
        return [f"{LINKS_FILE}: must run only on {sorted(LINKS_EVENTS)}, not {sorted(events)}"]
    return []


def triage_findings(workflows: Mapping[str, Workflow]) -> list[str]:
    workflow = workflows.get(TRIAGE_FILE)
    if workflow is None:
        return [f"{TRIAGE_FILE}: not found"]
    on = triggers(workflow)
    if not isinstance(on, Mapping) or on.get("milestone") != TRIAGE_MILESTONE_EVENTS:
        return [f"{TRIAGE_FILE}: milestone events must be {TRIAGE_MILESTONE_EVENTS}"]
    return []


def milestone_script(workflow: Workflow) -> str | None:
    """The inline Python of pr-milestones.yml's only step, or None."""
    reconcile = dict(jobs(workflow)).get("reconcile") or {}
    steps = reconcile.get("steps") or []
    if len(steps) != 1 or not isinstance(steps[0], Mapping):
        return None
    match = SCRIPT.search(str(steps[0].get("run", "")))
    return match.group(1) + "\n" if match else None


def milestone_findings(workflows: Mapping[str, Workflow]) -> list[str]:
    workflow = workflows.get(MILESTONES_FILE)
    if workflow is None:
        return [f"{MILESTONES_FILE}: not found"]
    findings = []
    if triggers(workflow) != MILESTONES_EVENTS:
        findings.append(f"{MILESTONES_FILE}: events changed: {triggers(workflow)!r}")
    if workflow.get("permissions") != {}:
        findings.append(f"{MILESTONES_FILE}: workflow-level permissions must be {{}}")
    if workflow.get("concurrency") != MILESTONES_CONCURRENCY:
        findings.append(f"{MILESTONES_FILE}: concurrency must be {MILESTONES_CONCURRENCY}")
    all_jobs = dict(jobs(workflow))
    if list(all_jobs) != ["reconcile"]:
        findings.append(f"{MILESTONES_FILE}: jobs must be exactly [reconcile], not {list(all_jobs)}")
    job = all_jobs.get("reconcile", {})
    if job.get("permissions") != {"pull-requests": "write"}:
        findings.append(f"{MILESTONES_FILE}: reconcile's permissions must be pull-requests: write")
    if squash(job.get("if", "")) != MILESTONES_IF:
        findings.append(f"{MILESTONES_FILE}: reconcile's condition must be: {MILESTONES_IF}")
    steps = job.get("steps") or []
    step = steps[0] if len(steps) == 1 and isinstance(steps[0], Mapping) else {}
    if not step or "uses" in step or "run" not in step:
        findings.append(f"{MILESTONES_FILE}: reconcile must have a single inline `run` step")
    elif "${{" in str(step["run"]):
        findings.append(f"{MILESTONES_FILE}: the inline code contains `${{{{`")
    elif milestone_script(workflow) is None:
        findings.append(f"{MILESTONES_FILE}: no `<<'PY'` Python script in the step")
    return findings


def strings(value: Any) -> Iterator[str]:
    """Every string in a parsed YAML value, keys excluded."""
    if isinstance(value, str):
        yield value
    elif isinstance(value, Mapping):
        for item in value.values():
            yield from strings(item)
    elif isinstance(value, list):
        for item in value:
            yield from strings(item)


def steps_of(job: Mapping[str, Any]) -> list[Mapping[str, Any]]:
    return [step for step in job.get("steps") or [] if isinstance(step, Mapping)]


def traffic_credential_findings(collect: Mapping[str, Any], name: str) -> list[str]:
    """The App's key reaches only the token step, and the token only the fetch step."""
    findings = []
    steps = steps_of(collect)
    token_steps = [step for step in steps if str(step.get("uses", "")).startswith(TOKEN_ACTION)]
    if len(token_steps) != 1:
        return [f"{name}: collect must create exactly one App token with {TOKEN_ACTION}"]
    token = token_steps[0]
    if token.get("id") != "app-token" or dict(token.get("with") or {}) != TOKEN_INPUTS:
        findings.append(f"{name}: the App token step must be `app-token` with exactly {TOKEN_INPUTS}")
    preflight = [step for step in steps if KEY_PRESENT in strings(step.get("env"))]
    if len(preflight) != 1 or steps.index(preflight[0]) > steps.index(token):
        findings.append(f"{name}: a preflight before the token step must test only whether "
                        "the private key is set")
    secret_uses = sorted(found for step in steps if step is not token
                         for text in strings(step) for found in SECRET.findall(text)
                         if text != KEY_PRESENT)
    if secret_uses:
        findings.append(f"{name}: secrets used outside the App token step: {secret_uses}")
    users = [step for step in steps if any("steps.app-token.outputs.token" in text
                                           for text in strings(step))]
    if len(users) != 1 or dict(users[0].get("env") or {}).get("TRAFFIC_TOKEN") != TRAFFIC_TOKEN_ENV:
        findings.append(f"{name}: only the fetch step may receive the App token, as TRAFFIC_TOKEN")
    return findings


def traffic_alert_findings(collect: Mapping[str, Any], alert: Mapping[str, Any], name: str) -> list[str]:
    """The alert job reads the exact history commit the collect job published."""
    findings = []
    if dict(collect.get("outputs") or {}).get("history-sha") != "${{ steps.publish.outputs.sha }}":
        findings.append(f"{name}: collect must output history-sha from the publish step")
    if "environment" in alert or any(SECRET.search(text) for text in strings(alert)):
        findings.append(f"{name}: the alert job must have no environment and no secret")
    steps = steps_of(alert)
    checkouts = [i for i, step in enumerate(steps)
                 if dict(step.get("with") or {}).get("path") == "history"]
    guards = [i for i, step in enumerate(steps)
              if "^[0-9a-f]{40}$" in str(step.get("run", ""))
              and dict(step.get("env") or {}).get("HISTORY_SHA") == HISTORY_SHA]
    if len(checkouts) != 1 or dict(steps[checkouts[0]].get("with") or {}).get("ref") != HISTORY_SHA:
        findings.append(f"{name}: the alert job must check out the history at {HISTORY_SHA}")
    elif not guards or guards[0] > checkouts[0]:
        findings.append(f"{name}: the alert job must validate the published SHA before its checkout")
    return findings


def traffic_findings(workflows: Mapping[str, Workflow]) -> list[str]:
    workflow = workflows.get(TRAFFIC_FILE)
    if workflow is None:
        return [f"{TRAFFIC_FILE}: not found"]
    all_jobs = dict(jobs(workflow))
    collect, alert = all_jobs.get("collect"), all_jobs.get("alert")
    if collect is None or alert is None:
        return [f"{TRAFFIC_FILE}: jobs collect and alert are required"]
    if collect.get("environment") != "analytics":
        return [f"{TRAFFIC_FILE}: collect must run in the analytics environment"]
    return (traffic_credential_findings(collect, TRAFFIC_FILE)
            + traffic_alert_findings(collect, alert, TRAFFIC_FILE))


def dependabot_findings(config: Any) -> list[str]:
    updates = config.get("updates") if isinstance(config, Mapping) else None
    if not isinstance(updates, list):
        return [f"{DEPENDABOT}: no updates list"]
    findings = []
    count = Counter(str(update.get("package-ecosystem")) for update in updates)
    for ecosystem, directory in PINNED_ECOSYSTEMS.items():
        if count[ecosystem] != 1:
            findings.append(f"{DEPENDABOT}: `{ecosystem}` is covered {count[ecosystem]} times; "
                            "expected once")
        for update in updates:
            if update.get("package-ecosystem") == ecosystem and update.get("directory") != directory:
                findings.append(f"{DEPENDABOT}: `{ecosystem}` must cover `{directory}`")
    for update in updates:
        ecosystem = update.get("package-ecosystem")
        if "milestone" in update:
            findings.append(f"{DEPENDABOT}: `{ecosystem}` sets a milestone")
        priorities = sorted(str(label) for label in update.get("labels") or []
                            if PRIORITY.match(str(label)))
        if priorities != ["P3"]:
            findings.append(f"{DEPENDABOT}: `{ecosystem}` has priorities {priorities}; "
                            "expected exactly ['P3']")
    return findings


def documented_versions(tooling: str) -> dict[str, str]:
    """The Tool versions rows of tools/README.md: tool -> version."""
    return dict(re.findall(r"^\| (Python|Node\.js) \| ([^ |]+) \|", tooling, re.M))


def runtime_findings(workflows: Mapping[str, Workflow], tooling: str) -> list[str]:
    """Every runtime is an exact patch version, the one tools/README.md records."""
    documented = documented_versions(tooling)
    findings = []
    for tool in RUNTIME_INPUTS.values():
        if not EXACT_VERSION.match(documented.get(tool, "")):
            findings.append(f"{TOOLING}: the {tool} row must give an exact X.Y.Z version")
    for name, workflow in sorted(workflows.items()):
        for job_id, job in jobs(workflow):
            for step in steps_of(job):
                for key, tool in RUNTIME_INPUTS.items():
                    value = dict(step.get("with") or {}).get(key)
                    if value is not None and str(value) != documented.get(tool):
                        findings.append(f"{name}: job {job_id}: {key} `{value}` must be the "
                                        f"exact version {TOOLING} records for {tool}")
    return findings


def findings_for(workflows: Mapping[str, Workflow], dependabot: Any, governance: str,
                 settings: str, tooling: str) -> list[str]:
    findings = []
    for name, workflow in sorted(workflows.items()):
        findings += action_findings(name, workflow) + permission_findings(name, workflow)
    findings += integrity_findings(workflows, governance, settings)
    findings += labels_sync_findings(workflows) + milestone_findings(workflows)
    findings += triage_findings(workflows) + codeql_findings(workflows)
    findings += links_findings(workflows)
    findings += traffic_findings(workflows)
    findings += runtime_findings(workflows, tooling)
    findings += dependabot_findings(dependabot)
    return findings


def workflow_files(root: Path) -> list[Path]:
    """Every workflow file: GitHub runs both the .yml and the .yaml extension."""
    folder = root / WORKFLOWS
    return sorted([*folder.glob("*.yml"), *folder.glob("*.yaml")])


def load(root: Path, yaml: Any) -> tuple[dict[str, Workflow], Any, str, str, str]:
    loader = strict_loader(yaml)
    workflows = {path.name: yaml.load(path.read_text(encoding="utf-8"), Loader=loader)
                 for path in workflow_files(root)}
    dependabot = yaml.load((root / DEPENDABOT).read_text(encoding="utf-8"), Loader=loader)
    return (workflows, dependabot, (root / GOVERNANCE).read_text(encoding="utf-8"),
            (root / SETTINGS).read_text(encoding="utf-8"),
            (root / TOOLING).read_text(encoding="utf-8"))


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    args = parser.parse_args(argv)
    yaml = load_yaml()
    if yaml is None:
        if os.environ.get("GITHUB_ACTIONS"):
            print("workflow-invariants: FAIL: PyYAML is not installed "
                  "(tools/requirements-quality-ci.txt)")
            return 1
        print("workflow-invariants: SKIPPED locally: PyYAML is not installed")
        return 2
    try:
        workflows, dependabot, governance, settings, tooling = load(args.root, yaml)
    except yaml.YAMLError as error:
        print(f"workflow-invariants: FAIL: {error}")
        return 1
    findings = findings_for(workflows, dependabot, governance, settings, tooling)
    for finding in findings:
        print(finding)
    print(f"workflow-invariants: {'FAIL' if findings else 'PASS'} "
          f"({len(workflows)} workflows, {len(findings)} findings)")
    return 1 if findings else 0


if __name__ == "__main__":
    sys.exit(main())
