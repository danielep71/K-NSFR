<a id="repository-traffic-history"></a>

# 📈 Repository traffic history

**K-NSFR** · [Overview](../README.md) · [Documentation](../README.md#documentation) · [Governance](GOVERNANCE.md)

[![Stage](https://img.shields.io/badge/Stage-Repository_setup-BF8700?style=flat-square)](../README.md#status)

GitHub keeps repository traffic for 14 days. The **Traffic history** workflow
([.github/workflows/traffic-history.yml](../.github/workflows/traffic-history.yml))
collects it once a day, with the repository's stars, forks, watchers and open
issues, and merges it into CSV files on the data-only `traffic-history`
branch, so the history outlives that window. It also writes badge files
there, and opens an issue when traffic stands out (*Traffic alert* below). It
is repository analytics, not CI: it builds and tests nothing.

The workflow is on `main` since the integration of 2026-11-03 (#107). The
schedule and manual dispatch described below also require the credential
setup; until it exists the collect job fails at its *Check the App
configuration* step. No live collection, push or alert is verified;
operational validation remains tracked in #16.

<a id="schedule"></a>

## ⏱️ Schedule

The workflow is scheduled daily at **06:17 UTC** and can be started from
**Actions → Traffic history → Run workflow** on `main`. GitHub may start
scheduled runs late; each collection records the time it was actually taken.
A manual run on any other branch is skipped.

<a id="credential-setup-and-renewal"></a>

## 🔑 Credential setup and renewal

The traffic API needs **Administration: read** on the repository, which the
workflow token cannot have. The collect job therefore creates a short-lived
installation token of a private GitHub App that I own, installed on this
repository only and allowed only *Administration: read* (#143). The token
expires within an hour and is revoked when the job ends; no personal access
token is used.

1. **Settings → Environments → `analytics`.** Under *Deployment branches and
   tags*, select *Selected branches and tags* and allow `main` only. A run that
   names a missing environment creates it without that restriction, so check
   it before adding the credentials.
2. **Settings → Developer settings → GitHub Apps → New GitHub App** under my
   account:
   - a name and a homepage URL;
   - *Webhook*: *Active* unchecked;
   - *Repository permissions*: **Administration: Read-only**, and nothing else
     (*Metadata: Read-only* is added by GitHub);
   - *Where can this GitHub App be installed?*: *Only on this account*.
3. In the App's settings, **Generate a private key**, which downloads a
   `.pem` file, and **Install App** on my account with *Only select
   repositories*: `K-NSFR`.
4. In the `analytics` environment:
   - add the **variable** `TRAFFIC_APP_CLIENT_ID` with the App's *Client ID*;
     it is not secret;
   - add the **secret** `TRAFFIC_APP_PRIVATE_KEY` with the whole content of
     the `.pem` file, then delete the file.

   Never put the key in a file of the repository, an issue, a comment or chat.
5. Run the workflow manually on `main` and check that it passes and that
   `traffic-history` gains a commit.

**Rotation:** generate a new private key in the App's settings, replace the
secret, run step 5, then delete the old key in the App's settings. Delete a key
at once if it may have been exposed, or if the workflow is changed by anyone
but me. Collections missed while the credential is invalid are lost once they
fall outside the 14-day window.

The collect job keeps the key and the token apart from everything else:

- *Check the App configuration* reads the client ID and only whether the key
  is set (a boolean), never the key itself. A missing value fails here, naming
  the setting.
- *Create the App installation token* (`actions/create-github-app-token`,
  pinned by commit) is the only step that receives the key. It asks for a
  token on this repository with *Administration: read* only.
- *Fetch traffic from the GitHub API* is the only step that receives the token,
  as `TRAFFIC_TOKEN`, and it runs `curl` only. The merge script runs without
  it, and the push uses the workflow token with `contents: write`.

The workflow-invariant checks enforce these rules (`tools/check_workflows.py`).

<a id="stored-data"></a>

## 🗄️ Stored data

The `traffic-history` branch holds `README.md` and these files under `data/`.
It was created on 2026-10-29 as commit `7d19eb6`, on top of `main` at that
date, with only its README; the collector creates the files below on its first
run. It is never merged into `main` or a release branch, and its history is
not rewritten.

| File | Key | Contents |
| --- | --- | --- |
| `daily.csv` | `date` | Views and unique views, clones and unique clones per UTC day. `views_collected_at` and `clones_collected_at` give the collection that last changed each pair; `views_observed_at` and `clones_observed_at` give the latest collection that observed it, changed or not. |
| `snapshots.csv` | `snapshot_date` | One row per collection day: the first and last day the API reported, the 14-day totals, the repository's `stars`, `forks`, `watchers` and `open_issues` at collection time, `collected_at`, the collection that last changed that day's snapshot, referrers or paths, and `observed_at`, the latest collection that observed them. |
| `referrers.csv` | `snapshot_date`, `rank` | The top 10 referrers over the 14 days before each collection. |
| `paths.csv` | `snapshot_date`, `rank` | The top 10 repository paths over the same window. |
| `badges/*.json` | — | Badge files rewritten from the stored data after each merge (*Badges* below). |

Referrers and paths are retained per collection day, because the API gives
them only as 14-day totals, never per day. The repository counts come from
the repository API (`stargazers_count`, `forks_count`, `subscribers_count`,
`open_issues_count`); `open_issues` is GitHub's count, which includes open
pull requests.

<a id="badges"></a>

### Badges

The collector writes three shields.io *endpoint* files to `data/badges/`:
`clones` (every clone in `daily.csv`, summed since tracking began),
`visitors (14d)` (unique visitors over the API's 14-day window, from the latest
snapshot) and `visits` (daily unique visitors summed since tracking began).
The README does not display them: [documentation style](WRITING_STYLE.md#badge-and-glyph-rules)
admits a traffic badge only with the underlying data and a review. A badge
file changes only when its value does.

<a id="merge-rules"></a>

### Merge rules

`tools/traffic_history.py` merges each collection idempotently:

- **Overlapping windows:** for every date the collection reports, its values
  replace the stored ones, so a day that was incomplete when first collected
  is corrected later. Views and clones are merged separately, each with its
  own collection time, which stays unchanged when its values do not change,
  and its own observed-at watermark, which advances on every newer
  observation, even with identical values.
- **Older history:** dates the collection does not report keep their values.
  Nothing is deleted.
- **No double counting:** values are replaced, never added. Distinct visitors
  across days or windows cannot be derived by summing daily
  unique counts; the API's 14-day unique count is stored in `snapshots.csv`.
  The `visits` badge does sum daily unique visitors, so it counts daily visits,
  not distinct visitors over the whole stored history.
- **Same-day runs:** a collection day's snapshot, referrers and paths are one
  unit, replaced only if any of the three changed; an unchanged later
  collection advances only `observed_at`. Replaying a collection with its own
  collection time changes nothing.
- **Order:** "newest" means the latest collection time. Freshness is judged
  by the observed-at watermark, not by the time of the last change, so merging
  an older collection after a newer one never replaces what the newer one
  observed, even when the newer one found the same values or counts were
  later revised. It can only fill a date or metric that nothing newer
  reported. Rows written before the watermarks existed have them blank and
  fall back to the time of the last change.

<a id="zero-or-unavailable"></a>

### Zero or unavailable

The API reports each day of its window, days without activity included, as
`0`. So **`0`** means the day was collected and had no activity; **no row**
means no collection covered the day; **a blank metric** means one of the two
series was not reported for that day. The step summary of each run lists the
dates between the first and the last stored day that have no row.

<a id="visibility"></a>

### Visibility

The repository has been public since 2026-10-28, and so is the branch: anyone
can read its CSV files. The data is aggregate (counts, referrers and paths)
and identifies no visitor. The counts include my own visits and every clone
made by CI checkouts.

<a id="traffic-alert"></a>

## 🚨 Traffic alert

After each collection the **Traffic alert** job compares the latest snapshot
with the one before (`traffic_history.py alert`) and opens one issue, titled
`Traffic alert: <snapshot date>`, when any of these holds:

- **Views spike:** the 14-day total at least doubles and grows by at least
  100 views;
- **Stars or forks increase;**
- **New referrer:** a referrer never stored before brings at least 5 views.

Clones raise no alert: they include every CI checkout, so a rise says nothing
about readers. They stay in the history and the `clones` badge. A count blank
in the older snapshot raises nothing. At most one issue is
opened per snapshot date: the job skips a date whose issue exists, open or
closed. The issue is assigned to me, labelled `P3` and `repository`, and
filed under the open milestone with the earliest due date after today (UTC).
Its *Planning* section has the snapshot date as its start date and a target
end date a week later, or the day before the milestone's due date if that is
earlier, so it passes [Issue triage](GOVERNANCE.md#issues). A milestone due
on or before the snapshot date would leave no valid target end date, so the
script refuses it and the step fails.

The job picks the assignee and the milestone by its own rules:
`github.repository_owner`, which is my account, and the open milestone with
the earliest due date after today. Triage instead expects the assignee configured in the
Bug report form and the milestone of the active release. Both give the same
result while the active release's milestone has the earliest due date, so give
any later milestone a later due date.

The job uses the workflow token with `issues: write` and `contents: read`. It
has no environment, so it never receives the App's key or token, and it runs
only after a successful collection. It reads exactly the `traffic-history`
commit that collection published: the collect job outputs that commit's SHA,
and the alert job checks that it is a 40-character hex SHA before checking it
out, so another writer moving the branch in between cannot change what the
alert reads.

<a id="branch-protection"></a>

## 🛡️ Branch protection

The workflow pushes to `traffic-history` directly, so its rules differ from
those of `main` and `release/**`
([installed settings](GOVERNANCE.md#installed-settings-and-limits)):

| Rule | `traffic-history` |
| --- | --- |
| Deletion | Blocked. |
| Force pushes | Blocked. |
| Pull request required | No: the workflow commits directly. |
| Required status checks | None. |
| Bypass | None. |

<a id="failure-diagnosis-and-recovery"></a>

## 🩺 Failure diagnosis and recovery

Each run fails before anything is written if a response is missing or
invalid, and leaves the history unchanged. Read the failing step:

| Step and message | Cause | Action |
| --- | --- | --- |
| *Check the App configuration*: `Not set in the analytics environment: ...` | The named variable or secret is missing, or the run is not in the `analytics` environment. | Follow the credential setup. |
| *Create the App installation token* fails | The client ID or key is wrong, the key was deleted, or the App is not installed on this repository. | Check the App's installation and key, then replace the variable or secret. |
| *Fetch*: HTTP 401 | The token was rejected. | Run again; if it persists, rotate the key. |
| *Fetch*: HTTP 403 or 404 | The App lacks *Administration: read*, or its installation does not include this repository. | Fix the App's permissions or installation; a permission change must be accepted on the installation. |
| *Fetch*: HTTP 429 or 5xx after retries | Rate limit or GitHub outage. | Run it again later; the next run collects the same 14 days. |
| *Merge*: `ERROR: ...; the history was not changed` | The API returned an unexpected shape, or a history file no longer has the expected layout. | Inspect the response or the named file; a file edited by hand must keep its header and columns. |
| *Commit and push*: `Could not push the history` | `traffic-history` kept moving, or a rule rejected the push. | Check the branch's rules and recent commits, then run again. |
| *Traffic alert*: `The collect job published no traffic-history commit SHA` | The collect job ended without publishing. | Read the collect job's log; run the workflow again. |
| *Check out traffic-history* fails | The branch was deleted. | Restore it from the last known commit; never recreate it empty over existing history. |
| *Traffic alert*: `No open milestone due after …` | Every milestone is closed, has no due date, or is already due. | Open the next milestone with a later due date, then run the workflow again. |
| *Traffic alert*: `gh issue create` fails | A label, the assignee or the milestone no longer exists. | Restore it, then run the workflow again; the history is already saved. |

A failed run can always be repeated: the merge is idempotent, and each
collection covers the last 14 days. If the branch must be repaired, change it
with a new commit, keeping every header. The script lives on `main` and the
data on `traffic-history`, so work from a clone of `main` with the data branch
in a second worktree:

```sh
git worktree add ../history traffic-history
python tools/traffic_history.py init --data ../history/data
python tools/traffic_history.py merge --raw <responses> --data ../history/data --collected-at <UTC time>
```

`init` recreates missing CSV files with their headers and never overwrites an
existing one; the badge files are rewritten by the next `merge`. `merge` takes
a directory of saved API responses, as the workflow's fetch step writes them.
Commit the change in `../history` and push it to `traffic-history`.
