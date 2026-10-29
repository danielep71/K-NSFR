# traffic-history

Data-only branch of K-NSFR, reserved for repository traffic history.

Do not merge this branch into development branches, base development work on it,
or rewrite its history. The existing protection rules block deletion and force
pushes while allowing ordinary data updates.

## Collection status

Initialized on 2026-10-29. No traffic observations have been collected yet.
The default branch does not yet contain a traffic-collection workflow.
Missing observations must not be represented as zero traffic.

## Intended layout

| File | Contents |
| --- | --- |
| `data/daily.csv` | Views and clones per UTC day |
| `data/snapshots.csv` | One row per collection day, with the 14-day totals |
| `data/referrers.csv` | Top referrers per collection day |
| `data/paths.csv` | Popular paths per collection day |

Create these files through the collector when actual observations are available.
