<a id="documentation-style"></a>

# 🎨 Documentation style

**K-NSFR** · [Overview](../README.md) · [Documentation](../README.md#documentation) · [Governance](GOVERNANCE.md)

[![Stage](https://img.shields.io/badge/Stage-Repository_setup-BF8700?style=flat-square)](../README.md#status)

This is the presentation pattern for future K-NSFR Markdown. Keep the subject
matter precise: a polished document must never imply an implemented engine,
validated profile, Excel run, public release or production suitability before
the linked evidence exists.

<a id="page-structure"></a>

## 🧭 Page structure

- Start with one descriptive `#` heading and a relevant glyph. Keep the
  wording explicit so links and search remain useful.
- On pages under `docs/`, follow the title with the overview/documentation/
  governance navigation line and the setup-stage badge used on this page.
  Root policy pages use root-relative counterparts.
- Use a short introduction, descriptive `##` sections with distinct glyphs,
  and tables only for genuine comparisons or mappings.
- Keep warnings in GitHub callouts when a reader might mistake a target for
  delivered capability. Link to evidence or the owning issue for limitations.
- End related explanations with actionable local links. Require an explicit `<a id="stable-name"></a>` anchor for each linked
  section. Keep the ID unchanged when renaming or decorating headings.

<a id="badge-and-glyph-rules"></a>

## 🏷️ Badge and glyph rules

The README hero uses static, legible shields for **stage**, **declared host**
and **license**. Other pages use the compact stage badge. Use the same
vocabulary and colors across the repo: amber `BF8700` for setup, Excel green
`217346` for a declared VBA target, blue `0078D6` for Windows, and green
`2ea44f` for the selected MPL-2.0 license. A green validation badge requires
an actual current result, linked to the candidate and test evidence. Do not add
traffic, downloads, host-pass or regulatory-validation badges without the
underlying data and a current review.

Use glyphs as scanning aids: 🏦 project, 🎯 status/scope, 🧭 navigation/
decisions, 🏗️ architecture, 🧰 tools, ✅ checks, 🗺️ roadmap, ⚖️ governance/
license, 🔒 security, 🧪 synthetic data and 🔗 provenance. The text of each
heading carries the meaning even if a glyph does not render.

<a id="maintenance"></a>

## ✅ Maintenance

When changing a Markdown file, verify that local links resolve, fenced code
blocks close, the page renders with readable headings and tables, and the
repository's required integrity check passes. Update a badge only after the
fact it represents changes; never copy results from another project. The
source of truth for maturity is [scope](SCOPE.md), and validation boundaries
are in [architecture](ARCHITECTURE.md#validation-boundaries).
