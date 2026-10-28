<a id="source-control-and-generated-file-conventions"></a>

# 🧰 Source-control and generated-file conventions

**K-NSFR** · [Overview](../README.md) · [Documentation](../README.md#documentation) · [Governance](GOVERNANCE.md)

[![Stage](https://img.shields.io/badge/Stage-Repository_setup-BF8700?style=flat-square)](../README.md#status)

Repository source is canonical. Cross-platform text uses UTF-8 without a BOM,
LF in committed text, and a final newline. Remove trailing spaces or tabs except
for intentional Markdown hard breaks. Preserve the standard license text.
VBA and Windows-native source retain their host encoding contract: the editor
configuration does not force a charset, and no Git working-tree-encoding rule is
active. [Implementation conventions](IMPLEMENTATION.md) define the initial source
repertoire and host boundary; #13/#14 must record the exact code page and prove
source import/export on the supported host. ASCII remains a
portable subset for test examples, not an imposed source-encoding conversion.

`.editorconfig` defines editor defaults. `.gitattributes` normalizes repository
text to LF and checks VBA and Windows-native scripts out as CRLF for the intended
host. Cross-platform text checks out as LF regardless of core.autocrlf. No custom clean/smudge filter or
working-tree-encoding conversion is required. Use Git checkout rather than
assuming a downloaded source archive has the host line endings. Git conversion
tests are portable evidence; actual Excel import/export remains #13/#14 work.

The current layout is documented in [ARCHITECTURE.md](ARCHITECTURE.md). This
change adds only configuration, a committed-content check and its tests; future
source, fixtures and templates directories are created with real deliverables.

<a id="generated-files-and-intentional-binaries"></a>

## 📦 Generated files and intentional binaries

`.gitignore` excludes the named root build/release staging and local test-output
directories, caches, virtual environments, local editor files, credentials,
temporary Excel files and generated workbooks. Keep retained evidence in a
reviewed, trackable path; do not ignore documentation or evidence directories
wholesale. Exported VBA and .frx companions remain trackable. Ignore rules are
not secret scanning and do not protect
already tracked files or forced additions. Review staged changes before commits.

No binary is approved at this baseline: `tools/binary-allowlist.json` is empty.
For an intentional binary, obtain review of the exact path, purpose, source,
rights and data contents; add that path to the allowlist, add an exact negation
in .gitignore if needed, and ensure binary attributes prevent text conversion.
Never allow a template directory or extension wholesale. Future templates must
follow #13's manifest and data-policy review. Do not use git add -f as a substitute
for an approved exception. Existing license and attribution text remains intact.

<a id="validation"></a>

## ✅ Validation

After committing the candidate, run:

```sh
python tools/check_source_control.py --ref HEAD
python -m unittest discover -s tools -p 'test_source_control.py'
```

The checker reads blob objects and the binary allowlist from the specified
commit, not working-tree bytes. It therefore does not mistake intentional CRLF
checkout conversion for a committed whitespace defect. It rejects unexpected
binary formats according to committed attributes resolved in a disposable
repository with an empty index, no templates and isolated Git configuration.
Local/global/system attributes, environment-injected Git config and dirty or
staged attribute files do not affect that resolution. It also rejects, invalid cross-platform UTF-8,
BOM/NUL, unintended trailing whitespace, missing newlines, symlinks and submodules.
Host-text inspection is byte-preserving and does not certify a code page.
Run it after committing; uncommitted
changes are deliberately outside its scope. A failing scan reports paths and
categories without printing file contents.

The regression suite creates temporary repositories, fresh clones with
core.autocrlf=false/true/input, checks VBA CRLF and documentation LF checkouts,
and stages renormalized content to verify no blob changes. It also tests dirty
working trees, bad committed text, host encoding preservation, Markdown line
breaks, ignore rules, and exact binary exceptions. These fixtures
exist only in temporary test directories and are not Excel execution evidence.

For existing clones, start with a clean worktree. After pulling the attributes,
use a fresh clone for host work if existing file endings are uncertain. Review
any deliberate `git add --renormalize .` result before committing it; do not mix
unrelated edits into normalization. `git ls-files --eol` shows index/worktree
endings for troubleshooting. Required Repository integrity runs the committed
content check and regression suite on candidate PRs.

The attributes also classify binary merges, VBA diffs, language reporting and
archive exclusions. A source archive omits repository control files and GitHub
automation by design; use a Git clone for the repository validation commands.
No workbook exception is enabled until an actual K-NSFR template is approved.
