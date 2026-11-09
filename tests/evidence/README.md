<a id="host-evidence-records"></a>

# 🖥️ Host evidence records

**K-NSFR** · [Overview](../../README.md) · [Documentation](../../README.md#documentation) · [Host evidence](../../docs/validation/HOST_EVIDENCE.md)

[![Stage](https://img.shields.io/badge/Stage-Repository_setup-BF8700?style=flat-square)](../../README.md#status)

This folder holds one JSON record per tested commit, named
`<candidate SHA>.json`. [Host evidence](../../docs/validation/HOST_EVIDENCE.md)
defines the record and how to make one; the `host-evidence` gate validates
every record here against the Git history.

| Record | Host | Result |
| --- | --- | --- |
| [`5231985`](523198552eee8296b1dfa4e72344c2a8e5d55cfd.json) | Windows 11 Enterprise 25H2, Microsoft 365 Apps build 20430.20180, 64-bit; recorded 2026-11-09 | PASS (#12, #14) |
