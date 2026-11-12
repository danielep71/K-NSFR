<a id="regulatory-source-register"></a>

# 📚 Regulatory source register

**K-NSFR** · [Overview](../../README.md) · [Documentation](../../README.md#documentation) · [Governance](../GOVERNANCE.md)

[![Stage](https://img.shields.io/badge/Stage-Repository_setup-BF8700?style=flat-square)](../../README.md#status)

This register records the sources of the planned `eu-crr-standard` profile
([scope](../SCOPE.md#regulatory-profiles)): which texts are authoritative,
which only explain, which provisions each future requirement will trace to,
and what I decided about applicability. Tracking: #8.

It contains no formula, threshold, factor, percentage or parameter value.
Values are captured later, with citations, in a reviewed parameter set
([PROD-01](../BACKLOG.md#later-calculation-work)). No source listed here is
implemented, and no executable profile exists.

<a id="reading-the-register"></a>

## 🧭 Reading the register

Register version: **1**, compiled on **2026-11-05**. Every entry states how it
was verified:

- **Verified (EUR-Lex)**: checked by me in the official text on EUR-Lex on
  2026-11-05.
- **Verified (EBA)**: read on 2026-11-05 in the EBA Interactive Single
  Rulebook, which reproduces the CRR structure and links each article to its
  EUR-Lex consolidated source.
- **Verified (BIS)** / **Verified (Bundesbank)**: read on 2026-11-05 on that
  authority's official site.
- **To verify on EUR-Lex**: not yet checked in the official text. Nothing
  relies on such a value until it is checked.

A source is **normative** when it sets binding rules for the profile,
**reporting** when it binds only the supervisory reporting format, and
**explanatory** when it helps interpret but never overrides the normative
text.

<a id="profile-version-and-scope"></a>

## 📌 Profile version and scope

My decisions of 2026-11-05 ([ADR-009](../DECISIONS.md#recorded-decisions)):

- **Legal version.** Version 1 of `eu-crr-standard` is bound to the CRR
  consolidated version `02013R0575-20260626`, the version the EBA rulebook
  links on 2026-11-05, together with any later amending act verified in this
  register. A new register version revisits the date.
- **Entity scope.** A single institution on an individual basis. Application
  on a consolidated basis (Article 428a), the treatments within a group,
  institutional protection scheme or cooperative network (Articles 428g and
  428h), and the simplified calculation for small and non-complex
  institutions (Part Six Title IV Chapters 5 to 7) are outside version 1.
- **Basel.** The Basel standard is reference-only (`basel-reference`): no
  executable mapping or equivalence with the EU profile is claimed.

<a id="normative-sources"></a>

## ⚖️ Normative sources

| ID | Instrument and issuer | Version and dates | Applicable parts | Official URL | Verification |
| --- | --- | --- | --- | --- | --- |
| S-01 | Regulation (EU) No 575/2013 of the European Parliament and of the Council of 26 June 2013 on prudential requirements for credit institutions and investment firms (CRR) | Consolidated version `02013R0575-20260626`. Original: OJ L 176, 27.6.2013, p. 1 | Part Six Title I (Article 413), Title III (Articles 427–428), Title IV (Articles 428a–428az); Part Ten Title II (Article 510) | [ELI](http://data.europa.eu/eli/reg/2013/575/oj) · [consolidated version](https://eur-lex.europa.eu/legal-content/EN/TXT/?uri=CELEX:02013R0575-20260626) | OJ reference: Verified (EUR-Lex). Version identifier and article structure: Verified (EBA) |
| S-02 | Regulation (EU) 2019/876 of the European Parliament and of the Council of 20 May 2019 amending Regulation (EU) No 575/2013 as regards, among others, the net stable funding ratio (CRR2) | OJ L 150, 7.6.2019, p. 1. Entry into force 27 June 2019; application from 28 June 2021 under Article 3(2), subject to the exceptions of that Article | Inserted Part Six Title IV, which applies from 28 June 2021; its text is read through the S-01 consolidated version | [ELI](http://data.europa.eu/eli/reg/2019/876/oj) | Verified (EUR-Lex) |
| S-03 | Regulation (EU) 2025/1215 of 17 June 2025 amending Regulation (EU) No 575/2013 as regards the stable funding treatment of short-term transactions with financial customers | OJ L, 2025/1215, 25.6.2025. Entry into force 26 June 2025; application from 29 June 2025 | Article 510: makes permanent the required stable funding treatment of the relevant short-term transactions with financial customers, securities financing and unsecured transactions included; deletes Article 510(7) and (8), so the increases they scheduled do not apply; updates the EBA monitoring in Article 510(6). The rest of Article 510 is not made permanent by it | [ELI](http://data.europa.eu/eli/reg/2025/1215/oj) | Verified (EUR-Lex) |
| S-04 | Regulation (EU) 2024/1623 of the European Parliament and of the Council of 31 May 2024 amending Regulation (EU) No 575/2013 (CRR3) | OJ L, 2024/1623, 19.6.2024 | Article 1 point 208 amends Article 428(1)(k) (Title III, items requiring stable funding): the reference to undrawn credit facilities becomes a reference to buckets 4, 3 or 2 of Annex I. No direct amendment to Title IV (Articles 428a and following) or to Article 510; indirect effects of its other amendments to the CRR are not excluded | [ELI](http://data.europa.eu/eli/reg/2024/1623/oj) | Verified (EUR-Lex) |

<a id="reporting-sources"></a>

## 📋 Reporting sources

They define the supervisory templates, not the calculation. Version 1 uses
them only to name and order outputs; they never override S-01.

| ID | Instrument and issuer | Version and dates | Applicable parts | Official URL | Verification |
| --- | --- | --- | --- | --- | --- |
| R-01 | Commission Implementing Regulation (EU) 2021/451 (implementing technical standards on supervisory reporting) | Ceased to apply generally from 1 January 2025, with specific transitional exceptions for market risk; repealed with effect from 31 December 2026 (see R-02 and R-03) | NSFR templates C 80.00 (required stable funding), C 81.00 (available stable funding), C 82.00 and C 83.00 (simplified calculation) and C 84.00 (summary) | [ELI](http://data.europa.eu/eli/reg_impl/2021/451/oj) | Template codes: Verified (Bundesbank, [NSFR reporting page](https://www.bundesbank.de/en/tasks/banking-supervision/individual-aspects/liquidity/supervisory-reporting-of-institutions-with-regard-to-the-net-stable-funding-ratio)). Application and repeal dates: Verified (EUR-Lex) |
| R-02 | Commission Implementing Regulation (EU) 2024/3117 of 29 November 2024 on supervisory reporting, repealing Implementing Regulation (EU) 2021/451 | As amended by R-03. Its Article 25 originally repealed R-01 from 31 December 2025 | Article 17 and Annex I, section 7 *Reporting on stable funding*: C 80 and C 81 for the standard methodology, C 82 and C 83 for the simplified methodology for eligible institutions, C 84 for all institutions concerned. Article 22 has the EBA publish the templates and instructions; the annex numbering of R-01 does not carry over | [ELI](http://data.europa.eu/eli/reg_impl/2024/3117/oj) | Verified (EUR-Lex) |
| R-03 | Commission Implementing Regulation (EU) 2025/2475 amending Implementing Regulation (EU) 2024/3117 | As adopted | Article 1 point 3 replaces Article 25(1) and (2) of R-02, postponing the repeal of R-01 to 31 December 2026 | [ELI](http://data.europa.eu/eli/reg_impl/2025/2475/oj) | Verified (EUR-Lex) |

<a id="explanatory-sources"></a>

## 🔎 Explanatory sources

| ID | Source and issuer | Version and dates | Use | Official URL | Verification |
| --- | --- | --- | --- | --- | --- |
| E-01 | EBA Interactive Single Rulebook, Capital Requirements Regulation | Links the CRR consolidated version `02013R0575-20260626` on 2026-11-05 | Navigation of the CRR structure and of related EBA Q&As; never a substitute for the official text | [CRR tree](https://www.eba.europa.eu/regulation-and-policy/single-rulebook/interactive-single-rulebook/12674) | Verified (EBA) |
| E-02 | EBA Single Rulebook Q&A on Part Six Title IV | Each answer is dated individually | Interpretation of a provision, cited by Q&A ID when used | [CRR tree](https://www.eba.europa.eu/regulation-and-policy/single-rulebook/interactive-single-rulebook/12674) (Q&A links per article) | Verified (EBA) for the existence of Q&A links; no answer is adopted yet |
| E-03 | Basel Committee on Banking Supervision, *Basel III: the net stable funding ratio* (BCBS d295) | Published 31 October 2014; integrated into the consolidated Basel Framework | Reference-only background for `basel-reference` | [BIS](https://www.bis.org/bcbs/publ/d295.htm) | Verified (BIS) |
| E-04 | Basel Framework, chapter NSF *Net stable funding ratio*: NSF10 Definitions and applications, NSF20 Calculation and reporting, NSF30 Available and required stable funding, NSF99 Application guidance | Effective as of 15 Dec 2019; status current | Reference-only background for `basel-reference` | [BIS](https://www.bis.org/committees/bcbs/basel-framework/standard/nsf?allChapters=true) | Verified (BIS) |

<a id="structure-of-the-normative-text"></a>

## 🗂️ Structure of the normative text

CRR Part Six "Liquidity", as the EBA rulebook shows it for `02013R0575-20260626`
(Verified (EBA)):

| Provision | Heading | Articles |
| --- | --- | --- |
| Title I | Definitions and liquidity requirements | 411–414, including Article 413 *Stable funding requirement* |
| Title III | Reporting on stable funding | 427–428 |
| Title IV Chapter 1 | The net stable funding ratio | 428a *Application on a consolidated basis*, 428b *The net stable funding ratio* |
| Title IV Chapter 2 | General rules for the calculation of the net stable funding ratio | 428c–428h |
| Title IV Chapter 3 | Available stable funding: Section 1 *General provisions* and Section 2 *Available stable funding factors* | 428i–428j and 428k–428o |
| Title IV Chapter 4 | Required stable funding: Section 1 and Section 2 | 428p–428q and 428r–428ah |
| Title IV Chapter 5 | Derogation for small and non-complex institutions | 428ai |
| Title IV Chapter 6 | Available stable funding for the simplified calculation of the net stable funding ratio | 428aj–428ap |
| Title IV Chapter 7 | Required stable funding for the simplified calculation of the net stable funding ratio | 428aq–428az |
| Part Ten Title II | Article 510 *Net Stable Funding Requirements* (transitional provisions) | 510 |

<a id="concept-mapping"></a>

## 🧩 Concept mapping

Mapping version 1. Each row links a regulatory concept to the provision it
traces to, the planned requirement, the parameter set that will hold its
values, and the issue that will test it. "Out of v1" rows follow the scope
decision above.

| Concept | Normative locator | Planned requirement | Parameter set | Tests |
| --- | --- | --- | --- | --- |
| Stable funding requirement and the ratio | S-01 Articles 413, 428b | Result fields for available stable funding, required stable funding and the ratio, with diagnostics (#9) | None | Ratio and zero-denominator cases (#10) |
| General calculation rules | S-01 Article 428c | Input validation and gross-basis treatment rules in the contracts (#9) | None | Invalid-input cases (#10) |
| Derivative contracts | S-01 Article 428d | Out of contract v1: rejected with `NSFR-E012` (#9); admission needs its own issue | — | — |
| Netting of secured lending and capital market-driven transactions | S-01 Article 428e | Out of contract v1: rejected with `NSFR-E012` (#9); admission needs its own issue | — | — |
| Interdependent assets and liabilities | S-01 Article 428f | Out of v1: the treatment depends on a supervisory decision under that Article (to verify on EUR-Lex) | — | — |
| Institutional protection schemes, cooperative networks, group treatment | S-01 Articles 428g, 428h | Out of v1 (entity scope) | — | — |
| Available stable funding and residual maturity of liabilities | S-01 Articles 428i, 428j | Liability classification and residual-maturity inputs (#9) | ASF factors, cited per Articles 428k–428o | ASF classification and maturity boundary cases (#10) |
| Required stable funding, residual maturity and encumbrance of assets | S-01 Chapter 4 Section 1 (Articles 428p–428q); the exact provisions to confirm on EUR-Lex | Asset classification, residual-maturity and encumbrance inputs (#9) | RSF factors, cited per Articles 428r–428ah | RSF classification, maturity and encumbrance boundary cases (#10) |
| Required stable funding of short-term transactions with financial customers | S-01 Article 510, as amended by S-03 | Transaction-type, counterparty and residual-maturity inputs (#9) | Article 510 treatment, cited | Contract example `valid-basic`; cases planned under PROD-02 and PROD-03 |
| Simplified calculation for small and non-complex institutions | S-01 Articles 428ai–428az | Out of v1 | — | — |
| Application on a consolidated basis | S-01 Article 428a | Out of v1 (entity scope) | — | — |
| Reporting layout and items requiring stable funding | S-01 Articles 427–428, Article 428(1)(k) as amended by S-04; R-02 Article 17 and Annex I section 7 | Output naming and ordering only | None | Output layout review (#9) |

<a id="applicability-decisions"></a>

## ✅ Applicability decisions

Decided by me on 2026-11-05 for version 1. Every threshold or band below
stays in the cited provisions and enters a parameter set only with its
citation; this register states none.

| Topic | Decision |
| --- | --- |
| Funding horizon | The horizon and the residual-maturity bands are those set by the ASF and RSF provisions (S-01 Articles 428i–428ah). The register adds no horizon of its own. |
| Available stable funding | Liabilities and own funds are classified only by S-01 Chapter 3; each factor is a cited parameter. |
| Required stable funding | Assets and off-balance-sheet items are classified only by S-01 Chapter 4; each factor is a cited parameter. |
| Residual maturity | Measured as S-01 Article 428j requires for liabilities and own funds, and as the RSF provisions of Chapter 4 require for assets. No default maturity is inferred for a missing value. |
| Encumbrance | Each asset carries its encumbrance status and remaining encumbrance period as inputs; the treatment follows the RSF provisions. |
| Currencies | One reporting currency per run, supplied by the caller; no currency conversion (see [scope](../SCOPE.md#users-and-use-cases)). Reporting by significant currency is out of v1. |
| Entity scope | A single institution on an individual basis, as decided above. |
| National discretions and supervisory decisions | None exercised in v1; a treatment that depends on a supervisory decision (Article 428f, for example) is out of v1. |

<a id="unresolved-interpretations"></a>

## ❓ Unresolved interpretations and follow-on work

| Item | Status and owner |
| --- | --- |
| Supervisory decision under Article 428f, and the exact provisions of Chapter 4 Section 1 on the residual maturity and encumbrance of assets | To verify on EUR-Lex before PROD-01 sets the parameter set; contract 1.0.0 (#9) carries only the dates. |
| Indirect effects of S-04's other CRR amendments on the NSFR inputs | Reviewed under PROD-01, for the inputs that Article 428(1)(k) and the RSF provisions refer to. |
| Inclusion of derivative contracts (Article 428d) and netting (Article 428e) | Deferred from contract version 1 by my decision in #9: such inputs are rejected with `NSFR-E012`; admitting them needs its own issue. |
| Factor values for ASF and RSF | Captured with citations in the parameter set of [PROD-01](../BACKLOG.md#later-calculation-work), never in this register. |
| Consolidated basis, group and scheme treatments, simplified calculation | Out of v1; each needs its own scope decision and follow-on issue before admission. |

<a id="use-of-the-register"></a>

## 🛡️ Use of the register

- How the README, fixtures, examples and parameter sets cite these entries is
  an editorial rule, in
  [AGENTS.md](../../AGENTS.md#what-the-documentation-describes).
- A change to this register, including a new source version, needs my
  review and a new register version; explanatory sources never override a
  normative one.
