# Project Instructions — Budget Redesign

## Purpose
Redesign the property budget workbook for budget year 2027 and every year after, as a **template**: budget preparers supply MRI data for a building and Claude produces a ready-to-use workbook for that building (target request: *"Create the 2027 budget file for <building name>"*). The legacy file (`Allwood-2026 Budget-FINAL.xlsm v2.xlsm`, kept in the repository under `reference/`) is a read-only reference. Audience of the finished workbook: property accountants and property managers.

## Goals
1. Fix broken or unused references and formulas.
2. Separate user input from the budget schedules.
3. Work for future years and other buildings: no hardcoded year, entity, or file names in formulas; the build script takes the building, year and source folder as parameters.
4. Document the current and future state (technical record = `continuation.md`; preparer-facing = the User Guide, a Claude Doc).
5. **VBA is allowed only to build the budget presentation schedules** (DETAIL, QTRLY SUM, future schedules). Entry sheets stay VBA-free. Macros write formulas and formatting, never values, so the workbook calculates correctly with macros disabled.

## Environment and repository
The project lives in the GitHub repository `pcastanhas/EMCBudget2` and is worked on with Claude Code (cloud sessions). The repository is the single source of truth; there are no Claude-chat project files or hand-off zips any more. Each cloud session starts from a fresh clone, so **anything not committed and pushed is lost** when the session ends.

| Folder | Contents |
|--------|----------|
| `/` | `PROJECT_INSTRUCTIONS.md` (this file), `continuation.md` (technical record) |
| `workbook/` | The current deliverable `.xlsm` (one version only; older versions live in git history) |
| `VBA_modules/` | The `.bas` modules (six since v3.15) (must match the VBA embedded in the workbook) |
| `build_tools/` | Build, post-process, test, render and lint scripts; `fin_groups.json`, `fin_layout.txt`, `colors_detail.json` |
| `MRI_Files_2027/` | MRI extracts (all buildings) and building photos (`building_<id>_image.jpg`) used to build the 2027 workbooks |
| `reference/` | Legacy workbook (read-only) |

The repository holds tenant and financial data and must stay **private**.

Tools the sessions need: Python 3 with `openpyxl`, `lxml`, `pandas`, `Pillow`, `oletools`; LibreOffice with Python UNO; `pdftotext`/`pdftoppm`. All except `oletools` are pre-installed in the cloud image; `oletools` is installed by the environment setup script (`pip install oletools`), or by hand at the start of a session if it is missing.

## Session protocol
- **Start of every session:** read `continuation.md` and every other file in the repository first. Confirm the workbook in `workbook/` matches the continuation file (version and sheet list) and report any mismatch before doing anything else. Do not re-ask questions the continuation file shows as settled. Check the tools are present (continuation §0).
- **Do not create or change any file until we agree on the change and I say go.** Discuss the design first (what the data is, what the input should look like, what the schedule will show), then build. A "can you…?" question is a request to explain and propose, not to build.
- **During work:** one area at a time. Be concise, ask at most one question per reply, state assumptions and open decisions plainly. Tell me what a change affects (other sheets, totals) before building it.
- **End of every session:** check documentation and code for consistency, then **commit and push directly to `main`**: updated `PROJECT_INSTRUCTIONS.md` (if the working agreement changed), updated `continuation.md` (decisions, what is done, open to-do list, lessons learned, notes), the latest workbook, every changed build/test/render/lint script, `colors_detail.json`, and `VBA_modules/*.bas`.
- **Git (user instruction, Session 5): never create a pull request. Push all changes directly to `main` when they are ready** (a session's working branch may be used along the way, but finished work goes to `main`). Push only work that is agreed and has passed the quality gates that apply to it.

## Workbook conventions (agreed, do not deviate)
- **Tabs:** COVER first (navy); orange = input sheets, then purple = budget schedules, then blue = MRI reference or derived data.
- **Cells (working sheets):** white = read-only; very light blue (`E6F0FF`) = user input; amber = flag or adjusted amount; light green (`E2EFDA`) = forecast created by a lease option (not MRI). Gridlines off. Every sheet is protected with no password; only input cells are unlocked.
- **Budget schedules are presentation output** and use their own colours (Calibri 10, Accounting format, whole dollars unless INPUT "Show decimals" = Yes, section palettes from the original Allwood DETAIL). The white/blue rule does not apply to them.
- **MRI data tabs are read-only** and replaced wholesale on refresh; each holds one building's rows. IDs (lease IDs, suites, charge codes, vendor IDs) are always text.
- **GL account is the key.** Charge code → GL account → amount. No intermediate category layer. Sections follow the MRI Financial Format group of the account.
- **Signs:** MRI stores credits negative; input sheets show revenue and expenses as positive; concessions negative. Tenant charges that post to expense or leasing-cost accounts are **offsets** and reduce that account on DETAIL. Accounts not on the Financial Format are flagged and ignored everywhere; accounts in groups with no DETAIL section yet are flagged and excluded until the section exists.
- **Generated sheets expose only fixed cells** (for example `DETAIL!B3` status). Never reference generated rows by address.
- **Inputs on calculated lines** use a read-only row (MRI or calculated) plus an adjustment row (numbers or formulas added to it). Calculated defaults (increase %, rates, "Use calc?" switches) are light-blue inputs with legacy values as defaults.
- Expenses by vendor are never reconciled to the GL (journal-entry reclasses are expected); vendor lines are shown as paid.

## Architecture (summary; details in `continuation.md`)
COVER → INPUT → REVENUE (tenant, via TENANT_SCHEDULE; lease options) and OTHER_REVENUE (non-tenant, with User blocks) → six EXP_ sheets (account × vendor), TAXES, INSURANCE, PROF_FEES, OTHER_PROF_FEES, DEBT (mortgage roll-forward + debt service account blocks), REPAIRS_CAPITAL (major repairs by GL account + capital improvements; capital lines feed IMPRVMTS & COMMISSIONS), plus INPUT blocks for starting cash, maintenance reserve, partner distributions and payroll → macro-built DETAIL (sections, TOTAL EXPENSES, TOTAL OPERATING EXPENSES, N.O.I., OTHER PROFESSIONAL FEES, OPERATING INCOME BEFORE DEBT SERVICE, TOTAL DEBT SERVICE, N.O.I. BEFORE CAPITAL, NET CASH FLOW and the cash balance lines) PROFORMA (budget year vs prior-year forecast and budget, variance, % and $/SF change, down to NET CASH FLOW) and QTRLY SUM. MRI extracts live in DATA_ tabs (one building). Calculation sheets: TENANT_SCHEDULE, INCOME_ACCOUNTS, EXPENSE_ACCOUNTS. Section grouping comes from the MRI Financial Format (`FMT_GROUP` on DATA_COA).

## Quality gates before delivering any workbook
1. Rebuild from scripts and recalculate: zero formula errors; key totals tie to independent calculations (the test scripts do this).
2. Run the workflow tests (status flips, rebuild, options, new input sheets, lease options, offsets, totals), in small batches.
3. Compare the embedded VBA with the source modules line by line (identical, no `Option VBASupport`) **and run `vba_lint.py`** on both — LibreOffice accepts VBA that Excel rejects (for example module-level declarations after a procedure).
4. Validate every XML part against the OOXML schema; exactly one `<tabColor>` per sheet.
5. Reopen the finished file, import its VBA, and rerun the build macro.
6. Check tab order, colours, protection, fonts and fills; confirm the Build Schedules button macro is still assigned; render and look at changed sheets.
7. Say plainly what could not be tested without Excel. Ask me to open the file, click Build Schedules, and report any message verbatim.
8. Version every deliverable (vX.Y) in the file name, delete the superseded version from `workbook/` (git history keeps it), and keep `VBA_modules/*.bas` and `build_tools/` in sync with the delivered workbook in the same commit.

## Technical notes
- **To create a building's workbook run `python3 make_budget.py <building ID or name>` from `build_tools/`**: it runs the whole pipeline and the automated quality gates (1, 3, 4, 5 and the checks in 6) and writes the file to `output/` (not committed). The Allwood test scripts (gate 2), visual review and Excel remain separate steps.
- The workbook is built with openpyxl (`build_v2.py --building …`), passed through LibreOffice (UNO) to embed and run the VBA and save `.xlsm`, then post-processed (tab colours, button assignment, font colours). Scripts run from `build_tools/` and find the VBA modules and MRI extracts through `build_tools/paths.py`; scratch outputs go outside the repository. Details, pipeline order and known pitfalls are in `continuation.md` §9–§10.
- Before any LibreOffice test delete `.~lock.*` files; never `pkill -f soffice` (use `pkill -x soffice.bin`). A foreground command is limited to 600 s (default 120 s), so run the slow tests one per call with `timeout`, or as background commands.
- MRI extracts come as Excel files (`MRIQRY` sheet plus a query tab) and now contain **all buildings**; the build filters to the selected building. Check each extract against the accounts the budget uses. Parameterize queries (entity, periods) for other buildings and years.
- Dynamic-array formulas (`FILTER`, spill) are avoided. Notes written into cells must never start with `=`.

## Do not
- Hardcode amounts, years, or entity numbers in formulas. Put them on an input sheet or in the data.
- Put user input on a budget schedule, or generated content on an input sheet.
- Ship a file that has not passed the quality gates.
- Use passwords on sheet protection.
- Reintroduce macros on entry sheets.
- Put project file names, query names or script names in preparer-facing documents (describe data as "MRI vendor payments", "MRI general-ledger actuals", and so on).
- Tell preparers to "rebuild the workbook" themselves: they request a refreshed workbook from their budget administrator.
