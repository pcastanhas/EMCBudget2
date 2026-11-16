# Budget Redesign — Continuation File

**Written:** end of Session 4 (system date 11/12/2026). **Session 5 (11/16/2026): project moved from Claude chat to Claude Code + GitHub (`pcastanhas/EMCBudget2`); environment sections (§0, §4, §9, §12, §13) updated. No workbook change.**
**Current deliverable:** `E42_ALLWOOD_ASSOCIATES_2027_Budget_TEST_v3.13.xlsm` (entity 42 only; **29 sheets**). **Session 4: v3.12 replaced CAPITAL_PROJECTS with REPAIRS_CAPITAL (major repairs grid + capital improvements grid; §5.4); v3.13 adds collapsible detail rows on the 11 section sheets (§5.4). Neither is confirmed in Excel by the user yet.** Session 3 versions: v3.8 DEBT sheet + DEBT SERVICE section + N.O.I. BEFORE CAPITAL; v3.9 OTHER_PROF_FEES sheet + OPERATING INCOME BEFORE DEBT SERVICE; v3.10 CAPITAL_PROJECTS sheet + NET CASH FLOW / PROJECTED CASH BALANCE + starting-cash input; v3.11 payroll, partner distributions and maintenance reserve inputs + new cash lines. **User confirmed in Excel:** v3.8 ("looks good"), v3.9 ("it worked"), v3.11 ("that worked"); v3.10 was used on the way (not separately confirmed).
**User guide (draft v0.1):** a Claude Doc named "Budget Template User Guide (Preliminary)" (claude.ai artifact; link in the Session 2 chat). It has no file; ask the user for the link if you need to edit it.
**Legacy reference (read-only):** `reference/Allwood-2026 Budget-FINAL.xlsm v2.xlsm`
**Repository layout:** `build_tools/` (build, test, render and lint scripts), `VBA_modules/*.bas` (5 modules), `workbook/` (latest `.xlsm`), `MRI_Files_2027/` (MRI extracts + building photos), `reference/` (legacy workbook). See `PROJECT_INSTRUCTIONS.md` → Environment and repository.

> **STANDING INSTRUCTION (user, Session 2; updated Session 5 for GitHub): at the end of every session, commit and push ALL changed files** — `PROJECT_INSTRUCTIONS.md`, `continuation.md`, the latest workbook, every changed build/test/render/lint script, `colors_detail.json`, and `VBA_modules/*.bas` — **directly to `main` (no pull requests; user instruction, Session 5)**. Each cloud session starts from a fresh clone, so anything not pushed is lost. (The Session 2–4 zip hand-off is retired.)

---

## 0. How to resume (do this first)

1. Read this file and `PROJECT_INSTRUCTIONS.md`, then the other files in the repository.
2. Confirm the latest workbook is v3.13 and that its sheet list matches §2 (29 sheets, COVER first). Report any mismatch before doing anything else.
3. The repository is cloned fresh for each cloud session; work in place (no build folder to unpack). Check that `MRI_Files_2027/` holds the MRI extracts listed in §4 — without them nothing builds or tests.
4. Check tools: `python3 -c "import oletools, uno, openpyxl, lxml, pandas, PIL"`. LibreOffice + Python UNO, `openpyxl`, `lxml`, `pandas`, `Pillow`, `pdftotext` are in the cloud image; `oletools` comes from the environment setup script — if it is missing, `pip install oletools`.
5. **Paths:** all scripts take their locations from `build_tools/paths.py` (repo-relative): VBA source `VBA_modules/`, MRI extracts `MRI_Files_2027/` (override with `--src` on `build_v2.py` or the environment variable `BUDGET_SRC` for every script), `colors_detail.json` in `build_tools/`. Run the scripts from `build_tools/`; put scratch outputs (base `.xlsx`, `lo.xlsm`, test copies, PDFs) **outside the repository** (e.g. the session scratchpad) so they are not committed.
6. Before any LibreOffice test: `rm -f .~lock.*` in the folder of every file being loaded (stale lock files make `load()` return `None`; see §10). Never run `pkill -f soffice` (it kills your own shell because the command line contains "soffice"); use `pkill -x soffice.bin`.
7. Run tests in small batches with `timeout`; a foreground command is limited to 600 s (default 120 s — pass a larger timeout), and long tests can run as background commands.
8. Do not create or change files until the change is agreed and the user says go (standing rule). Ask at most one question per reply.
9. Start with the open items in §7 unless told otherwise. The user said the rest of Session 2 was entity 42 only; confirm before building other entities.

---

## 1. Goals

| # | Goal | Status |
|---|------|--------|
| 1 | Fix broken/unused references and formulas | New workbook built clean; legacy untouched |
| 2 | Separate user input from budget worksheets | Done: INPUT, REVENUE, OTHER_REVENUE, six EXP_ sheets, TAXES, INSURANCE, PROF_FEES are inputs; DETAIL/QTRLY SUM are macro-built output |
| 3 | Flexible for future years and other buildings | Build script is parameterized by building (`--building`), year, source folder; DATA tabs hold one building's rows; no entity/year hardcodes in formulas. Test scripts and some file names still year/entity specific (§7) |
| 4 | Document current and future state | User guide v0.1 written (Claude Doc). This file is the technical record. Legacy-disposition document and MRI refresh procedure still to write |
| 5 | VBA only to build presentation schedules | Holds. Entry sheets are VBA-free; macros write formulas and formatting only; workbook calculates with macros disabled (new accounts appear on DETAIL only after Build Schedules) |

Long-term vision (user): *"Create the 2027 budget file for WEST MOUNT PLEASANT ASSOCIATES"* → the template workbook with only that building's data. Today: `python3 build_v2.py out.xlsx --building 05 …` + the LibreOffice pipeline does this; the one-command wrapper (`make_budget.py`) is the next tooling task.

---

## 2. Current workbook (v3.13), tab order is required

| # | Tab | Colour | Role |
|---|-----|--------|------|
| 1 | COVER | Navy `1F497D` | Building name, budget year, photo, status panel, linked list of all other sheets (28 links) |
| 2 | INPUT | Orange | Settings, Lease Options, Build Schedules button |
| 3 | REVENUE | Orange | Tenant revenue |
| 4 | OTHER_REVENUE | Orange | Non-tenant income |
| 5–10 | EXP_UTILITIES, EXP_JANITORIAL, EXP_ELEVATOR, EXP_HVAC, EXP_LANDSCAPING, EXP_GENERAL_RM | Orange | Expense input by GL account and vendor |
| 11 | TAXES | Orange | Real estate taxes, escrow, assessments |
| 12 | INSURANCE | Orange | Insurance, consulting |
| 13 | PROF_FEES | Orange | Management fee, accounting, bank charges, other G&A |
| 14 | OTHER_PROF_FEES | Orange | OTHER PROFESSIONAL FEES accounts: prior-year forecast × (1 + 1%), override by month, adjustments, 5 User blocks |
| 15 | DEBT | Orange | Primary mortgage inputs and roll-forward; blocks for the DEBT SERVICE accounts (interest, principal) with forecast, adjustment and override lines |
| 16 | REPAIRS_CAPITAL | Orange | Major repairs (40 lines, by GL account → that account on DETAIL + memo line on EXP_ blocks) and capital improvements (25 lines → one capital account, default 1345000, DETAIL section 3) |
| 17 | DETAIL | Purple | Macro-built budget by account and month |
| 18 | QTRLY SUM | Purple | Macro-built by quarter (unformatted, no summary lines yet) |
| 19–21 | TENANT_SCHEDULE, INCOME_ACCOUNTS, EXPENSE_ACCOUNTS | Blue | Calculations |
| 22–29 | DATA_COA, DATA_CHARGECODES, DATA_BUDGET_PY, DATA_ACTUALS, DATA_BALANCES, DATA_SUITES, DATA_TENANT_CHARGES, DATA_VENDOR | Blue | MRI data (one building) |

All sheets protected, no password. Input cells unlocked. DATA tabs are protected too (UNO tests must `unprotect("")` before writing).

**What the user has confirmed in Excel:** v2.6 and v2.7 ("that worked"), v2.8, v2.9, v3.0–v3.2 ("looks good"). v3.4 failed to compile in Excel (module-level `Private` declarations placed after procedures); fixed in v3.5. The user kept working on v3.5/v3.6 without reporting errors; the user's small screenshot "from v3.7" showed no error. **Session 3: v3.8 ("looks good"), v3.9 ("it worked") and v3.11 ("that worked") were opened in Excel with no problems reported;** v3.10 was not separately confirmed. Always ask for a Build Schedules click and any verbatim message after the next change.

---

## 3. Conventions (agreed)

- **Tabs:** orange input (front, after COVER), purple schedules, blue MRI/derived; COVER navy.
- **Cells (working sheets):** white = read-only; light blue `E6F0FF` = input; amber `FFD966` = flag/adjustment; **light green `E2EFDA` = forecast from a lease option (not MRI)**. Gridlines off.
- **Budget schedules** (DETAIL, QTRLY SUM) use their own palette (Calibri 10, Accounting format, whole dollars unless INPUT!C6 = Yes).
- **IDs are text** (lease IDs, suites, charge codes, vendor IDs).
- **GL account is the key.** No category layer.
- **Signs:** MRI revenue negative; input sheets show revenue as **positive**, concessions **negative**, expenses **positive**. Tenant charges that post to an *expense or leasing-cost account* are **offsets**: DETAIL shows expense less offset (leasing: cost less offset).
- **Fixed-cell contract:** other sheets reference only fixed cells of generated sheets (DETAIL!B3 status, DETAIL!B4 last built).
- **Account classification (VBA `RevSec`):** account on REVENUE/OTHER_REVENUE → section by MRI Financial Format group (`FMT_GROUP`): TYPE INCOME or group INCOME → section 1; the five INPUT leasing-cost accounts or group TENANT IMP. & COMM → 3; the six expense groups → 4–9; TAXES 10; INSURANCE 11; PROFESSIONAL FEES/G&A 12. **Not on the Financial Format → flagged "IGNORED - NOT ON FINANCIAL FORMAT", excluded from every schedule, status CHECK DATA. On the format but group has no DETAIL section yet → flagged "IGNORED - SECTION NOT BUILT YET", excluded, no status alert** (user instruction; Allwood: 1527000 Due to/frm Tenant, $92).

---

## 4. MRI data

Source extracts now contain **all buildings** (user changed them in Session 2); the build filters to one building. They go in `MRI_Files_2027/` in the repository. File names (the build also accepts the old names):

| File | Contents / facts |
|------|------------------|
| `BUILDING_LIST.xlsx` | `bldgid`, `bldgname` only (59 buildings). **No management fee rates**; they are typed on PROF_FEES (Allwood 3.0% / 0.0%) |
| `CHART_OF_ACCOUNT.xlsx` | Derived DATA_COA tab (Sheet1, 13 columns) not the raw MRI extract; the build reads columns A–F and **re-derives G–K from `fin_groups.json`**. `HAS_ACTIVITY` is chart-wide and informational |
| `CHARGE_CODE_LIST.xlsx` | 456 codes, shared by all buildings |
| `FY_2026_BUDGETS.xlsx` | 13 entities; filtered by `entityid` |
| `FY_2025_ACTUALS.xlsx`, `Q1-Q3_2026_ACTUALS.xlsx` | Cash basis actuals, all entities |
| `2026_STARTING_BALANCES.xlsx` | 1/1/2026 balances (DATA_BALANCES; unused by schedules yet) |
| `SUITE_LIST.xlsx` | Now has Cleanable values on some suites (unused) |
| `TENANT_LIST.xlsx` | 1,291 rows across 32 buildings; one blank-amount Allwood row dropped vs v2.6 (harmless). Capacity 400 rows per building (build stops if exceeded) |
| `Q1-Q3_2026_EXPENSES_BY_VENDOR.xlsx` | **Vendor × GL account × month by check period** (`schk`/`hist`), all 36 entities, Jan–Sep 2026, `acctnum` has `MR` prefix. Ties to GL actuals to the penny on 37 of 43 Allwood expense accounts; the rest differ by journal-entry reclasses (user: no reconciling line wanted) |
| `building_42_image.jpg` | Cover photo (name pattern `building_<id>_image.jpg/.png`) |

Rules learned: RTRIM padded IDs; MRI sheet is `MRIQRY` (query text on `QUERY`/`Query`); all extracts have year-bearing file names (`…_2026_…`) that will need renaming or a script option next year. Allwood: Billing SF stays manual (157,500 vs MRI leasable 164,367).

---

## 5. Design decisions by sheet

### 5.1 INPUT
C3 Budget Year, C4 Building Name, C5 Building SF (billing; script default = MRI leasable, `--billing-sf` overrides; Allwood 157,500), C6 Show decimals, C8–C12 MRI reference/SF check, C15 actuals-through month (script default = last month with activity), C18/C19 status/last built, rows 20–21 button, **C24:C28 leasing-cost GL accounts**, **Lease Options rows 30–35: C31 Expired lease handling (Terminate/MTM/Renew), C32 renewal % (3.0%), C33:C35 accounts not carried forward (default 4150000)**, legend from row 37. Names: `BudgetYear`, `BuildingName`, `BillingSF`, `ActualsThru`, `LeaseOption`, `RenewPct`, `OPT_EXCL`.
**Session 3 INPUT blocks (v3.10–v3.11), all light-blue inputs with MRI-derived defaults computed by `build_v2.py`:** **C38 starting cash** (`StartCash`; default = MRI CASH BALANCES bank accounts excluding names containing "security": starting balance + actual activity through the actuals-through month = $1,326,172.89 for Allwood; includes the maintenance reserve; it is a Sept balance, not a 12/31 forecast — user should replace) and C39 read-only MRI reference (sums `EXPENSE_ACCOUNTS!AA`). **MAINTENANCE RESERVE rows 41–46:** C42 reserve balance at start (`ResStart`, default = MRI account named *reserve* = $100,000), C43:C46 quarterly additions (`ResAmt`, default 0) with posting months D43:D46 (`ResMonth`, default 3/6/9/12). **PARTNER DISTRIBUTIONS rows 48–53:** C49 Yes/No (`DistOn`), C50:C53 quarterly amounts (`DistAmt`, default = prior-year net activity on partner capital accounts 3xxxxxx excluding 3999000 ÷ quarters elapsed = $180,000) with payment months D50:D53 (`DistMonth`, default 1/4/7/10 — MRI posts distributions in Jan/Apr/Jul/Oct). **PAYROLL rows 55–61:** account fixed at build (HVAC & PAYROLL group, name contains "payroll", most activity = 5340000 Payroll Reimbursement); C56 annual budget (`PayAnnual`, default = prior-year budget $65,000), C57:D58 four 5-pay-week month numbers (`PayWeekMonths`, default 3/6/8/11 = legacy 2026), C59 summer-help monthly amount (`PaySummerAmt`, 0), C60/D60 summer first/last month (`PaySummerFrom/To`, 6/8), C61 read-only weeks budgeted (=48+count, flag E61 if not 52). **Legend moved to row 63.**

### 5.2 REVENUE
- In-place blocks are **sized to each lease's charge-code count** (no empty MRI lines, no empty blocks); status flags "MORE CODES THAN LINES" / "more MRI leases than lease blocks" if MRI later adds a code/lease (needs a refreshed workbook). 15 new-tenant slots follow.
- Block: header (suite, lease ID, tenant, SF, start, expires; **G/H blue override "If expired"/"Renewal %"; I/J applied option/%**), one MRI line + adjustment row per code, 3 ADD lines (code list), Total Rent for Lease, 5 COST lines.
- **Lease options:** TENANT_SCHEDULE column AG "FORECAST BASE" = amount in the expiry month (0 if vacated, ended, or account excluded). MRI line = schedule amount + (option ≠ Terminate and month after expiry: base × (1+% if Renew)). Light-green conditional format + flag "FORECAST - MTM/RENEWAL". Per-lease override cells; INPUT sets default.
- **Columns V/W (added v3.4):** V "Detail status" (IGNORED flags), W "Detail amount" = signed amount DETAIL uses (income and COST rows +, tenant-billed offsets on expense/leasing accounts −, ignored 0). Status check 2 compares DETAIL total with `SUM(REVENUE!W)` + OTHER_REVENUE inputs + section sheets.
- Scanned range rows 8–1500; tags in E: `MRI`, `ADJ`, `ADD`, `CALC`, `COST`.

### 5.3 OTHER_REVENUE
One block per INCOME-group account with prior-year actual/budget (8 for Allwood; count computed at build) + **5 "User" blocks** (dropdown of all income-group accounts, column C input; flags DUPLICATE ACCOUNT / NOT AN INCOME ACCOUNT; block numbers `U1..U5` as text so they are not counted as automatic blocks). Names `INC_LIST_NUM`, `INC_LIST_TXT`. Rows 9–600.

### 5.4 Expense sheets (EXP_*), TAXES, INSURANCE, PROF_FEES, OTHER_PROF_FEES, DEBT, REPAIRS_CAPITAL
- **EXP_ sheets:** block per account in the sheet's Financial Format group (TYPE EXPENSE) with prior-year actual, budget or vendor activity; per vendor a read-only payments row (`VEND`, vendor ID in column D) and a blue budget row (`INPUT`); 3 open lines; total; **read-only tenant-billed offset row (`OFFSET`) and net row (`NET`)**; flag `NEW VENDOR - REBUILD` in T of the block header when MRI has more vendor rows than the block shows. DETAIL reads `INPUT` rows. Rows 9–1500.
- **TAXES:** panel (assessed value/rate optional, increase 3.5%, four pay months 2/5/8/11, escrow deposit default = max monthly escrow posting, BIA increase 2.5%); blocks 1200000, 5710000, 5730000, 5735000, 5745000 with calculated rows; DETAIL reads `FINAL` rows. **Quirk (legacy behaviour):** with escrow > 0 the tax payments net inside the escrow row, so TOTAL TAXES ≈ 12 × deposit + assessments; a higher tax bill does not change the total.
- **INSURANCE / PROF_FEES:** blocks with header cells H (Increase %, defaults 2.15% insurance, 1.00% consulting, 5% accounting, 0% others) and J (Use calc? Yes/No); calculated row = prior-year forecast × (1+%); vendor rows + adjustment rows + 3 open lines; DETAIL reads `FINAL`. PROF_FEES panel (rows 9–13): fee % on base rent (3.0%), fee % on operating expense (0.0%), base rent accounts G12/H12 (4100000, 4150000), OE account G13 (4210000); blocks start row 16. Management fee = rates × REVENUE+OTHER_REVENUE monthly amounts of those accounts.
- **Known limitation:** prior-year forecast = actuals through the actuals-through month + budget after, so an installment budgeted in a past but unpaid month drops out (Allwood insurance Sept installment: default $38,041 ≈ $25k low). Documented in the user guide.
- `EXPENSE_ACCOUNTS` (calc) lists qualifying expense-group accounts incl. TAXES, INSURANCE, PROF_FEES groups and whether each has a block (`EA_NOBLOCK` feeds a status check).

- **OTHER_PROF_FEES (v3.9, orange, before DEBT):** rows 6–7 totals (FCST / FINAL), blocks from row 9 (12 rows each + 2 gap): one per OTHER PROFESSIONAL FEES-group EXPENSE account with prior-year actual/budget/vendor activity, excluding the INPUT leasing-cost accounts (5820400 stays in section 3) = 8 for Allwood (5485050, 5780000, 5780100 BAIT, 5810000, 5815000, 5815010, 5815050, 5820000), then **5 User blocks** (dropdown `OPF_LIST_TXT` built on `EXPENSE_ACCOUNTS!W:Y`; flags DUPLICATE ACCOUNT / NOT AN OTHER PROF FEE ACCT in T, counted by status check 9). Block header: H **Increase %** (default 1.00%), J Use calc?; rows: ACT, BUD, FCST, **CALC** (= forecast × (1+H)), **OVR** (light-blue override by month: blank = use calculated, typed incl. 0 replaces it), **USED** (override else calc), adjustment row + 3 open lines (`INPUT`), total (`FINAL`, = USED + INPUT rows; D = account). No vendor rows. Default total $124,619.54 (forecast $123,385.68 × 1.01). Legacy BAIT installments typed on the OVR line reproduce $139,900. Legacy-only accounts with no activity (5813000, 5816000, 5820700, 5828000, 5853100 fitness offset) get no automatic block (use User blocks; the fitness offset is not built).
- **REPAIRS_CAPITAL (v3.12, orange, after DEBT; replaces CAPITAL_PROJECTS; modelled on the Eisenhower REPAIRS-CAPITAL sheet):** ONE LINE PER PROJECT, no add/delete (sheet protected, no password; only light-blue cells unlocked). **Major repairs** rows 12–51 (40 lines; header row 11), **capital improvements** rows 59–83 (25 lines; header row 58; GL account for all capital lines in **D56**, default 1345000, name in F56). Summary rows 7 (repairs) and 8 (capital), totals rows 52 and 84. Columns: A line, **B Include? (Yes/No)**, C description, **D GL # (calc)**, **E tag (calc: REPAIR / IGNORED for repairs; FINAL for capital)**, F GL account (repairs: dropdown `RC_LIST_TXT` = EXPENSE accounts in the six expense sections, 167 for Allwood; capital: locked name), G year, H Quote/Estimate/Allowance (capital: priority 1/2/3/Next year), I total cost, J spent in prior years, K budget-year amount (blank = cost less spent; typed incl. 0 replaces), L amount used (calc), M start month, N months (default 1), O remaining after budget year (calc), **P:AA 12 monthly input cells**, **AB:AM budget by month (calc; the columns DETAIL/EXP_ read)**, AN total, **AO Line check**, AP notes. **All-or-nothing monthly rule:** if ANY of P:AA holds a number those 12 cells are the schedule (blank month = 0; start month and months ignored); otherwise L spreads evenly from the start month over N months with the rounding remainder in the last month. Include = No → the line budgets 0. **Line check flags** (prefix **CHECK -** raises DETAIL status check 14; **NOTE -** is information only): CHECK: no GL account, account not budgetable (tag IGNORED, line excluded from DETAIL), no start month and no monthly amounts, typed months ≠ budget-year amount (typed schedule still used), capital GL account not in the chart (D56 → AO56; capital lines then drop off DETAIL); NOTE: excluded (Include = No), spread runs past December, budget-year amount exceeds remaining cost, possible duplicate (same amount used and start month in either grid, e.g. sealant 45,000 in May). **Header must not start with the word CHECK** (COUNTIF is case-insensitive: 'Check' matched 'CHECK*'), hence 'Line check'. **How it feeds:** capital lines keep tag FINAL and the account in D (blank unless the account is in the chart) → `AddFromSheet` with fixed section 3 and `ExpSumTerm(15,…)`; major repairs (tag REPAIR) → `AddFromSheet` with section classified per account (`RevSec`), and `RepSumTerm` is added in `AcctMonthTerm` for section 3 and all expense sections (not income) and in `AcctTotalTerm`. **EXP_ sheets (six only):** each block has a locked MEMO row (tag `MEMO`) 'major repairs from REPAIRS_CAPITAL' between the OFFSET and NET rows; NET = inputs + tenant-billed offset + major repairs; block spacing is now rt+5 (INSURANCE/PROF_FEES unchanged). Dropped from the old sheet: the prior-year reference block. Repairs are limited to the six expense sections (no TAXES/INSURANCE/PROF/OPF/DEBT accounts).
- **Collapsible detail rows (v3.13, `add_outline` in `build_v2.py`, all 11 section sheets: six EXP_, TAXES, INSURANCE, PROF_FEES, DEBT, OTHER_PROF_FEES):** Excel row outline level 1, **collapsed by default**. Per block the ACT / BUD / FCST rows plus the FIRST VEND row (when it directly follows FCST) are one group whose visible summary row is that vendor's budget (INPUT) row; every other VEND row is its own one-row group above its budget row; where a CALC/payroll row follows FCST the group is ACT/BUD/FCST only (summary = that row). The budget row label becomes `<vendor> - <year> budget` (or `adjustment (+/-)` on INSURANCE/PROF_FEES), a formula using the same DATA_VENDOR lookup as the VEND row, normal font. CALC/OVR/USED/INPUT/FINAL/OFFSET/MEMO/NET rows are never grouped. Hidden rows still count in every SUMIFS/DETAIL formula; tests read rows by tag, not visibility. **Protection:** Excel keeps the +/- buttons dead on a protected sheet unless `EnableOutlining` is set in code each time the workbook opens, so `modBuildAll.Auto_Open` unprotects (no password), sets `EnableOutlining = True` and re-protects with `UserInterfaceOnly:=True` on `ExpSheetName(1..11)`. With macros disabled the groups stay collapsed and cannot be expanded (formulas unaffected). **Not tested in Excel.** `postprocess.py` `mark_collapsed` writes `collapsed="true"` on the summary row (LibreOffice writes `false`, which would show the button in the wrong state). EXP_ memo/NET labels shortened: '… major repairs (REPAIRS_CAPITAL, read-only)', '… net of tenant-billed offset + major repairs (as on DETAIL)'.
- **Payroll (v3.11):** the 5340000 block on EXP_HVAC gets one extra locked row tagged `INPUT` (so DETAIL sums it): `=N(PayAnnual)/52*IF(COUNTIF(PayWeekMonths,month)>0,5,4)+IF(month in summer range, PaySummerAmt, 0)`. Leave that account's vendor budget lines blank (legacy had payroll only from these inputs); open lines remain for adjustments. Default reproduces legacy months $5,000/$6,250 = $65,000.
- **DEBT (Session 3, orange, after OTHER_PROF_FEES):** layout like INSURANCE (rows 6–7 all-account totals tagged FCST / FINAL; month columns G:R; blocks from row 37; sheet order is now PROF_FEES, OTHER_PROF_FEES, DEBT, REPAIRS_CAPITAL). **Panel (blue inputs)** G10 lender (default = largest vendor on the principal account), **G11 interest GL account** (default = DEBT SERVICE expense account with most activity, 5910000), **G12 principal GL account** (default = DEBT SERVICE liability account, 2600000), **G13 interest rate** (default = rate implied by the latest MRI actual month, rounded to 5 decimals: Allwood 3.720%), G14 implied rate (read-only), **G15 day-count basis** (30/360 default, Actual/360, Actual/365), **G16 payment override** (blank = use G17), G17 latest MRI actual P&I payment (SUMPRODUCT on DATA_ACTUALS for the actuals-through month), G18 payment used, G19 maturity date (optional; I19 flags maturity in/before the budget year), **G20 opening-balance override** (T35 flags a difference from the prior-year ending balance). **Roll-forward (read-only):** rows 23–28 prior year (opening balance = −MRI starting balance on DATA_BALANCES for the principal account; actual interest/principal through ActualsThru, then calculated), rows 30–35 budget year (opening = prior-year ending unless G20 is typed). Interest = `ROUND(opening × rate × days / (360 or 365), 2)`; principal = `MAX(0, MIN(payment − interest, opening))`; days = 30 or calendar days. **Blocks** (one per DEBT SERVICE format-group account with activity, **any account type**): header H increase % and J Use calc; ACT / BUD / FCST rows, **CALC row** (interest account → row 33, principal account → row 34, any other debt account → prior-year forecast × (1 + H)), an adjustment row, three open lines, total tagged `FINAL` (DETAIL reads it). No vendor or offset rows. Account order on DETAIL is by account number (2600000 before 5910000). With no debt accounts the sheet shows the empty panel and DETAIL prints no debt section.

### 5.5 DETAIL (macro-built)
Sections (VBA `SecOrder`): 1 INCOME, six expense sections (4–9, palette EXPSUB), **TOTAL EXPENSES**, 10 TAXES, 11 INSURANCE, 12 PROFESSIONAL FEES/G&A, **TOTAL OPERATING EXPENSES**, **N.O.I. BEFORE OTHER PROFESSIONAL FEES & DEBT SERVICE**, **14 OTHER PROFESSIONAL FEES (TOTAL OTHER PROFESSIONAL FEES)**, **OPERATING INCOME BEFORE DEBT SERVICE** (= N.O.I. − fees), **13 DEBT SERVICE (TOTAL DEBT SERVICE)**, **N.O.I. BEFORE CAPITAL** (= operating income before debt service − total debt service), then 3 IMPRVMTS & COMMISSIONS (REVENUE cost lines + REPAIRS_CAPITAL capital lines; TOTAL IMPRVMTS & COMMISSIONS; shown only when amounts exist), then the cash block (always printed): **NET CASH FLOW** (= N.O.I. before capital − total imprvmts & commissions), plain lines *Less: partner distributions* (`=IF(DistOn="Yes",SUMIFS(DistAmt,DistMonth,month),0)`) and *Less: additions to maintenance reserve* (`SUMIFS(ResAmt,ResMonth,month)`), **NET CASH FLOW AFTER DISTRIBUTIONS & RESERVES**, **MAINTENANCE RESERVE BALANCE** (`ResStart` + cumulative additions), **PROJECTED OPERATING CASH BALANCE** (`StartCash − ResStart` + cumulative net after distributions & reserves), **PROJECTED CASH BALANCE** (= operating + reserve; reserve additions are transfers). Balance rows: total column = December, quarter columns = quarter-end. Section numbers: 13 DEBT, 14 OTHER_PROF_FEES (display order via `SecOrder`: …12, 14, 13, then 3). VBA helpers `DetailFlowRow`, `DetailBalanceRow`, `DetailTotalCashRow`, `BalanceEndCols`. The old "TENANT CHARGES POSTED TO EXPENSE…" section was removed (user): offsets are netted in the expense/leasing rows. Summary lines always print (zero if empty). Account row formula by section: 1 = REVENUE + OTHER inputs; 3 = COST rows − non-COST REVENUE; 4–13 = section-sheet amount − non-COST REVENUE for the account (flagged IGNORED rows excluded).
Status formula (DETAIL!B3) has **15 checks** (check 9 also counts OTHER_PROF_FEES User-block flags; check 10 covers expense, DEBT SERVICE and OTHER PROFESSIONAL FEES accounts without a block; **check 14** = a REPAIRS_CAPITAL line with a `CHECK -` flag in column AO; **check 15** = a distribution/reserve amount with no valid month 1–12) (display option, rebuild needed, unknown charge code, orphan REVENUE amount, orphan OTHER_REVENUE amount, too many income accounts, too many leases, too many charge codes/non-monthly, duplicate/non-income User block, expense account without block, vendor without line, TAXES pay month, not-on-format account).
QTRLY SUM: same accounts/sections by quarter (uses `AcctMonthTerm`); **no summary lines, still unformatted**.

### 5.6 COVER
Navy title band (`=BuildingName`), year band, photo (`--image` or `building_<id>_image.*`; grey placeholder if none), building panel (ID `=INPUT!C8`, billing SF, status `=DETAIL!B3`, last built `=DETAIL!B4`), worksheet index with `location` hyperlinks grouped by tab colour. `postprocess.py` finds the button drawing by name because the cover image becomes `drawing1.xml`.

### 5.7 VBA (5 modules in `VBA_modules/`)
`modBuildAll` (button + quiet entry), `modBuildDetail` (DETAIL, status, `DetailSummaryRow`), `modBuildQtrly`, `modBuildCommon` (constants **and module-level `Private` variables at the TOP**, `RevSec`, `GroupSec` (DEBT SERVICE = 13), `AddFromSheet`, `CollectAccounts`, `AcctMonthTerm`, `ExpSumTerm`, `AcctTotalTerm`, `ExpSheetName` 1–12 (10 = DEBT, 11 = OTHER_PROF_FEES, 12 = REPAIRS_CAPITAL; constants `RC_SHEET`, `RC_MCOL`=28), `SheetSec` (12 → section 3), **`SheetMonthCol`/`SheetTotCol` (REPAIRS_CAPITAL months AB:AM, total AN; every other sheet G:R / S), `RepSumTerm` (major repairs by account and month)**; `AddFromSheet` now takes the total column, `ExpTag`, `SecOrder`, titles), `modFormat` (palette, `SecPalette`, `PaintRow`).

---

## 6. Regression targets (entity 42, v3.11 defaults; **v3.12 matches all 46 DETAIL total rows exactly with both new grids empty**)

| Item | Amount |
|------|--------|
| DETAIL total income | $3,475,097 (base rent $3,221,163; concessions −$119,405; OE $166,629; electric $206,710; other tenant charges now netted into expense rows) |
| TAXES default total | $468,810 (budget-year tax $514,199 = prior-year forecast $496,811 × 1.035; escrow $38,729) |
| INSURANCE default | $41,660 (insurance $38,041 + consulting $3,619) |
| PROF_FEES default | $116,166 (management fee $93,053 = 3% × net base rent; accounting $22,601; bank $512) |
| TOTAL OPERATING EXPENSES (defaults) | **$680,930** ($615,930 + payroll $65,000); N.O.I. before other fees & debt **$2,794,167.25** (before payroll was added in v3.11: $615,930 / $2,859,167.25) |
| Lease options (base rent / total income) | Terminate $3,221,163 / $3,475,097; MTM $3,727,621 / $4,029,657; Renew 3% $3,742,815 / $4,046,294; Renew 5% $3,752,944 / $4,057,385 |
| Legacy tie (TAXES inputs 19,000,000 × 2.531%, +3.5%, escrow 38,729) | payments $120,222.50 ×2, $128,638.07 ×2; escrow row −$32,973; taxes $497,721 |
| Offsets | 5225000 tenant-billed $10,001; 5435000 $705 (net into Janitorial / General R&M) |
| **Debt (defaults: 3.72%, 30/360, payment $61,499.97 = Jan–Sep 2026 actual)** | 2026 interest $400,197.25, principal $337,802.39 (Jan–Sep actual + Oct–Dec calculated; ties to the legacy chain); **2026 ending = 2027 opening balance $10,573,976.03**; 2027 interest $387,414.51, principal $350,585.13, **TOTAL DEBT SERVICE $737,999.64**; 2027 ending balance $10,223,390.90 |
| **Other professional fees (defaults, +1%)** | 5485050 $222.20, 5780000 $1,212, 5780100 BAIT $112,402.90, 5810000 $2,322.20, 5815000 $5,085.90, 5815010 $251.49, 5815050 $1,311.74, 5820000 $1,811.09 = **$124,619.52**; typing the legacy BAIT installments (33,500 Jan, 39,400 Apr, 33,500 Jun, 33,500 Sep) = $139,900 |
| **Payroll 5340000 (defaults)** | $5,000/month, $6,250 in Mar/Jun/Aug/Nov = $65,000 (legacy tie) |
| **Operating income / N.O.I. before capital** | OPERATING INCOME BEFORE DEBT SERVICE **$2,669,547.71**; **N.O.I. BEFORE CAPITAL $1,931,548.08** |
| **Cash block (defaults)** | starting cash $1,326,172.89 (MRI, Sept); reserve $100,000; distributions $720,000 (Jan/Apr/Jul/Oct); net cash flow $1,931,548.08; net after distributions & reserves $1,211,548.08; Dec **operating cash $2,437,720.97, reserve $100,000, PROJECTED CASH BALANCE $2,537,720.97** |
| **Capital projects (test)** | 25,000 (Apr) + 72,000 (Aug) + 45,000 over 3 months from May + 10,000 over 3 from Oct on 1345000 and 10,000 (Feb) on 5860000 → total imprvmts $162,000; NET CASH FLOW drops by the same monthly amounts |
| Entity 05 (smoke test v3.10/v3.11) | builds, 0 formula errors; DEBT 2636000 + 5910000 (CONNECTONE BANK, implied 5.942%), payroll default $31,000, distributions $41,667/quarter, reserve $0; status CHECK DATA only for budget-only TAXES account 5710100 (fixed TAXES block list). Its old v2.6 income check ($1,087,970) predates all this |

---

## 7. Open items

### Session 5 (migration to Claude Code + GitHub)
- **Done:** documents updated for the repository workflow (commit/push replaces the zip hand-off). `oletools` installed in the session; gate check on v3.13: embedded VBA identical to `VBA_modules/*.bas` (all 5), no `VBASupport`, `vba_lint.py` clean on both; workbook has the 29 sheets in §2.
- **Done (user):** uploaded the ten MRI extracts (§4) and `building_42_image.jpg` to `MRI_Files_2027/` and the legacy workbook to `reference/`; added `pip install oletools` to the cloud environment's setup script (confirm at the start of Session 6 that `import oletools` works without a manual install). Check: every extract opens with its `MRIQRY` sheet and has entity 42 rows; `build_v2.py --building 42 --billing-sf 157500 --src ../MRI_Files_2027` builds the base workbook (29 sheets, cover photo found). Keep the repository private.
- **Done (scripts):** chat-sandbox paths replaced by repo-relative ones (new `build_tools/paths.py`; `build_v2.py`, `final_build.py`, `make_xlsm.py`, `parse_fin.py`, the three render scripts and seven tests use it). `lo_driver.py` uses the optional `office.soffice` helper when present and a plain headless environment (`SAL_USE_VCLPLUGIN=svp`) otherwise. `.gitignore` added (`__pycache__`, lock files). **Pipeline check in the new environment:** build → `final_build.py` (status OK - schedules current) → `postprocess.py` reproduces v3.13: same 29 sheets, all 1,700 DETAIL cells identical, `colors_detail.json` unchanged.
- **Tests (Session 5, new environment): all 17 pass** (run one after another as a background command, ~32 min in total; slowest `test_repairs` 277 s, `test_fees` 195 s, `test_debt` 161 s, `test_cashflow` 155 s). Default totals equal §6 (total income $3,475,097.30, TOTAL OPERATING EXPENSES $680,930.05, N.O.I. $2,794,167.25, other professional fees $124,619.54). No workbook change, so v3.13 stays the deliverable.

### Waiting on a user decision
1. **July reset rule** for OE/electric reimbursement (flat in-place today).
2. **Crowe concession end date** (−$15,589.17 from 8/1/2027, open-ended).
3. Whether QTRLY SUM should get the TOTAL EXPENSES / TOTAL OPERATING EXPENSES / N.O.I. lines (offered, not answered).
4. "Back to cover" links on each sheet (offered).
5. Whether the user wants User blocks (pick-an-account) on the six EXP_ sheets for accounts with no history (built only for OTHER_REVENUE and OTHER_PROF_FEES).
6. Whether to raise CHECK DATA for "SECTION NOT BUILT YET" accounts (currently flag only).

### Session 4 (done / open)
- **Done:** REPAIRS_CAPITAL built to the design agreed with the user (one line per project; 12 optional monthly cells, all-or-nothing; Include?; flags; EXP_ memo lines; capital account cell). All 17 test scripts pass on v3.13 (`test_repairs` and `test_outline` new). VBA identical to source and lint-clean; schema-valid worksheets/workbook/styles; one tabColor per sheet; button still assigned; reopen + rebuild OK.
- **Open:** (0) user must also confirm in Excel that the +/- outline buttons work on the protected section sheets (depends on `Auto_Open`; fallback if it fails: leave those sheets unprotected, which breaks the input-in-allowed-areas standard, so avoid). (1) user must open v3.13 in Excel, click Build Schedules, and report any message verbatim (nothing Excel-specific tested: data-validation list on a defined name, conditional formats, status formula length 4,582 chars vs the 8,192 limit). (2) Possible-duplicate rule only compares amount used + start month (not lines with typed monthly amounts and no start month). (3) Load a real 2027 repairs/capital list for Allwood (both grids start empty). (4) User guide: add REPAIRS_CAPITAL. (5) Repairs on TAXES/INSURANCE/PROF/OPF/DEBT accounts are deliberately not allowed.

### Session 3 decisions (done) and open
- **Done (user decisions):** debt on its own DEBT sheet with account blocks and manual overrides; the 12/31 prior-year ending balance equals the 1/1 opening balance; OTHER_PROF_FEES defaults to prior-year forecast + 1% with a **monthly override line** on every block; capital items: leasing costs **stay on REVENUE cost lines**, a **CAPITAL_PROJECTS** sheet was added for building improvements; DETAIL finished with the two N.O.I. lines, NET CASH FLOW and cash block; payroll, distributions and reserves inputs added (user asked for these three only).
- **Legacy INPUT audit (done in Session 3, 36 cells):** covered now: year, building, cash (starting cash), mortgage (DEBT sheet), tax escrow (TAXES), assessed value/rate/increase/payments (TAXES), payroll, distributions, reserves. **Not built (user parked):** (a) **Snow Removal seasonal spread** — legacy spreads an annual amount ($43,000, account 5385000) by month: Jan 38.75%, Feb 37.71%, Mar 14.54%, Apr 0%, Dec 9% (names JANPER…DECPER); (b) **executive-summary / leasing inputs** — min $ for a repair on the exec summary ($12,000), min/max market rent ($22/$25 psf), min/max TI ($15/$35 psf); (c) **display/unused:** prior-year CAM budget bill $9.40 psf (BUDGET REVIEW, COMPARISONS), refi anticipated?, amortization text, R/E tax escrow balance $114,837 (unused), electric billing rate $1.75, loss factor 16%, common-area 15%, cleaning rate $0.076 psf (only the cleaning rate is displayed), security-deposit balance (legacy used it only to back S/D out of cash; new starting cash already excludes S/D bank accounts), prior-year quarterly distribution amounts (BUDGET REVIEW only). **Still to audit:** inputs on other legacy tabs (OP_EX, LEASPROJ, REPAIRS-CAPITAL major repairs list that feeds the expense accounts, 5 yr CapPLan, HVAC, FITNESS CTR).
- **Open (new):** debt accounts list in account-number order (2600000 before 5910000) — legacy/format order is interest first; **QTRLY SUM still has no summary lines and no cash lines** (legacy QTRLY cash block: cash excl. maintenance reserve, less distributions, reserve, use of TI/maint reserve, additions; **legacy quirk:** it subtracts reserve additions from total cash and also adds them to the reserve, which double-counts — new DETAIL treats additions as transfers); the starting-cash default is a September MRI balance (user must enter expected 12/31 cash); 5-week payroll months default to the legacy 2026 months; legacy fitness-center offset (40% of fitness revenue) not built; TAXES sheet has fixed blocks for five accounts (breaks other buildings, e.g. 05's 5710100); one INPUT label (B39) is clipped by its neighbour (cosmetic).
- Legacy facts (reference): debt rows 150–164 (interest = ROUND(balance × rate/360 × 30, 2), principal = 61,499.97 − interest hardcoded, 5857000 row dead); OTHER PROFESSIONAL FEES rows 135–148 (typed monthly BAIT installments, annual/12, fixed monthly, prior-year-month lookup, a `=row21*0.4` BAIT formula pointing at an empty row, Fit Ctr = 40% of fitness revenue); IMPRVMTS rows 166–173 (1345000 from REPAIRS-CAPITAL capital list rows 37–49 with month flags; 1365000/2105000/5390000/5820400/5860000 from REVENUE cost rows fed by LEASPROJ per-suite assumptions); NET CASH FLOW row 176 = N.O.I. before capital − total imprvmts; PROJECTED CASH BALANCE row 179 = STCASH (PVExecSum J37) + cumulative; legacy payroll OP_EX row 44 (5340000) = ANNPAY/52 × (5 if month in PAYMOS else 4) + SUMMER in Jun–Aug; snow removal OP_EX row 55 = annual × month % (Jan/Feb/Mar/Apr/Dec).

### Build next (suggested order)
1. **`make_budget.py` wrapper**: one command (building name/ID, year, source folder) → build → LibreOffice → post-process → gates → output named by building and year; regenerates `colors_detail.json` per building (it is a per-build file; **re-run `final_build.py` on the Allwood base before committing, because smoke tests on other buildings overwrite it**).
2. **Format QTRLY SUM** and add its summary and cash lines (TOTAL EXPENSES, TOTAL OPERATING EXPENSES, N.O.I. lines, net cash flow, distributions, reserves) — user has not answered whether to add them.
3. Remaining legacy inputs the user parks/approves: Snow Removal seasonal spread, executive-summary and leasing-rate inputs, display items; **audit OP_EX / LEASPROJ / REPAIRS-CAPITAL inputs**; per-lease leasing-cost assumptions (legacy LEASPROJ) if the user wants them in place of typed REVENUE cost lines.
4. **Sections for DISTRIBUTIONS/(CONTRIBUTIONS), OTHER DISB/(RECEIPTS), CASH BALANCES** (so flagged tenant accounts such as 1527000 land somewhere).
5. Make TAXES blocks data-driven (other buildings); make tests building-independent (several scripts hard-code Allwood row addresses/expected amounts: `test_or.py`, `test_expenses.py` totals text, `test_lease_option.py` account list, `test_fees.py`, `test_repairs.py`, `test_cashflow.py`, `test_debt.py` use Allwood account numbers and amounts); year-independent file names; rebuild entity 05 as a deliverable.
6. MRI upload sheet (legacy `FrmBudgets` replacement; sign rule: flip every 4xxx except 4310000, 4315000, 4320000, 4340000).
7. Documentation: finish/review the user guide (add DEBT, OTHER_PROF_FEES, REPAIRS_CAPITAL, cash inputs); legacy-sheet disposition; MRI refresh procedure; per-building rates (management fee).
8. Print layout/page breaks; "Back to cover" links; July reset rule and Crowe concession end date (still waiting on the user).

### Known data/format notes
- 5260100 Electric Survey Offset (TYPE EXPENSE, group INCOME) lands in INCOME. 5820400 Legal-Leasing is in OTHER PROFESSIONAL FEES in the format but is a leasing-cost account on INPUT (section 3).
- 4310000 Electric Reimbursement stays in INCOME (TYPE INCOME) although the format prints it under UTILITIES ("net electric" presentation still open).
- Vendor payments are cash basis by check period and include non-GL items outside the six sections (debt service, TI, taxes, distributions: $1.98M on 28 other accounts for Allwood) that later sections will use.
- Two accounts are included by two format lines (2635000, 2650000); 535 COA accounts have no `FMT_GROUP`.
- TENANT_SCHEDULE capacity 400 rows; REVENUE/EXP sheets scanned to row 1500; OTHER_REVENUE to 600.

---

## 8. Legacy workbook findings (reference)

33 sheets, one VBA project; ten external links; ~140 `#REF!`/external names; cached errors on several sheets; DETAIL row 21 check always FALSE; BAIT row uses a boolean row; month-date drift; square footage in four places; REVENUE fed by thousands of typed constants. Management fee logic: rate (BLDGLIST MFBR/MFOP) × monthly base rent+concession (REVENUE row 675) + MFOP × operating expense (row 676). Taxes: assessed value × rate × (1+increase), quarterly payments (prior-year/4 ×2, remainder/2 ×2) in Feb/May/Aug/Nov, escrow row = monthly escrow − taxes. Insurance/accounting: reference-month × (1 + typed %).

---

## 9. Build pipeline

Scripts live in `build_tools/`; VBA source in `VBA_modules/`; MRI extracts in `MRI_Files_2027/`; locations come from `build_tools/paths.py` (§0 step 5). Run from `build_tools/`:

| Script | Purpose |
|--------|---------|
| `parse_fin.py`, `fin_layout.txt`, `fin_groups.json` | Financial Format → account group map (`build_v2.py` re-derives COA columns G–K from the JSON) |
| **`build_v2.py`** | `python3 build_v2.py OUT.xlsx --building <ID or name> [--year 2027] [--src DIR] [--billing-sf N] [--actuals-thru N] [--image FILE] [--test]`. Resolves the building from `BUILDING_LIST`, filters every extract to it, sizes REVENUE/OTHER_REVENUE/EXP blocks from the data, writes all sheets incl. COVER |
| `lo_driver.py`, `make_xlsm.py` | LibreOffice headless helpers; embed the VBA project |
| `final_build.py` | Embed VBA, run `BuildAllSchedulesQuiet`, add the button, export DETAIL font colours to `colors_detail.json`, save `.xlsm` |
| `postprocess.py` | Tab colours (all 29 sheets), button macro assignment (drawing found by name), DETAIL font colours into styles, **`mark_collapsed` (collapsed flag on outline summary rows, v3.13)** |
| `final_reopen.py` | Reopen finished file, import its VBA, rerun the build |
| `vba_lint.py` | **Excel-only compile checks LibreOffice does not enforce** (module-level declarations after procedures; undeclared assigned variables). Run on `vba/*.bas` and on the VBA extracted from the finished file |
| `render_sheets.py` | `python3 render_sheets.py book out.pdf "SHEET:c1,r1,c2,r2" … [--set "INPUT!C31=MTM"] [--build]` → PDF for a visual check (`pdftoppm -png` to view) |
| `render.py`, `render_cover.py` | Older fixed-area renderers (kept for reference) |
| Tests (all take the base `.xlsx` path; they embed the VBA themselves) | `test_flow2` (status flips, rebuild), `test_dec` (show decimals), `test_or` (OTHER_REVENUE), `test_capacity` (lease/code capacity flags), `test_user_blocks`, `test_lease_option` (Terminate/MTM/Renew vs independent calc), `test_expenses` (vendor tie-out, flow, new-vendor/no-block flags), `test_taxes`, `test_insurance`, `test_proffees`, `test_offsets` (offset netting, not-on-format flag), `test_totals` (TOTAL EXPENSES / OPERATING EXPENSES / N.O.I.), **`test_debt`** (roll-forward vs independent calc from the raw extracts; rate, payment, balance override, day count, Use calc, actuals-through, maturity flag, DETAIL/QTRLY flow) |
| Session 4 tests | `test_outline` (reads the finished .xlsm, no LibreOffice: reference/VEND rows grouped+hidden, input/total rows visible, summary rows collapsed, vendor budget labels) — run as `python3 test_outline.py FILE.xlsm` (takes the finished workbook, not the base .xlsx) |
| More tests (Session 3) | `test_fees` (OTHER_PROF_FEES: defaults, increase %, override, Use calc, adjustments, User blocks, flags, the two N.O.I. lines), `test_repairs` (Session 4; replaces `test_capital`: repairs → DETAIL accounts + EXP_ memo lines, capital → section 3, rebuild flips, all-or-nothing months, Include = No, flags, duplicate note, capital account change, protection), `test_cashflow` (payroll, distributions, reserves → cash lines); `test_debt` and `test_fees` take ~4 min each, so run them alone (chat tool calls were limited to 300 s; Claude Code allows up to 600 s per foreground command) |

Run order (from `build_tools/`, `$S` = a scratch folder outside the repo): `build_v2.py $S/base.xlsx --building 42 --billing-sf 157500` → `final_build.py $S/base.xlsx $S/lo.xlsm` → `postprocess.py $S/lo.xlsm $S/out.xlsm colors_detail.json` → gates (§12) → save to `workbook/` with the next version number, delete the superseded version, sync scripts → commit and push to `main`.

---

## 10. Lessons learned

**Session 4:**
- **LibreOffice writes `collapsed="false"` on outline summary rows** (and drops the flag you set in openpyxl); patch it in `postprocess.py`. Group/hidden/outlineLevel survive the LibreOffice round trip.
- Excel's `EnableOutlining` is not saved in the file; it must be re-applied in `Auto_Open` with `UserInterfaceOnly:=True` after an Unprotect.
- A batch of three tests in one call can exceed the 300 s limit even when each takes ~100 s: run `test_user_blocks` + `test_lease_option` together at most, and `test_cashflow`, `test_fees`, `test_debt`, `test_flow2`, `test_or`, `test_repairs` one per call.
- **A stray character after a VBA line-continuation underscore (`_\`) makes LibreOffice run the macro silently and build NOTHING (DETAIL stays 'Not built yet'); `vba_lint.py` does not catch it.** When a build produces no schedule, diff the patched VBA lines for `_\` first. Patch VBA through Python with raw strings (`r'''…'''`).
- `COUNTIF(range,"CHECK*")` is case-insensitive: a column header 'Check' counts. Name flag columns something that does not start with the flag prefix.
- Layout constants shared by several builders (`RC_*`) must be defined before the first builder that uses them (placed beside `EXP_FIRST`).
- A flag formula that nests many IFs: count `(` vs `)` programmatically after generating it (one extra `)` made every flag `#VALUE!` and tripped the status).
- Accounts that the status checks count (`D` ≥ 1,000,000 with a tag) must also be accounts DETAIL can show: make the helper blank when the account is invalid (capital D = account only if in `COA_ACCTNO`; repair tag = IGNORED), otherwise the status shows REBUILD NEEDED instead of CHECK DATA.
- `test_flow2`, `test_or`, `test_cashflow`, `test_fees` exceed 90 s: run them one per call with `timeout 280`.


**Session 5:**
- Claude Code cloud sessions start from a fresh clone each time: only what is committed and pushed survives. Packages installed by hand (e.g. `oletools`) do not persist; put them in the environment setup script.
- Foreground commands may run up to 600 s (pass `timeout`); long tests can also run as background commands.

**Packaging (earlier sessions):** no `Option VBASupport 1` in exported VBA (feed module source with `\n`); LibreOffice drops tab colours on protected sheets (postprocess sets them; never add a second `<tabColor>`); stale theme colours (export live colours); VBA project must be created with `VBACompatibilityMode`, `ProjectName = "VBAProject"`, modules registered `ModuleType = 1`; UNO writes bypass protection (but see below).

**Session 2:**
- **Excel is stricter than LibreOffice for VBA.** Module-level `Private`/`Public`/`Dim` declarations placed after the first procedure compile in LibreOffice but give "Variable not defined" in Excel. Keep all module-level declarations at the top. `vba_lint.py` now catches this; run it on every build.
- **Stale LibreOffice lock files** (`.~lock.<file>#`) after a killed run make `loadComponentFromURL` return `None` (symptom: `'NoneType' object has no attribute 'BasicLibraries'`). Delete them before each run. Do not use `pkill -f soffice`.
- **Sheet protection blocks UNO writes on DATA tabs** (test steps that edit DATA_COA etc. must call `sheet.unprotect("")`). Writing `""` to an unprotected cell works.
- **openpyxl treats a string starting with `=` as a formula** (note texts like "= prior-year tax…" produced `#VALUE!`). Never start note text with `=`.
- **Drawing numbering changes** when another sheet has an image (cover photo); locate the button drawing by its shape name.
- A test tool call is limited to 300 s; run LibreOffice tests one at a time with `timeout`.
- Internal hyperlinks created with openpyxl `Hyperlink(location=…)` survive the LibreOffice round trip (`location` attribute kept).
- Tenant-billed charges on expense accounts must net against the expense (user decision); the earlier "show as billed" section double-counted conceptually and was removed.
- When the status check compares DETAIL with sources, any change to sign/netting rules needs a matching change in the source total (`AcctTotalTerm` uses `REVENUE!W`).
- Taxes escrow nets tax payments; test with the escrow deposit, not the tax increase.

**Excel/data behaviour (earlier):** `SUMIFS` criteria coerce numeric-looking text (use `"L"&leaseID`); avoid dynamic arrays; MRI budgets are rounded; filter end dates `>= 1/1/<year>`.

**Process:** compare embedded VBA to source, schema-validate XML, reopen and rebuild, run lint and tests, render and look; remove superseded versions; keep `VBA_modules` in sync; keep this file current.

---

**Session 3:**
- Add a new section in this order: sheet builder in `build_v2.py` (+ sheet list, tab colour, `postprocess.py` colour dict, COVER `SHEET_INFO`, EXPENSE_ACCOUNTS qualification / block count / `BUILT_GROUPS`), then VBA (`ExpSheetName`, the two `For i = 1 To N` loops, `GroupSec`, `SecOrder`, titles, `SecPalette`, `secTot` array size, the `oi` loops in DETAIL **and** QTRLY, summary line), then a test. `ExpTag` returns `FINAL` for any section ≥ 10, so a new section sheet reads its total rows tagged `FINAL`.
- A balance-sheet account (2600000 mortgage payable, TYPE LIABILITY) can sit in a DETAIL section; the EXPENSE_ACCOUNTS qualifier was widened to "EXPENSE type in the six/TAXES/INSURANCE/PROF groups **or** any type in DEBT SERVICE".
- Avoid circular references in calculation panels: the latest-payment cell reads DATA_ACTUALS directly (SUMPRODUCT with a column-number mask), not the roll-forward row that itself uses the payment.
- DATA_ACTUALS signs: payments to a liability are positive (debit); DATA_BALANCES liability balances are negative (credit) → opening balance = −starting balance.
- The `render_sheets.py` PDF has one page per sheet *with content*; locate a sheet's page with `pdftotext -f N` before viewing.
- Schema validation of `xl/sharedStrings.xml` against `sml.xsd` fails on the legal `xml:space` attribute in v3.7 through v3.11 alike (tooling quirk); validate worksheets, workbook and styles only.
- A stray `ws["I25"] = None` left in a builder blanked a roll-forward cell: after adding a sheet, print the whole roll-forward row, not only the ends.

**Session 3 (v3.8–v3.11):**
- The tool-call time limit is 300 s: run slow tests (`test_debt`, `test_fees`, `test_repairs`, `test_offsets`, `test_cashflow`, `test_flow2`, `test_or`) one per call with `timeout 280`; run the others in batches of 3–4 with `timeout 70–90`; write output to a file and `grep`/`tail` it.
- LibreOffice leaves `pkill`-able `soffice.bin` processes and lock files after a crash or timeout. Before every UNO run: `pkill -x soffice.bin; rm -f out/.~lock.* .~lock.*` (lock files for a loaded file live **in that file's folder**). Symptom: `embed_vba` fails with `'NoneType' object has no attribute 'BasicLibraries'`.
- Blank-versus-zero override pattern: `=IF(ISNUMBER(override),override,calc)` (blank = use calc, typed 0 = zero). "Use calc? = No" only zeroes the calc row.
- Adding a section sheet: sheet creation order, tab colour (`build_v2.py` list + `postprocess.py` dict), COVER `SHEET_INFO`, `EXPENSE_ACCOUNTS` qualification / block count (`_blocks` list) / `BUILT_GROUPS` (O column) / helper columns (W:AA now used), VBA (`ExpSheetName`, both `For i = 1 To N` loops, `GroupSec`, `SecOrder`, `SheetSec`, titles, `SecPalette`, `secTot` array, the `oi` loops in DETAIL **and** QTRLY, status-formula arrays), then a test. `ExpTag` returns `FINAL` for sections ≥ 10.
- A section sheet can feed an existing DETAIL section (REPAIRS_CAPITAL capital lines → 3) by passing a fixed section to `AddFromSheet` (`SheetSec`) and adding its `ExpSumTerm` in `AcctMonthTerm`; `AcctTotalTerm` must loop over it too (status check 2).
- In test helper scans, reset the current-block variable when a User block starts (User-block account cells are `""`, so the previous block would otherwise absorb their tags).
- Status order matters: REBUILD NEEDED (check 2) outranks data checks, so a bad User-block pick shows REBUILD NEEDED until rebuilt.
- Smoke tests on other buildings (`build_v2.py --building 05`) and on a synthetic no-debt case (`--fin-groups /tmp/fin_nodebt.json`, DEBT SERVICE group removed) are cheap and caught nothing new but confirm graceful degradation; they overwrite `colors_detail.json`.
- Every complete building (04, 05, 07, 12, 16, 21, 22, 23, 25, 33, 42, 49) has debt accounts; buildings missing from the budget/tenant extracts (08, 13, 18, 19, 20, 29…) cannot be built.
- Legacy payroll month rule uses 52 weeks (48 + 4×5-week months); fewer or more 5-week months leaves the annual budget short/over (INPUT C61 flags it).

## 11. Fixed cell and range contract (additions in bold)

| Where | What |
|-------|------|
| `INPUT!C3:C6`, `C15` | Names `BudgetYear`, `BuildingName`, `BillingSF`, `ShowDecimals`, `ActualsThru` |
| **`INPUT!C31`, `C32`, `C33:C35`** | **`LeaseOption`, `RenewPct`, `OPT_EXCL`** |
| `INPUT!C24:C28` | Leasing-cost GL accounts |
| `INPUT!C8` | MRI building ID (used by COVER) |
| `DETAIL!B3` / `B4` | Status / last built (also read by COVER and QTRLY SUM) |
| REVENUE rows 8–1500; D, E, S, U; **V (status), W (detail amount); header I/J applied option/%** | |
| OTHER_REVENUE rows 9–600; D, E, S, U; User blocks have text in U | |
| **EXP_ sheets, TAXES, INSURANCE, PROF_FEES rows 9–1500; D account; E tag; S total; T flags.** Tags: `ACT BUD FCST VEND INPUT CALC FINAL OFFSET NET`. DETAIL reads `INPUT` (EXP_) or `FINAL` (TAXES/INSURANCE/PROF_FEES) | |
| **TAXES panel G7:G17 (H12:H15 amounts); PROF_FEES panel G10:G13/H12; INSURANCE/PROF_FEES block header H (increase %), J (Use calc)** | |
| **DEBT panel G10:G20 (inputs G10:G13, G15, G16, G19, G20; read-only G14, G17, G18), roll-forward rows 23–28 (prior year) and 30–35 (budget year), account blocks from row 37; CALC row reads rows 33/34; total rows tagged `FINAL`** | |
| **OTHER_PROF_FEES blocks from row 9 (12 rows + 2 gap): header H increase % / J Use calc, ACT/BUD/FCST/CALC/OVR/USED, INPUT rows, FINAL total; User blocks `U1..U5` in column U; T flags; names `OPF_LIST_NUM/TXT`** | |
| **REPAIRS_CAPITAL (v3.12): repair lines rows 12–51, capital lines rows 59–83, capital account D56; D = GL #, E = tag (REPAIR/IGNORED/FINAL), L used, AB:AM budget by month (`RC_MCOL` = 28), AN total, AO line check; EXPENSE_ACCOUNTS AB:AD = repair dropdown list (names `RC_LIST_NUM`, `RC_LIST_TXT`)** | |
| **INPUT C38 `StartCash`, C42 `ResStart`, C43:C46 `ResAmt` / D43:D46 `ResMonth`, C49 `DistOn`, C50:C53 `DistAmt` / D50:D53 `DistMonth`, C56 `PayAnnual`, C57:D58 `PayWeekMonths`, C59 `PaySummerAmt`, C60 `PaySummerFrom`, D60 `PaySummerTo`; legend from row 63** | |
| **EXPENSE_ACCOUNTS: N11 DEBT SERVICE, N12 OTHER PROFESSIONAL FEES, O2:O14 `BUILT_GROUPS`, W:Y OPF dropdown list, Z:AA cash-account flag and balance** | |
| **TENANT_SCHEDULE AG (forecast base)** | |
| Names | `TS_*`, `CC_*`, `SUIT_*`, `COA_ACCTNO`, `COA_NAME`, **`COA_TYPE`, `COA_FMT`, `BUILT_GROUPS` (EXPENSE_ACCOUNTS O2:O13, now incl. DEBT SERVICE)**, `IA_ACCTNO`, `IA_RANK`, **`INC_LIST_NUM`, `INC_LIST_TXT`, `EA_NOBLOCK`** |

---

## 12. Verification checklist (before every delivery)

1. Rebuild from scripts; recalculate; zero formula errors (`#…` and `Err:`).
2. Key totals tie to independent calculations (§6) — the tests do this for lease options, taxes, insurance, fees, offsets, totals, vendor lines.
3. Run all 17 tests (batches, with `timeout`; `pkill -x soffice.bin` and remove lock files first; slow ones one per call).
4. Embedded VBA identical to `VBA_modules/*.bas` (extract with `oletools.olevba`), no `VBASupport`; **`vba_lint.py` clean on both**.
5. Validate every worksheet/workbook/styles XML part against `sml.xsd`; exactly one `<tabColor>` per sheet.
6. Reopen the finished file, import its VBA, rerun the build.
7. Check tab order, colours, protection, fonts/fills; confirm the button macro is still assigned (drawing found by name).
8. Render changed sheets and look (`render_sheets.py`).
9. Say plainly what cannot be tested without Excel; ask the user to open the file, click Build Schedules, and report any message verbatim.
10. Version every deliverable, remove superseded versions, sync `VBA_modules/` and `build_tools/`.
11. **At the end of the session: commit and push PROJECT_INSTRUCTIONS.md, continuation.md, the latest workbook, changed scripts and tests, `colors_detail.json` and VBA modules directly to `main` (no pull request).**

---

## 13. Repository inventory

`PROJECT_INSTRUCTIONS.md`, `continuation.md`, `workbook/E42_ALLWOOD_ASSOCIATES_2027_Budget_TEST_v3.13.xlsm`, `VBA_modules/` (5 `.bas`), `build_tools/` (`build_v2.py`, `postprocess.py`, `final_build.py`, `final_reopen.py`, `lo_driver.py`, `make_xlsm.py`, `parse_fin.py`, `vba_lint.py`, `render_sheets.py`, `render.py`, `render_cover.py`, 17 `test_*.py` (incl. `test_debt.py`, `test_fees.py`, `test_repairs.py` (replaced `test_capital.py`), `test_cashflow.py`, `test_outline.py`), `fin_groups.json`, `fin_layout.txt`, `colors_detail.json`). `MRI_Files_2027/` (the ten MRI extracts in §4 and `building_42_image.jpg`), `reference/Allwood-2026 Budget-FINAL.xlsm v2.xlsm` (legacy workbook, read-only). The user guide is a Claude Doc (not in the repository; ask the user for the link).
