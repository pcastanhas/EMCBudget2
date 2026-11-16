import sys, datetime as dt
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.utils import get_column_letter as CL
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule

import argparse, os, json as _json
ap = argparse.ArgumentParser(description="Build the budget workbook (base .xlsx) for ONE building from all-building MRI extracts.")
ap.add_argument("out", help="output .xlsx path")
ap.add_argument("--building", required=True, help="MRI building/entity ID (e.g. 05) or building name (e.g. 'WEST MOUNT PLEASANT ASSOCIATES')")
ap.add_argument("--year", type=int, default=2027, help="budget year (default 2027)")
ap.add_argument("--src", default=None, help="folder holding the MRI extract files (default: $BUDGET_SRC, else <repo>/MRI_Files_2027)")
ap.add_argument("--billing-sf", type=float, default=None, help="billing/CAM SF for INPUT!C5 (default: MRI leasable SF of the building)")
ap.add_argument("--actuals-thru", type=int, default=None, help="months of prior-year-of-budget actuals for INPUT!C15 (default: last month with activity)")
ap.add_argument("--image", default=None, help="building photo for the COVER sheet (default: building_<id>_image.jpg/.png in the source folder or next to this script)")
ap.add_argument("--fin-groups", default=None, help="fin_groups.json (default: next to this script)")
ap.add_argument("--test", action="store_true", help="scratch copy with a test tenant filled in")
ARGS = ap.parse_args()
OUT = ARGS.out
TEST = ARGS.test
BUDGET_YEAR = ARGS.year
HERE = os.path.dirname(os.path.abspath(__file__))
FIN_GROUPS = ARGS.fin_groups or os.path.join(HERE, "fin_groups.json")
from paths import MRI_DIR
SRC_DIR = ARGS.src or MRI_DIR
UP = SRC_DIR.rstrip("/") + "/"

FONT = "Arial"
f_norm = Font(name=FONT, size=10)
f_bold = Font(name=FONT, size=10, bold=True)
f_title = Font(name=FONT, size=14, bold=True)
f_sub = Font(name=FONT, size=10, italic=True, color="595959")
f_lbl = Font(name=FONT, size=8, italic=True, color="7F7F7F")
f_note = Font(name=FONT, size=9, italic=True, color="7F7F7F")
INPUT_FILL = PatternFill("solid", fgColor="E6F0FF")
AMBER = PatternFill("solid", bgColor="FFD966", fgColor="FFD966")
RED = PatternFill("solid", bgColor="F8CBAD", fgColor="F8CBAD")
thin = Side(style="thin", color="000000")
hair = Side(style="hair", color="BFBFBF")
NUM0 = '#,##0;(#,##0);"-"'
NUM2 = '#,##0.00;(#,##0.00);"-"'
DATEF = "mm/dd/yyyy"
UNLOCK = Protection(locked=False)
ORANGE, PURPLE, BLUE = "FFA500", "7030A0", "0070C0"

wb = Workbook()
ws_in = wb.active
ws_in.title = "INPUT"
ws_cov = wb.create_sheet("COVER", 0)
ws_rev = wb.create_sheet("REVENUE")
ws_ld = wb.create_sheet("LEASE_DEALS")
ws_or = wb.create_sheet("OTHER_REVENUE")
EXP_SECTIONS = [(4, "UTILITIES", "EXP_UTILITIES"), (5, "JANITORIAL & CLEANING", "EXP_JANITORIAL"), (6, "ELEVATOR", "EXP_ELEVATOR"),
                (7, "HVAC & PAYROLL", "EXP_HVAC"), (8, "LANDSCAPING & GROUNDS", "EXP_LANDSCAPING"), (9, "GENERAL R&M", "EXP_GENERAL_RM")]
ws_exp = {n: wb.create_sheet(n) for _, _, n in EXP_SECTIONS}
ws_tax = wb.create_sheet("TAXES")
ws_ins = wb.create_sheet("INSURANCE")
ws_pro = wb.create_sheet("PROF_FEES")
ws_opf = wb.create_sheet("OTHER_PROF_FEES")
ws_debt = wb.create_sheet("DEBT")
ws_cap = wb.create_sheet("REPAIRS_CAPITAL")
ws_det = wb.create_sheet("DETAIL")
ws_pf = wb.create_sheet("PROFORMA")
ws_q = wb.create_sheet("QTRLY SUM")
ws_ts = wb.create_sheet("TENANT_SCHEDULE")
ws_ds = wb.create_sheet("DEAL_SCHEDULE")
ws_ia = wb.create_sheet("INCOME_ACCOUNTS")
ws_ea = wb.create_sheet("EXPENSE_ACCOUNTS")
data_names = ["DATA_COA", "DATA_CHARGECODES", "DATA_BUDGET_PY", "DATA_ACTUALS",
              "DATA_BALANCES", "DATA_SUITES", "DATA_TENANT_CHARGES", "DATA_VENDOR"]
ws_data = {n: wb.create_sheet(n) for n in data_names}
ws_in.sheet_properties.tabColor = ORANGE      # input sheets: orange, at the front
ws_rev.sheet_properties.tabColor = ORANGE     # REVENUE is an input sheet
ws_ld.sheet_properties.tabColor = ORANGE      # LEASE_DEALS: one line per deal (input)
ws_ds.sheet_properties.tabColor = BLUE        # DEAL_SCHEDULE: calculation
ws_or.sheet_properties.tabColor = ORANGE      # OTHER_REVENUE is an input sheet
ws_ia.sheet_properties.tabColor = BLUE        # INCOME_ACCOUNTS is MRI-derived data
ws_ea.sheet_properties.tabColor = BLUE
for w in list(ws_exp.values()) + [ws_tax, ws_ins, ws_pro, ws_opf, ws_debt, ws_cap]:
    w.sheet_properties.tabColor = ORANGE      # expense / tax input sheets
for w in (ws_det, ws_pf, ws_q):
    w.sheet_properties.tabColor = PURPLE      # budget schedules (macro-built)
ws_ts.sheet_properties.tabColor = BLUE        # TENANT_SCHEDULE is MRI-derived data
for w in ws_data.values():
    w.sheet_properties.tabColor = BLUE


def put(ws, ref, val, font=f_norm, fmt=None, inp=False, align=None):
    c = ws[ref]
    c.value = val
    c.font = font
    if fmt:
        c.number_format = fmt
    if inp:
        c.fill = INPUT_FILL
        c.protection = UNLOCK
    if align:
        c.alignment = Alignment(horizontal=align)
    return c


def name(nm, ref):
    wb.defined_names[nm] = DefinedName(nm, attr_text=ref)


# ------------------------------------------------------------------ MRI data tabs
def write_df(ws, df, fmts=None, text_cols=(), note=None):
    fmts = fmts or {}
    for j, col in enumerate(df.columns, 1):
        c = ws.cell(1, j, col)
        c.font = f_bold
        c.border = Border(bottom=thin)
    for i, row in enumerate(df.itertuples(index=False), 2):
        for j, v in enumerate(row, 1):
            if v is None or (isinstance(v, float) and pd.isna(v)) or v is pd.NaT:
                continue
            if isinstance(v, pd.Timestamp):
                v = v.to_pydatetime()
            c = ws.cell(i, j, v)
            c.font = f_norm
            col = df.columns[j - 1]
            if col in text_cols:
                c.number_format = "@"
            elif col in fmts:
                c.number_format = fmts[col]
    for j, col in enumerate(df.columns, 1):
        w = max(len(str(col)), *(len(str(x)) for x in df.iloc[:200, j - 1].tolist())) + 2
        ws.column_dimensions[CL(j)].width = min(max(w, 9), 48)
    ws.freeze_panes = "A2"
    if note:
        nc = len(df.columns) + 2
        ws.cell(1, nc, note + " READ-ONLY: to refresh, Review > Unprotect Sheet (no password), paste new data, then protect again.").font = f_note
        ws.column_dimensions[CL(nc)].width = 90
    ws.protection.sheet = True
    ws.protection.formatColumns = False
    ws.protection.formatRows = False



def _file(*names):
    for n in names:
        if os.path.exists(UP + n):
            return UP + n
    raise SystemExit("MRI extract not found in %s: tried %s" % (UP, ", ".join(names)))


def _read(path, key_cols=(), **kw):
    """Read the MRIQRY sheet (or the first sheet), all text unless dtype given."""
    xl = pd.ExcelFile(path)
    sh = "MRIQRY" if "MRIQRY" in xl.sheet_names else xl.sheet_names[0]
    df = pd.read_excel(path, sheet_name=sh, **kw)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _strip(df, cols):
    for c in cols:
        df[c] = df[c].astype(str).str.strip()
    return df


# ---- resolve the building (ID or name) against BUILDING_LIST
bl = _strip(_read(_file("BUILDING_LIST.xlsx"), dtype=str), ["bldgid", "bldgname"])
_b = ARGS.building.strip()
_hit = bl[bl.bldgid.str.lstrip("0").str.upper() == _b.lstrip("0").upper()] if _b.replace(" ", "").isalnum() and len(_b) <= 4 else bl.iloc[0:0]
if _hit.empty:
    _hit = bl[bl.bldgname.str.upper() == _b.upper()]
if _hit.empty:
    _hit = bl[bl.bldgname.str.upper().str.contains(_b.upper(), regex=False)]
if len(_hit) != 1:
    raise SystemExit("Building '%s' matched %d buildings. Candidates:\n%s" % (_b, len(_hit), bl.to_string(index=False)))
BLDG_ID, BLDG_NAME = _hit.iloc[0].bldgid, _hit.iloc[0].bldgname
PULLED = dt.date.today().strftime("%m/%d/%Y")
print("building:", BLDG_ID, BLDG_NAME, "| budget year", BUDGET_YEAR, "| source", UP)


def only_bldg(df, col):
    """Keep only the selected building; IDs compare as stripped text."""
    df[col] = df[col].astype(str).str.strip()
    out = df[df[col] == BLDG_ID].copy()
    if out.empty:
        raise SystemExit("No rows for building %s in column %s" % (BLDG_ID, col))
    return out.reset_index(drop=True)


# chart of accounts and charge codes are building-independent
coa = _read(_file("CHART_OF_ACCOUNT_1.xlsx", "CHART_OF_ACCOUNT.xlsx"), dtype=str)
coa = coa[[c for c in coa.columns[:6]]]            # A-F only; G-K are re-derived below
coa = coa[coa["ACCTNUM"].notna()]
for c in coa.columns:
    coa[c] = coa[c].str.strip()
coa["ACCTNO"] = coa["ACCTNO"].astype(int)
_fg = {int(k): v for k, v in _json.load(open(FIN_GROUPS))["accounts"].items()}
for _col, _key in (("FMT_GROUP", "group"), ("FMT_PARENT", "parent"), ("FMT_SUBTOTAL", "subtotal"), ("FMT_LINE", "line"), ("FMT_NOTE", "note")):
    coa[_col] = coa["ACCTNO"].map(lambda a, k=_key: (_fg.get(a, {}).get(k) or None))
write_df(ws_data["DATA_COA"], coa, text_cols=("ACCTNUM", "FORMATTED"),
         note="SOURCE A-F: MRI GACC chart of accounts (all buildings share one chart), pulled %s. HAS_ACTIVITY is chart-wide, informational only. SOURCE G-K: MRI Financial Format EMCSTD_V2 preview (finformat.pdf), " % PULLED +
              "resolved with parse_fin.py: FMT_GROUP = the format title the account prints under; FMT_PARENT = EXPENSES for expense sub-groups; "
              "FMT_SUBTOTAL = the subtotal line that rolls it up; FMT_LINE = format line number. Blank = account is not in the format. Re-derive G-K when the format changes. ACCTNO is numeric.")
cc = _strip(_read(_file("CHARGE_CODE_1.xlsx", "charge_code_1.xlsx", "CHARGE_CODE_LIST.xlsx"), dtype=str), None or [])
for c in cc.columns:
    cc[c] = cc[c].astype(str).str.strip()
write_df(ws_data["DATA_CHARGECODES"], cc, text_cols=tuple(cc.columns),
         note="SOURCE: MRI INCH/GLMT charge codes (shared by all buildings), pulled %s. Code must stay TEXT (leading zeros, 2E1). 70 codes default to MR4520000 (MRI fallback)." % PULLED)

# ---- building-specific extracts, filtered to the selected building
bud = only_bldg(_read(_file("FY_2026_BUDGETS.xlsx", "FY_BUDGET_1.xlsx"), dtype={"entityid": str, "budtype": str, "acctnum": str}), "entityid")
for c in ("budtype", "acctnum"):
    bud[c] = bud[c].str.strip()
bud["Year"] = bud["Year"].astype(int)
if set(bud.Year) != {BUDGET_YEAR - 1}:
    print("WARNING: budget extract year(s) %s, expected %d" % (sorted(set(bud.Year)), BUDGET_YEAR - 1))
write_df(ws_data["DATA_BUDGET_PY"], bud, text_cols=("entityid", "budtype", "acctnum"), fmts={m: NUM2 for m in bud.columns[4:]},
         note="SOURCE: MRI BUDGETS, entity %s, PRO, cash basis, pulled %s. Prior-year budget as loaded in MRI. Revenue negative." % (BLDG_ID, PULLED))
a25 = _read(_file("FY_2025_ACTUALS.xlsx", "2025_actuals.xlsx"), dtype={"entityid": str, "acctnum": str})
a26 = _read(_file("Q1-Q3_2026_ACTUALS.xlsx", "Q1-Q3_ACTUALS.xlsx"), dtype={"entityid": str, "acctnum": str})
act = only_bldg(pd.concat([a25, a26], ignore_index=True), "entityid")
act["acctnum"] = act["acctnum"].str.strip()
act["Year"] = act["Year"].astype(int)
write_df(ws_data["DATA_ACTUALS"], act, text_cols=("entityid", "acctnum"), fmts={m: NUM2 for m in act.columns[3:]},
         note="SOURCE: MRI GLSUM, entity %s, cash basis, pulled %s. Monthly activity; revenue negative. Year-end adjusting entries sit in December." % (BLDG_ID, PULLED))
bal = only_bldg(_read(_file("2026_STARTING_BALANCES.xlsx", "2026_beg_balances.xlsx"), dtype={"entityid": str, "acctnum": str}), "entityid")
bal["acctnum"] = bal["acctnum"].str.strip()
write_df(ws_data["DATA_BALANCES"], bal, text_cols=("entityid", "acctnum"), fmts={"starting_balance": NUM2},
         note="SOURCE: MRI GLSUM balance-forward rows, entity %s, period 202601. Balances at 1/1/2026; debit +, credit -. Pulled %s." % (BLDG_ID, PULLED))
su = only_bldg(_read(_file("SUITE_LIST.xlsx", "SUITES_fa575b.xlsx"), dtype={"bldgid": str, "suitid": str}), "bldgid")
su["suitid"] = su["suitid"].str.strip()
write_df(ws_data["DATA_SUITES"], su, text_cols=("bldgid", "suitid"), fmts={c: "#,##0" for c in ("Leasable", "Cleanable", "Usable")},
         note="SOURCE: MRI SUIT/SSQF, building %s, latest measurement as of %s. Leasable=GLA, Usable=USA. Suite ID must stay TEXT." % (BLDG_ID, PULLED))
te = only_bldg(_read(_file("TENANT_LIST.xlsx", "TENANTS_65d009.xlsx"), dtype=str), "bldgid")
for c in ("leasid", "suitid", "occpname", "inccat", "frquency"):
    te[c] = te[c].astype(str).str.strip()
for c in ("rentstrt", "expir", "vacate", "effdate", "enddate"):
    te[c] = pd.to_datetime(te[c], errors="coerce")
te["amount"] = pd.to_numeric(te["amount"])
if len(te) > 400:
    raise SystemExit("Building %s has %d tenant charge rows; TENANT_SCHEDULE capacity is 400" % (BLDG_ID, len(te)))
write_df(ws_data["DATA_TENANT_CHARGES"], te, text_cols=("bldgid", "leasid", "suitid", "occpname", "inccat", "frquency"),
         fmts={**{c: DATEF for c in ("rentstrt", "expir", "vacate", "effdate", "enddate")}, "amount": NUM2},
         note="SOURCE: MRI LEAS/CMRECC, building %s: charges effective <= 12/31/%d and not ended before 1/1/%d, current/new leases, pulled %s. " % (BLDG_ID, BUDGET_YEAR, BUDGET_YEAR, PULLED) +
              "Lease ID / suite / charge code must stay TEXT. Capacity 400 rows (A2:L401).")
_vf = _file("Q1-Q3_2026_EXPENSES_BY_VENDOR.xlsx", "EXPENSES_BY_VENDOR.xlsx")
ven = _read(_vf, dtype={"entityid": str, "vendid": str, "vendnme1": str, "acctnum": str})
ven["entityid"] = ven["entityid"].astype(str).str.strip()
ven = ven[ven["entityid"] == BLDG_ID].copy().reset_index(drop=True)
for c in ("vendid", "vendnme1", "acctnum"):
    ven[c] = ven[c].astype(str).str.strip()
VEN_MONTHS = list(ven.columns[4:16])
for c in VEN_MONTHS:
    ven[c] = pd.to_numeric(ven[c]).fillna(0.0)
if len(ven) > 2990:
    raise SystemExit("Building %s has %d vendor rows; DATA_VENDOR capacity is 2,990" % (BLDG_ID, len(ven)))
write_df(ws_data["DATA_VENDOR"], ven, text_cols=("entityid", "vendid", "vendnme1", "acctnum"), fmts={m: NUM2 for m in VEN_MONTHS},
         note="SOURCE: MRI checks (SCHK/HIST) by vendor and GL account, entity %s, cash basis by check period, pulled %s. Positive = expense. One row per vendor + account. Capacity 2,990 rows." % (BLDG_ID, PULLED))
# INPUT defaults derived from the data
MRI_LEASABLE = float(pd.to_numeric(su["Leasable"]).sum())
BILLING_SF = ARGS.billing_sf if ARGS.billing_sf is not None else MRI_LEASABLE
_mons = list(act.columns[3:15])
_py = act[act["Year"] == BUDGET_YEAR - 1]
if _py.empty:
    raise SystemExit("No %d actuals for building %s" % (BUDGET_YEAR - 1, BLDG_ID))
_has = [i + 1 for i, m in enumerate(_mons) if pd.to_numeric(_py[m]).abs().sum() > 0.004]
ACTUALS_THRU = ARGS.actuals_thru if ARGS.actuals_thru is not None else (max(_has) if _has else 12)
print("billing SF %s (MRI leasable %s) | actuals through month %d | tenant rows %d | suites %d" % (BILLING_SF, MRI_LEASABLE, ACTUALS_THRU, len(te), len(su)))
RE_DEFAULT = "30"

# ------------------------------------------------------------------ names
name("BudgetYear", "INPUT!$C$3")
name("BuildingName", "INPUT!$C$4")
name("BillingSF", "INPUT!$C$5")
name("ShowDecimals", "INPUT!$C$6")
name("ActualsThru", "INPUT!$C$15")
name("IA_ACCTNO", "INCOME_ACCOUNTS!$A$2:$A$2001")
name("IA_RANK", "INCOME_ACCOUNTS!$G$2:$G$2001")
for nm, col in (("TS_LEASE", "A"), ("TS_SUITE", "B"), ("TS_TENANT", "C"), ("TS_START", "D"), ("TS_EXPIRY", "E"),
                ("TS_CODE", "G"), ("TS_KEYRANK", "P"), ("TS_RANK", "S")):
    name(nm, f"TENANT_SCHEDULE!${col}$2:${col}$401")
name("CC_CODE", "DATA_CHARGECODES!$A$2:$A$700")
name("CC_DESC", "DATA_CHARGECODES!$B$2:$B$700")
name("CC_ACCT", "DATA_CHARGECODES!$D$2:$D$700")
name("SUIT_ID", "DATA_SUITES!$B$2:$B$200")
name("SUIT_GLA", "DATA_SUITES!$C$2:$C$200")
name("COA_ACCTNO", "DATA_COA!$C$2:$C$2000")
name("COA_NAME", "DATA_COA!$D$2:$D$2000")
name("COA_TYPE", "DATA_COA!$E$2:$E$2000")
name("COA_FMT", "DATA_COA!$G$2:$G$2000")
name("BUILT_GROUPS", "EXPENSE_ACCOUNTS!$O$2:$O$14")

# ------------------------------------------------------------------ INPUT
ws = ws_in
ws.sheet_view.showGridLines = False
for col, w in zip("ABCDE", (3, 46, 26, 4, 78)):
    ws.column_dimensions[col].width = w
put(ws, "B1", "BUDGET INPUT", f_title)
put(ws, "B2", "Light-blue cells are the only cells you edit. White cells are calculated or read-only.", f_sub)
put(ws, "B3", "Budget Year", f_bold)
put(ws, "C3", BUDGET_YEAR, f_bold, "0", inp=True, align="right")
put(ws, "B4", "Building Name", f_bold)
put(ws, "C4", BLDG_NAME, f_bold, inp=True, align="right")
put(ws, "B5", "Building SF (billing / CAM)", f_bold)
put(ws, "C5", BILLING_SF, f_bold, "#,##0", inp=True, align="right")
put(ws, "E5", "Manual input. Denominator for all $/SF. Compared with MRI leasable SF below.", f_note)
put(ws, "B6", "Show decimals on budget schedules", f_bold)
put(ws, "C6", "No", f_bold, inp=True, align="right")
put(ws, "E6", "Yes = amounts show cents; No = whole dollars. Rebuild the schedules after changing it.", f_note)
dv_dec = DataValidation(type="list", formula1='"Yes,No"', allow_blank=False)
ws.add_data_validation(dv_dec)
dv_dec.add("C6")
put(ws, "B7", "MRI REFERENCE (read-only)", f_bold)
ws["B7"].border = Border(bottom=thin)
ws["C7"].border = Border(bottom=thin)
put(ws, "B8", "MRI building ID")
put(ws, "C8", '=IFERROR(TRIM(DATA_SUITES!A2),"")', align="right")
put(ws, "B9", "MRI leasable SF (sum of suites)")
put(ws, "C9", "=SUM(DATA_SUITES!C2:C200)", fmt="#,##0")
put(ws, "B10", "Billing SF minus MRI leasable SF")
put(ws, "C10", "=C5-C9", fmt="#,##0;(#,##0);0")
put(ws, "B11", "Difference as % of MRI leasable SF")
put(ws, "C11", "=IF(C9=0,0,C10/C9)", fmt="0.0%;(0.0%);0.0%")
put(ws, "B12", "Billing SF check", f_bold)
put(ws, "C12", '=IF(C5=C9,"MATCH","DIFFERENT")', f_bold, align="right")
put(ws, "E12", "Highlights whenever the billing SF differs from MRI's leasable total.", f_note)
ws.conditional_formatting.add("C5", FormulaRule(formula=["$C$5<>$C$9"], fill=AMBER))
ws.conditional_formatting.add("C10:C12", FormulaRule(formula=["$C$5<>$C$9"], fill=AMBER))
put(ws, "B14", "Budget period")
put(ws, "C14", '=TEXT(DATE(C3,1,1),"mm/dd/yyyy")&" - "&TEXT(DATE(C3,12,31),"mm/dd/yyyy")', align="right")
put(ws, "B15", "Current-year actuals through month (1-12)", f_bold)
put(ws, "C15", ACTUALS_THRU, f_bold, "0", inp=True, align="right")
put(ws, "E15", "Months of the year before the budget year that have actuals. OTHER_REVENUE uses actuals through this month, budget after it.", f_note)
dv_am = DataValidation(type="whole", operator="between", formula1="1", formula2="12", allow_blank=False)
ws.add_data_validation(dv_am)
dv_am.add("C15")
put(ws, "B17", "BUDGET SCHEDULES (built by macro)", f_bold)
ws["B17"].border = Border(bottom=thin)
ws["C17"].border = Border(bottom=thin)
put(ws, "B18", "Schedule status")
put(ws, "C18", "=DETAIL!$B$3", f_bold, align="right")
put(ws, "B19", "Schedules last built")
put(ws, "C19", "=DETAIL!$B$4", align="right")
put(ws, "E18", "Rebuild after changing MRI data, a charge code on a tenant line, or any account.", f_note)
put(ws, "E19", "Click the Build Schedules button, or press Alt+F8 and run BuildAllSchedules.", f_note)
ws.row_dimensions[20].height = 22
ws.row_dimensions[21].height = 22
ws.conditional_formatting.add("C18", FormulaRule(formula=['LEFT($C$18,2)<>"OK"'], fill=AMBER))
put(ws, "B23", "LEASING COST ACCOUNTS (GL numbers used on REVENUE cost rows)", f_bold)
ws["B23"].border = Border(bottom=thin)
ws["C23"].border = Border(bottom=thin)
LC = [("Tenant improvements", 1365000), ("Commissions payable", 2105000), ("Legal - leasing", 5820400),
      ("Architectural", 5860000), ("Permits - space alterations", 5390000)]
for i, (lbl, acct) in enumerate(LC):
    r = 24 + i
    put(ws, f"B{r}", lbl)
    put(ws, f"C{r}", acct, fmt="0", inp=True, align="right")
    put(ws, f"E{r}", f'=IFERROR(INDEX(COA_NAME,MATCH(C{r},COA_ACCTNO,0)),"not in chart of accounts")', f_note)
put(ws, "B30", "LEASE OPTIONS (leases that expire during the budget year)", f_bold)
ws["B30"].border = Border(bottom=thin)
ws["C30"].border = Border(bottom=thin)
put(ws, "B31", "Expired lease handling", f_bold)
put(ws, "C31", "Terminate", f_bold, inp=True, align="right")
put(ws, "E31", "Terminate = charges stop at lease expiry (MRI). MTM = keep every charge at its last amount to December. Renew = last amount x (1 + renewal %) to December.", f_note)
dv_opt = DataValidation(type="list", formula1='"Terminate,MTM,Renew"', allow_blank=False)
ws.add_data_validation(dv_opt)
dv_opt.add("C31")
put(ws, "B32", "Renewal increase % (one-time step)", f_bold)
put(ws, "C32", 0.03, f_bold, "0.0%", inp=True, align="right")
put(ws, "E32", "Applies to the Renew option. Each lease can override both settings on its REVENUE block header (light-blue cells).", f_note)
dv_pct = DataValidation(type="decimal", operator="between", formula1="0", formula2="1", allow_blank=False)
ws.add_data_validation(dv_pct)
dv_pct.add("C32")
put(ws, "B33", "Not carried forward (GL accounts)")
for _r in (33, 34, 35):
    put(ws, f"C{_r}", 4150000 if _r == 33 else None, fmt="0", inp=True, align="right")
put(ws, "E33", "Charges on these accounts (e.g. rent concessions) are NOT extended under MTM or Renew. Up to 3 accounts.", f_note)
put(ws, "E34", '=IF(C33="","",IFERROR(INDEX(COA_NAME,MATCH(C33,COA_ACCTNO,0)),"not in chart of accounts"))', f_note)
name("LeaseOption", "INPUT!$C$31")
name("RenewPct", "INPUT!$C$32")
name("OPT_EXCL", "INPUT!$C$33:$C$35")
def cash_default():
    """MRI operating cash (CASH BALANCES group, excluding security-deposit accounts): starting balance + actual activity through the actuals-through month."""
    py = BUDGET_YEAR - 1
    am = list(act.columns[3:15])
    a_py = act[act["Year"] == py]
    tot_ = 0.0
    for r_ in coa[coa["FMT_GROUP"] == "CASH BALANCES"].itertuples():
        if "security" in str(r_.NAME).lower():
            continue
        s_ = float(pd.to_numeric(bal[bal["acctnum"] == r_.ACCTNUM]["starting_balance"]).sum())
        x_ = a_py[a_py["acctnum"] == r_.ACCTNUM]
        a_ = float(sum(pd.to_numeric(x_[c]).sum() for c in am[:ACTUALS_THRU]))
        tot_ += s_ + a_
    return round(tot_, 2)


put(ws, "B37", "CASH BALANCE (DETAIL: projected cash balance)", f_bold)
ws["B37"].border = Border(bottom=thin)
ws["C37"].border = Border(bottom=thin)
put(ws, "B38", "Cash at beginning of budget year", f_bold)
put(ws, "C38", cash_default(), f_bold, "#,##0.00", inp=True, align="right")
put(ws, "E38", "Expected cash on hand at 12/31 of the prior year, including the maintenance reserve below; DETAIL's PROJECTED CASH BALANCE starts here. Default = MRI operating cash through the actuals-through month (next line) - replace it with your expected year-end cash.", f_note)
put(ws, "B39", "  MRI operating cash through actuals-through month (read-only)")
put(ws, "C39", "=SUM(EXPENSE_ACCOUNTS!$AA$2:$AA$2001)", f_norm, "#,##0.00", align="right")
put(ws, "E39", "CASH BALANCES group bank accounts (security-deposit accounts excluded): MRI starting balance + actual activity through the actuals-through month.", f_note)
name("StartCash", "INPUT!$C$38")
def _py_sum(df_, acct_, cols, thru=None):
    x_ = df_[(df_["Year"] == BUDGET_YEAR - 1) & (df_["acctnum"] == acct_)]
    cc = cols[:thru] if thru else cols
    return float(sum(pd.to_numeric(x_[c]).sum() for c in cc))


def reserve_default():
    """Maintenance reserve balance (CASH BALANCES accounts named *reserve*): MRI starting balance + actual activity through the actuals-through month."""
    am = list(act.columns[3:15])
    tot_ = 0.0
    for r_ in coa[coa["FMT_GROUP"] == "CASH BALANCES"].itertuples():
        if "reserve" in str(r_.NAME).lower():
            tot_ += float(pd.to_numeric(bal[bal["acctnum"] == r_.ACCTNUM]["starting_balance"]).sum()) + _py_sum(act, r_.ACCTNUM, am, ACTUALS_THRU)
    return round(tot_, 2)


def dist_default():
    """Partner distributions per quarter: net prior-year activity on the partner capital accounts (DISTRIBUTIONS group, 3xxxxxx, not current P/L) through the actuals-through month / quarters elapsed."""
    am = list(act.columns[3:15])
    tot_ = 0.0
    for r_ in coa[(coa["FMT_GROUP"] == "DISTRIBUTIONS/(CONTRIBUTIONS)")].itertuples():
        if str(int(r_.ACCTNO)).startswith("3") and int(r_.ACCTNO) != 3999000:
            tot_ += _py_sum(act, r_.ACCTNUM, am, ACTUALS_THRU)
    q_ = max(1, -(-ACTUALS_THRU // 3))
    return max(0.0, round(tot_ / q_, 0))


def payroll_defaults():
    """Payroll account = HVAC & PAYROLL group account with 'payroll' in its name and the most prior-year activity; annual budget default = its prior-year budget total (else prior-year forecast)."""
    am, bm = list(act.columns[3:15]), list(bud.columns[4:16])
    best, bt = None, 0.0
    for r_ in coa[(coa["FMT_GROUP"] == "HVAC & PAYROLL") & (coa["TYPE"] == "EXPENSE")].itertuples():
        if "payroll" in str(r_.NAME).lower():
            t_ = abs(_py_sum(act, r_.ACCTNUM, am)) + abs(float(sum(pd.to_numeric(bud[(bud["Year"] == BUDGET_YEAR - 1) & (bud["acctnum"] == r_.ACCTNUM)][c]).sum() for c in bm)))
            if t_ > bt:
                best, bt = r_, t_
    if best is None:
        return None, 0.0, ""
    b_ = float(sum(pd.to_numeric(bud[(bud["Year"] == BUDGET_YEAR - 1) & (bud["acctnum"] == best.ACCTNUM)][c]).sum() for c in bm))
    f_ = _py_sum(act, best.ACCTNUM, am, ACTUALS_THRU) + float(sum(pd.to_numeric(bud[(bud["Year"] == BUDGET_YEAR - 1) & (bud["acctnum"] == best.ACCTNUM)][c]).sum() for c in bm[ACTUALS_THRU:]))
    return int(best.ACCTNO), round(b_ if b_ > 0 else f_, 0), str(best.NAME)


RES_DEF, DIST_DEF = reserve_default(), dist_default()
PAYROLL_ACCT, PAYROLL_ANNUAL, PAYROLL_NAME = payroll_defaults()
print(f"INPUT defaults: reserve {RES_DEF:,.2f} | distributions {DIST_DEF:,.0f}/quarter | payroll acct {PAYROLL_ACCT} {PAYROLL_NAME!r} annual {PAYROLL_ANNUAL:,.0f}")
dv_m12 = DataValidation(type="whole", operator="between", formula1="1", formula2="12", allow_blank=True, showErrorMessage=True,
                        errorTitle="Month", error="Enter a month number from 1 to 12.")
dv_amt0 = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True, showErrorMessage=True,
                         errorTitle="Amount", error="Enter a positive amount (or leave blank).")
dv_yn2 = DataValidation(type="list", formula1='"Yes,No"', allow_blank=False, showErrorMessage=True)
for v_ in (dv_m12, dv_amt0, dv_yn2):
    ws.add_data_validation(v_)
ws.column_dimensions["D"].width = max(ws.column_dimensions["D"].width or 0, 11)
# -- maintenance reserve (rows 41-46)
put(ws, "B41", "MAINTENANCE RESERVE", f_bold)
ws["B41"].border = Border(bottom=thin)
ws["C41"].border = Border(bottom=thin)
put(ws, "D41", "Month", f_bold, align="center")
ws["D41"].border = Border(bottom=thin)
put(ws, "B42", "Reserve balance at beginning of budget year", f_bold)
put(ws, "C42", RES_DEF, f_bold, "#,##0.00", inp=True, align="right")
put(ws, "E42", "Included in the starting cash above. DETAIL shows operating cash (starting cash less this reserve) and the reserve separately. Default = MRI maintenance-reserve account.", f_note)
for q_ in range(4):
    r_ = 43 + q_
    put(ws, f"B{r_}", f"Addition to reserve - quarter {q_ + 1}")
    put(ws, f"C{r_}", 0, f_bold, "#,##0", inp=True, align="right")
    put(ws, f"D{r_}", 3 * (q_ + 1), f_bold, "0", inp=True, align="center")
dv_amt0.add("C43:C46")
dv_m12.add("D43:D46")
put(ws, "E43", "Cash moved from operating cash into the reserve. Month = the month it is posted (default: last month of the quarter).", f_note)
name("ResStart", "INPUT!$C$42")
name("ResAmt", "INPUT!$C$43:$C$46")
name("ResMonth", "INPUT!$D$43:$D$46")
# -- partner distributions (rows 48-53)
put(ws, "B48", "PARTNER DISTRIBUTIONS", f_bold)
ws["B48"].border = Border(bottom=thin)
ws["C48"].border = Border(bottom=thin)
put(ws, "D48", "Month", f_bold, align="center")
ws["D48"].border = Border(bottom=thin)
put(ws, "B49", "Partner distributions anticipated?", f_bold)
put(ws, "C49", "Yes" if DIST_DEF > 0 else "No", f_bold, inp=True, align="right")
dv_yn2.add("C49")
put(ws, "E49", "No = ignore the four amounts below. Default amounts = prior-year distributions to date divided by the quarters elapsed.", f_note)
for q_ in range(4):
    r_ = 50 + q_
    put(ws, f"B{r_}", f"Distribution - quarter {q_ + 1}")
    put(ws, f"C{r_}", DIST_DEF, f_bold, "#,##0", inp=True, align="right")
    put(ws, f"D{r_}", 3 * q_ + 1, f_bold, "0", inp=True, align="center")
dv_amt0.add("C50:C53")
dv_m12.add("D50:D53")
put(ws, "E50", "Month = the month it is paid (default: first month of the quarter, as MRI shows prior-year distributions).", f_note)
name("DistOn", "INPUT!$C$49")
name("DistAmt", "INPUT!$C$50:$C$53")
name("DistMonth", "INPUT!$D$50:$D$53")
# -- payroll (rows 55-60)
put(ws, "B55", ("PAYROLL (GL %d %s)" % (PAYROLL_ACCT, PAYROLL_NAME)) if PAYROLL_ACCT else "PAYROLL (no payroll account found in the HVAC & PAYROLL group)", f_bold)
ws["B55"].border = Border(bottom=thin)
for c_ in "CDEF":
    ws[f"{c_}55"].border = Border(bottom=thin)
put(ws, "B56", "Annual payroll budget", f_bold)
put(ws, "C56", PAYROLL_ANNUAL if PAYROLL_ACCT else None, f_bold, "#,##0", inp=True, align="right")
put(ws, "E56", "Each month = annual amount / 52 x pay weeks (4 or 5) + summer help. Default = the prior-year budget for the account. Shown on the payroll line of the account's EXP_ block.", f_note)
dv_amt0.add("C56")
put(ws, "B57", "Months with 5 pay weeks (month numbers)", f_bold)
put(ws, "B58", "  (four months)")
for ref_, m_ in zip(("C57", "D57", "C58", "D58"), (3, 6, 8, 11)):
    put(ws, ref_, m_, f_bold, "0", inp=True, align="center")
dv_m12.add("C57:D58")
put(ws, "E57", "Default = the prior year's months (Mar, Jun, Aug, Nov) - update for the budget-year pay calendar.", f_note)
put(ws, "B59", "Summer help - monthly amount", f_bold)
put(ws, "C59", 0, f_bold, "#,##0", inp=True, align="right")
dv_amt0.add("C59")
put(ws, "B60", "Summer help months (first, last)", f_bold)
put(ws, "C60", 6, f_bold, "0", inp=True, align="center")
put(ws, "D60", 8, f_bold, "0", inp=True, align="center")
dv_m12.add("C60:D60")
put(ws, "B61", "  Pay weeks budgeted (should be 52)")
put(ws, "C61", "=48+COUNT(C57:D58)", f_norm, "0", align="right")
put(ws, "E61", '=IF(C61=52,"","NOT 52 WEEKS - the annual amount will not be fully budgeted: list exactly four 5-week months")', f_bold)
name("PayAnnual", "INPUT!$C$56")
name("PayWeekMonths", "INPUT!$C$57:$D$58")
name("PaySummerAmt", "INPUT!$C$59")
name("PaySummerFrom", "INPUT!$C$60")
name("PaySummerTo", "INPUT!$D$60")
put(ws, "B66", "LEGEND", f_bold)
ws["B66"].border = Border(bottom=thin)
ws["C66"].border = Border(bottom=thin)
put(ws, "B67", "Light blue background = user input")
ws["C67"].fill = INPUT_FILL
put(ws, "B68", "White background = read-only (locked)")
put(ws, "B69", "Amber = flag / adjusted amount")
ws["C69"].fill = PatternFill("solid", fgColor="FFD966")
put(ws, "B70", "Light green = forecast from a lease option (not from MRI)")
ws["C70"].fill = PatternFill("solid", fgColor="E2EFDA")
put(ws, "B71", "Tabs: orange = input | purple = budget schedules | blue = MRI reference data")
put(ws, "B72", "Budget schedules (DETAIL, PROFORMA, QTRLY SUM) are presentation output and use their own color scheme.", f_note)
ws.protection.sheet = True
ws.protection.formatColumns = False
ws.protection.formatRows = False

# ------------------------------------------------------------------ TENANT_SCHEDULE
ws = ws_ts
hdr = ["LEASE ID", "SUITE", "TENANT", "LEASE START", "LEASE EXPIRES", "VACATE", "CHARGE CODE", "EFFECTIVE", "AMOUNT", "ENDS",
       "GL ACCOUNT", "KEY LEASE", "KEY LEASE|CODE", "FIRST LEASE+CODE", "CODE RANK IN LEASE", "KEY LEASE|RANK",
       "FIRST OF LEASE", "SORT KEY", "LEASE RANK"]
for j, h in enumerate(hdr, 1):
    c = ws.cell(1, j, h)
    c.font = f_bold
    c.border = Border(bottom=thin)
    c.alignment = Alignment(wrap_text=True, vertical="bottom")
TSM0 = 20  # column T = Jan
for m in range(12):
    c = ws.cell(1, TSM0 + m, f"=DATE(BudgetYear,{m + 1},1)")
    c.number_format = "mmm-yy"
    c.font = f_bold
    c.border = Border(bottom=thin)
c = ws.cell(1, TSM0 + 12, "TOTAL")
c.font = f_bold
c.border = Border(bottom=thin)
FBC = CL(TSM0 + 13)   # forecast base column (AG)
c = ws.cell(1, TSM0 + 13, "FORECAST BASE (last amount at expiry)")
c.font = f_bold
c.border = Border(bottom=thin)
c.alignment = Alignment(wrap_text=True, vertical="bottom")
D = "DATA_TENANT_CHARGES!"
for r in range(2, 402):
    ws[f"A{r}"] = f'=IF({D}B{r}="","",{D}B{r}&"")'
    ws[f"B{r}"] = f'=IF($A{r}="","",{D}C{r}&"")'
    ws[f"C{r}"] = f'=IF($A{r}="","",{D}D{r}&"")'
    ws[f"D{r}"] = f'=IF($A{r}="","",{D}E{r})'
    ws[f"E{r}"] = f'=IF($A{r}="","",{D}F{r})'
    ws[f"F{r}"] = f'=IF($A{r}="","",IF({D}G{r}="","",{D}G{r}))'
    ws[f"G{r}"] = f'=IF($A{r}="","",{D}H{r}&"")'
    ws[f"H{r}"] = f'=IF($A{r}="","",{D}I{r})'
    ws[f"I{r}"] = f'=IF($A{r}="","",N({D}J{r}))'
    ws[f"J{r}"] = f'=IF($A{r}="","",IF({D}L{r}="","",{D}L{r}))'
    ws[f"K{r}"] = f'=IF($A{r}="","",IFERROR(VALUE(MID(INDEX(CC_ACCT,MATCH($G{r},CC_CODE,0)),3,7)),0))'
    ws[f"L{r}"] = f'=IF($A{r}="","","L"&$A{r})'
    ws[f"M{r}"] = f'=IF($A{r}="","",$L{r}&"|"&$G{r})'
    ws[f"N{r}"] = f'=IF($A{r}="","",IF(COUNTIF($M$2:$M{r},$M{r})=1,1,0))'
    ws[f"O{r}"] = f'=IF($N{r}=1,COUNTIFS($L$2:$L{r},$L{r},$N$2:$N{r},1),"")'
    ws[f"P{r}"] = f'=IF($N{r}=1,$L{r}&"|"&$O{r},"")'
    ws[f"Q{r}"] = f'=IF($A{r}="","",IF(COUNTIF($L$2:$L{r},$L{r})=1,1,0))'
    ws[f"R{r}"] = f'=IF($Q{r}=1,IFERROR(VALUE($B{r}),90000)+ROW()/100000,"")'
    ws[f"S{r}"] = f'=IF($Q{r}=1,1+COUNTIFS($R$2:$R$401,"<"&$R{r}),"")'
    for m in range(12):
        col = CL(TSM0 + m)
        ws[f"{col}{r}"] = (f'=IF($A{r}="",0,IF(AND($H{r}<={col}$1,'
                           f'$H{r}=_xlfn.MAXIFS($H$2:$H$401,$M$2:$M$401,$M{r},$H$2:$H$401,"<="&{col}$1),'
                           f'OR($J{r}="",$J{r}>={col}$1),$E{r}>={col}$1,OR($F{r}="",$F{r}>={col}$1)),$I{r},0))')
        ws[f"{col}{r}"].number_format = NUM2
    ws[f"{CL(TSM0 + 12)}{r}"] = f"=SUM({CL(TSM0)}{r}:{CL(TSM0 + 11)}{r})"
    ws[f"{CL(TSM0 + 12)}{r}"].number_format = NUM2
    _lm = f'DATE(YEAR($E{r}),MONTH($E{r}),1)'
    ws[f"{FBC}{r}"] = (f'=IF(OR($A{r}="",NOT(ISNUMBER($E{r})),$F{r}<>""),0,'
                       f'IF(AND($H{r}<={_lm},$H{r}=_xlfn.MAXIFS($H$2:$H$401,$M$2:$M$401,$M{r},$H$2:$H$401,"<="&{_lm}),'
                       f'OR($J{r}="",$J{r}>={_lm}),COUNTIF(OPT_EXCL,$K{r})=0),$I{r},0))')
    ws[f"{FBC}{r}"].number_format = NUM2
    for col in "DEFHJ":
        ws[f"{col}{r}"].number_format = DATEF
    ws[f"I{r}"].number_format = NUM2
    ws[f"K{r}"].number_format = "0"
    ws[f"R{r}"].number_format = "0.00000"
    for col in range(1, TSM0 + 14):
        ws.cell(r, col).font = f_norm
for col, w in zip("ABCDEFGHIJKLMNOPQRS", (10, 8, 30, 12, 12, 11, 8, 12, 11, 12, 11, 10, 15, 9, 8, 14, 9, 10, 7)):
    ws.column_dimensions[col].width = w
for m in range(14):
    ws.column_dimensions[CL(TSM0 + m)].width = 12
ws.row_dimensions[1].height = 42
ws.freeze_panes = "D2"
ws.cell(1, TSM0 + 15, "Calculation sheet (read-only). One row per MRI charge row. A charge applies in a month when it is the latest effective row "
                      "for that lease+code, has not ended, and the lease has not expired or vacated. FORECAST BASE = the amount in effect in the lease's expiry month (used by the Lease Options on INPUT; 0 if the lease has a vacate date, the charge had ended, or its account is excluded on INPUT). Capacity 400 rows.").font = f_note
ws.protection.sheet = True
ws.protection.formatColumns = False
ws.protection.formatRows = False

# ------------------------------------------------------------------ REVENUE
ws = ws_rev
ws.sheet_view.showGridLines = False
COSTS = [("Tenant Improvements", 24), ("Commissions Payable", 25), ("Legal", 26), ("Architecture", 27),
         ("Permits - Space Alterations", 28)]
G0 = 7
N_ADD = 3


def lease_block_sizes(te_df):
    """Charge-code count per lease, in TENANT_SCHEDULE lease-rank order.
    Rank rule (TENANT_SCHEDULE col R/S): suite number (90000 if not numeric) + first row/100000.
    Each in-place block gets exactly one MRI line per distinct charge code of its lease."""
    keys = []
    seen = {}
    for i, row in enumerate(te_df.itertuples(index=False)):
        if row.leasid in seen:
            continue
        seen[row.leasid] = i
        try:
            sv = float(str(row.suitid).strip())
        except ValueError:
            sv = 90000.0
        keys.append((sv + (i + 2) / 100000.0, row.leasid))
    keys.sort()
    return [int(te_df.loc[te_df.leasid == lid, "inccat"].nunique()) for _, lid in keys]


BLOCK_LINES = lease_block_sizes(te)
N_INPLACE = len(BLOCK_LINES)
print("in-place blocks:", N_INPLACE, "| MRI lines per block:", BLOCK_LINES)
LAST = 1500
for col, w in zip("ABCDEF", (8, 9, 38, 11, 11, 11)):
    ws.column_dimensions[col].width = w
for m in range(12):
    ws.column_dimensions[CL(G0 + m)].width = 11
ws.column_dimensions["S"].width = 13
ws.column_dimensions["T"].width = 11
ws.column_dimensions["U"].width = 7
put(ws, "A1", '=BuildingName&" - PROJECTED REVENUES"', f_title)
put(ws, "A2", '="BUDGET FOR THE YEAR "&BudgetYear', f_bold)
put(ws, "A3", "Each lease lists its MRI charge codes (one line per code). Every MRI line has an adjustment row: budget amount = MRI + adjustment. "
              "Type numbers or formulas (e.g. =G11*0.1) in the light-blue rows. Light green = forecast from the Lease Options on INPUT (not from MRI).", f_sub)
for col, t in {"B": "Code", "C": "Tenant / Charge", "D": "GL Acct", "E": "Source"}.items():
    put(ws, f"{col}5", t, f_bold)
for m in range(12):
    put(ws, f"{CL(G0 + m)}5", f"=DATE(BudgetYear,{m + 1},1)", f_bold, "mmm-yy", align="right")
put(ws, "S5", "Total", f_bold, align="right")
put(ws, "T5", "Flag", f_bold)
put(ws, "U5", "Block", f_bold)
for col in range(1, 22):
    ws.cell(5, col).border = Border(bottom=thin)
put(ws, "C6", "TOTAL - ALL LEASES (Total Rent for Lease)", f_bold)
for m in range(12):
    col = CL(G0 + m)
    put(ws, f"{col}6", f'=SUMIFS({col}$8:{col}${LAST},$D$8:$D${LAST},"TOTAL")', f_bold, NUM0)
put(ws, "S6", "=SUM(G6:R6)", f_bold, NUM0)


def line_cells(r, kind):
    for m in range(12):
        c = ws.cell(r, G0 + m)
        c.number_format = NUM0
        c.font = f_norm
        c.border = Border(bottom=hair)
        if kind == "input":
            c.fill = INPUT_FILL
            c.protection = UNLOCK
    s = ws.cell(r, 19, f"=SUM(G{r}:R{r})")
    s.number_format = NUM0
    s.font = f_norm
    s.border = Border(bottom=hair)


def flag(r, text):
    ws.cell(r, 20, f'=IF(SUMPRODUCT(ABS(G{r}:R{r}))>0,"{text}","")').font = f_bold


def desc_f(r):
    return f'=IF($B{r}="","",IFERROR(INDEX(CC_DESC,MATCH($B{r},CC_CODE,0)),"(code not in MRI charge list)"))'


def acct_f(r):
    return f'=IF($B{r}="","",IFERROR(VALUE(MID(INDEX(CC_ACCT,MATCH($B{r},CC_CODE,0)),3,7)),"CODE?"))'


def adj_row(ra, rm):
    put(ws, f"C{ra}", "      adjustment (+/-)", f_note)
    put(ws, f"D{ra}", f'=IF($B{rm}="","",$D{rm})')
    put(ws, f"E{ra}", "ADJ", f_note)
    line_cells(ra, "input")
    flag(ra, "ADJUSTED")


def add_rows(ra0):
    for i in range(N_ADD):
        r = ra0 + i
        put(ws, f"B{r}", None, inp=True)
        ws[f"B{r}"].number_format = "@"
        put(ws, f"C{r}", None, f_note)
        ws[f"C{r}"] = desc_f(r)
        put(ws, f"D{r}", acct_f(r))
        put(ws, f"E{r}", "ADD", f_note)
        line_cells(r, "input")
        flag(r, "ADDED")
        ADD_ROWS.append(r)


def total_and_costs(rt, first, last):
    put(ws, f"C{rt}", "Total Rent for Lease", f_bold)
    put(ws, f"D{rt}", "TOTAL", f_note)
    for m in range(12):
        col = CL(G0 + m)
        put(ws, f"{col}{rt}", f"=SUM({col}{first}:{col}{last})", f_bold, NUM0)
    put(ws, f"S{rt}", f"=SUM(G{rt}:R{rt})", f_bold, NUM0)
    for col in range(1, 22):
        ws.cell(rt, col).border = Border(top=thin)
    for i, (nm, irow) in enumerate(COSTS):
        r = rt + 1 + i
        put(ws, f"C{r}", nm)
        put(ws, f"D{r}", f"=INPUT!$C${irow}", fmt="0")
        put(ws, f"E{r}", "COST", f_note)
        line_cells(r, "input")
    return rt + 1 + len(COSTS) + 1


ADD_ROWS = []
CODE_CELLS = []
GREEN_RULES = []
OPT_DV = DataValidation(type="list", formula1='"Terminate,MTM,Renew"', allow_blank=True)
PCT_DV = DataValidation(type="decimal", operator="between", formula1="0", formula2="1", allow_blank=True)
ws.add_data_validation(OPT_DV)
ws.add_data_validation(PCT_DV)


def write_inplace(r0, k, n_lines):
    v = r0 + 1
    for col, t in {"A": "Suite", "B": "Lease ID", "C": "Tenant", "D": "Sq Ft", "E": "Lease start", "F": "Expires"}.items():
        put(ws, f"{col}{r0}", t, f_lbl)
    put(ws, f"U{v}", k, f_lbl)
    put(ws, f"B{v}", f'=IFERROR(INDEX(TS_LEASE,MATCH($U{v},TS_RANK,0)),"")', f_bold)
    put(ws, f"A{v}", f'=IF($B{v}="","",INDEX(TS_SUITE,MATCH($B{v},TS_LEASE,0)))', f_bold)
    put(ws, f"C{v}", f'=IF($B{v}="","",INDEX(TS_TENANT,MATCH($B{v},TS_LEASE,0)))', f_bold)
    put(ws, f"D{v}", f'=IF($B{v}="","",IFERROR(INDEX(SUIT_GLA,MATCH($A{v}&"",SUIT_ID,0)),"n/a"))', f_bold, "#,##0")
    put(ws, f"E{v}", f'=IF($B{v}="","",INDEX(TS_START,MATCH($B{v},TS_LEASE,0)))', f_bold, DATEF)
    put(ws, f"F{v}", f'=IF($B{v}="","",INDEX(TS_EXPIRY,MATCH($B{v},TS_LEASE,0)))', f_bold, DATEF)
    for col, t in {"G": "If expired:", "H": "Renewal %", "I": "Applied", "J": "Applied %"}.items():
        put(ws, f"{col}{r0}", t, f_lbl, align="right")
    put(ws, f"G{v}", None, f_bold, inp=True, align="right")
    put(ws, f"H{v}", None, f_bold, "0.0%", inp=True, align="right")
    put(ws, f"I{v}", f'=IF($B{v}="","",IF(SUMPRODUCT((DEAL_LEASEID=$B{v}&"")*(DEAL_TYPE="Renewal"))>0,"Terminate",IF(OR($G{v}="MTM",$G{v}="Renew"),$G{v},IF($G{v}="Terminate","Terminate",IF(OR(LeaseOption="MTM",LeaseOption="Renew"),LeaseOption,"Terminate")))))', f_bold, align="right")
    put(ws, f"J{v}", f'=IF($B{v}="","",IF(ISNUMBER($H{v}),$H{v},RenewPct))', f_bold, "0.0%", align="right")
    OPT_DV.add(f"G{v}")
    PCT_DV.add(f"H{v}")
    for col in range(1, 22):
        ws.cell(v, col).border = Border(bottom=thin)
    base = r0 + 2
    for i in range(n_lines):
        rm, ra = base + 2 * i, base + 2 * i + 1
        put(ws, f"F{rm}", i + 1, f_lbl)
        put(ws, f"B{rm}", f'=IF($B${v}="","",IFERROR(INDEX(TS_CODE,MATCH("L"&$B${v}&"|"&$F{rm},TS_KEYRANK,0)),""))')
        put(ws, f"C{rm}", desc_f(rm))
        put(ws, f"D{rm}", acct_f(rm))
        put(ws, f"E{rm}", "MRI", f_note)
        for m in range(12):
            col = CL(G0 + m)
            ws[f"{col}{rm}"] = (f'=IF($B{rm}="",0,SUMIFS(TENANT_SCHEDULE!{CL(TSM0 + m)}$2:{CL(TSM0 + m)}$401,'
                                f'TENANT_SCHEDULE!$M$2:$M$401,"L"&$B${v}&"|"&$B{rm})'
                                f'+IF(OR($I${v}="Terminate",{col}$5<=$F${v}),0,'
                                f'SUMIFS(TENANT_SCHEDULE!${FBC}$2:${FBC}$401,TENANT_SCHEDULE!$M$2:$M$401,"L"&$B${v}&"|"&$B{rm})'
                                f'*IF($I${v}="Renew",1+$J${v},1)))')
        line_cells(rm, "calc")
        ws.cell(rm, 20, f'=IF(AND($I${v}<>"Terminate",SUMPRODUCT((G$5:R$5>$F${v})*(G{rm}:R{rm}<>0))>0),"FORECAST - "&IF($I${v}="Renew","RENEWAL","MTM"),"")').font = f_bold
        adj_row(ra, rm)
    GREEN_RULES.append((v, base, base + 2 * n_lines - 1))
    ra0 = base + 2 * n_lines
    # capacity flag: the lease has more distinct charge codes in MRI than this block has MRI lines
    ws[f"T{v}"] = (f'=IF($B{v}="","",IF(COUNTIFS(TENANT_SCHEDULE!$L$2:$L$401,"L"&$B{v},TENANT_SCHEDULE!$N$2:$N$401,1)'
                   f'>COUNTIF($E{base}:$E{ra0 - 1},"MRI"),"MORE CODES THAN LINES",""))')
    ws[f"T{v}"].font = f_bold
    add_rows(ra0)
    rt = ra0 + N_ADD
    for i in range(N_ADD):
        CODE_CELLS.append(f"B{ra0 + i}")
    return total_and_costs(rt, base, ra0 + N_ADD - 1)


# ------------------------------------------------------------------ LEASE_DEALS (input: one line per new lease / renewal / expansion) and DEAL_SCHEDULE (calc)
# Session 5 design (user): replaces the 15 new-tenant slots on REVENUE. One row per deal, LEASPROJ-style. The $/SF charge columns are the charge codes
# THIS building's tenants use (most-used first, max DL_MAX_CODES, rent and free-rent codes excluded) plus DL_OTHER columns where a code is picked.
DL_MAX_CODES, DL_OTHER, N_DEALS = 10, 2, 40
DL_RENT_DEFAULT, DL_FREE_DEFAULT = "00", "05"
_codes_used = (te.assign(_c=te["inccat"].astype(str).str.strip()).groupby("_c")["leasid"].nunique()
               .drop([DL_RENT_DEFAULT, DL_FREE_DEFAULT], errors="ignore"))
DL_CODES = [c for c, _ in sorted(_codes_used.items(), key=lambda kv: (-kv[1], kv[0]))][:DL_MAX_CODES]
N_DL_CODES = len(DL_CODES) + DL_OTHER
# LEASE_DEALS layout
DL_SET = {"rent": 4, "free": 5, "lap": 6, "ti": 7, "comm": 8, "emc": 9}       # settings rows (value in column E, second value in G)
DL_GRP, DL_HDR, DL_CODEROW, DL_ACCTROW = 11, 12, 13, 14
DL_FIRST = 15
DL_LAST = DL_FIRST + N_DEALS - 1
DLC = {k: CL(i) for i, k in enumerate(["no", "suite", "tenant", "type", "lease", "sf", "start", "end", "term", "ner", "rent", "incr",
                                        "f1m", "f1n", "f1p", "f2m", "f2n", "f2p"], 1)}
_c0 = 19                                                                                       # first charge-code column (S)
DL_CODECOLS = [CL(_c0 + j) for j in range(N_DL_CODES)]
_c1 = _c0 + N_DL_CODES
for _k in ["legal", "arch", "permit", "ti", "commpct", "emcpct", "check", "notes"]:
    DLC[_k] = CL(_c1)
    _c1 += 1
DL_LASTCOL = _c1 - 1
# DEAL_SCHEDULE layout: per deal one IDX row (lease month number, 0 = not active) and one row per role
DS_ROLES = ([("RENT", "INC"), ("FREE", "INC")] + [(f"C{j + 1}", "INC") for j in range(N_DL_CODES)] +
            [("LEGAL", "COST"), ("ARCH", "COST"), ("PERMIT", "COST"), ("TI", "COST"), ("COMM", "COST"), ("EMC", "COST")])
DS_COST_INPUT_ROW = {"TI": 24, "COMM": 25, "EMC": 25, "LEGAL": 26, "ARCH": 27, "PERMIT": 28}    # INPUT leasing-cost account rows
DS_FIRST = 7
DS_PER_DEAL = 1 + len(DS_ROLES)
DS_LAST = DS_FIRST + N_DEALS * DS_PER_DEAL - 1
DS_LIST_COL = "W"                                                                               # hidden list of in-place MRI lease IDs (renewal dropdown)


def code_acct_f(code_ref):
    return f'IF({code_ref}="","",IFERROR(VALUE(MID(INDEX(CC_ACCT,MATCH({code_ref}&"",CC_CODE,0)),3,7)),"CODE?"))'


def code_desc_f(code_ref):
    return f'IF({code_ref}="","",IFERROR(INDEX(CC_DESC,MATCH({code_ref}&"",CC_CODE,0)),"(code not in MRI charge list)"))'


def build_lease_deals(ws):
    ws.sheet_view.showGridLines = False
    widths = {"no": 4, "suite": 8, "tenant": 30, "type": 11, "lease": 11, "sf": 9, "start": 11, "end": 11, "term": 6, "ner": 8, "rent": 9,
              "incr": 7, "f1m": 6, "f1n": 6, "f1p": 6, "f2m": 6, "f2n": 6, "f2p": 6, "legal": 8, "arch": 8, "permit": 8, "ti": 8,
              "commpct": 8, "emcpct": 8, "check": 46, "notes": 30}
    for k, w in widths.items():
        ws.column_dimensions[DLC[k]].width = w
    for c in DL_CODECOLS:
        ws.column_dimensions[c].width = 10
    put(ws, "A1", '=BuildingName&" - LEASE DEALS (new leases, renewals, expansions)"', f_title)
    put(ws, "A2", "One line per deal. Light-blue cells are inputs. Amounts are calculated on DEAL_SCHEDULE and appear on REVENUE (section LEASE DEALS) and DETAIL. "
                  "Charges start the first month on or after the lease start date and end in the month of the lease end date; amounts are $ per square foot per year "
                  "(costs: $ per square foot, paid once). In-place MRI leases stay on REVENUE.", f_sub)
    put(ws, "B3", "SETTINGS", f_bold)
    for col in range(2, 9):
        ws.cell(3, col).border = Border(bottom=thin)
    dv_code = DataValidation(type="list", formula1="=CC_CODE", allow_blank=True, showErrorMessage=False)
    dv_m = DataValidation(type="whole", operator="between", formula1="0", formula2="24", allow_blank=False, showErrorMessage=True,
                          errorTitle="Months", error="Enter a whole number of months from 0 to 24.")
    dv_type = DataValidation(type="list", formula1='"New lease,Renewal,Expansion"', allow_blank=True)
    dv_lease = DataValidation(type="list", formula1="=DEAL_LEASE_LIST", allow_blank=True, showErrorMessage=True,
                              errorTitle="MRI lease", error="Pick the in-place MRI lease that this deal renews.")
    dv_pct = DataValidation(type="decimal", operator="between", formula1="0", formula2="1", allow_blank=True, showErrorMessage=True,
                            errorTitle="Percent", error="Enter a percentage between 0% and 100%.")
    dv_lm = DataValidation(type="whole", operator="between", formula1="1", formula2="240", allow_blank=True, showErrorMessage=True,
                           errorTitle="Lease month", error="Enter a lease month number (1 = first month of rent).")
    for v_ in (dv_code, dv_m, dv_type, dv_lease, dv_pct, dv_lm):
        ws.add_data_validation(v_)
    rows = [("rent", "Base rent charge code", DL_RENT_DEFAULT, None, "@", "The annual increase % applies to this code."),
            ("free", "Free-rent (concession) charge code", DL_FREE_DEFAULT, None, "@", "Free rent posts here as a negative amount."),
            ("lap", "Legal, architect and permits paid: months before lease start", 2, None, "0", "Paid once, in the month of the lease start less these months (legacy LEASPROJ: 2)."),
            ("ti", "Tenant improvements paid: months after lease start", 0, None, "0", "0 = the lease start month."),
            ("comm", "Broker commission: months before / after start (half each)", 3, 3, "0",
             "Commission = % x (NER, or base rent if NER is blank) x SF x lease term in years (legacy LEASPROJ: 3 / 3)."),
            ("emc", "EMC commission paid: months after lease start", 0, None, "0", "Same base as the broker commission.")]
    for k, lbl, v1, v2, fmt, note in rows:
        r_ = DL_SET[k]
        put(ws, f"B{r_}", lbl)
        put(ws, f"E{r_}", v1, f_bold, fmt, inp=True, align="right")
        if k in ("rent", "free"):
            dv_code.add(f"E{r_}")
            put(ws, f"F{r_}", f"={code_desc_f(f'$E${r_}')}", f_note)
            put(ws, f"I{r_}", f'="GL "&{code_acct_f(f"$E${r_}")}', f_note)
        else:
            dv_m.add(f"E{r_}")
        if v2 is not None:
            put(ws, f"G{r_}", v2, f_bold, fmt, inp=True, align="right")
            dv_m.add(f"G{r_}")
        put(ws, f"{'K' if k in ('rent', 'free') else 'I'}{r_}", note, f_note)
    name("DL_RENTCODE", f"LEASE_DEALS!$E${DL_SET['rent']}")
    name("DL_FREECODE", f"LEASE_DEALS!$E${DL_SET['free']}")
    # header rows
    grp = {"no": "DEAL", "rent": "BASE RENT", "f1m": "FREE RENT 1", "f2m": "FREE RENT 2", "legal": "LEASING COSTS ($/SF, paid once)", "commpct": "COMMISSIONS"}
    for k, t in grp.items():
        put(ws, f"{DLC[k]}{DL_GRP}", t, f_bold)
    put(ws, f"{DL_CODECOLS[0]}{DL_GRP}", "OTHER CHARGES ($/SF/yr) - charge codes used by this building's tenants", f_bold)
    hdr = {"no": "#", "suite": "Suite", "tenant": "Tenant / description", "type": "Type", "lease": "MRI lease renewed", "sf": "Sq Ft",
           "start": "Lease start", "end": "Lease end", "term": "Term (yrs)", "ner": "NER $/SF (opt.)", "rent": "Rent $/SF/yr", "incr": "Annual incr %",
           "f1m": "From lease month", "f1n": "Months", "f1p": "% of rent (blank = 100%)", "f2m": "From lease month", "f2n": "Months", "f2p": "% of rent (blank = 100%)",
           "legal": "Legal", "arch": "Architect", "permit": "Permits", "ti": "TI", "commpct": "Broker %", "emcpct": "EMC %",
           "check": "Line status", "notes": "Notes"}
    for k, t in hdr.items():
        c = put(ws, f"{DLC[k]}{DL_HDR}", t, f_bold, align="center")
        c.alignment = Alignment(horizontal="center", vertical="bottom", wrap_text=True)
    ws.row_dimensions[DL_HDR].height = 54
    put(ws, f"A{DL_CODEROW}", "Charge code", f_lbl)
    put(ws, f"A{DL_ACCTROW}", "GL account", f_lbl)
    put(ws, f"{DLC['rent']}{DL_CODEROW}", "=DL_RENTCODE", f_lbl, align="center")
    put(ws, f"{DLC['rent']}{DL_ACCTROW}", f"={code_acct_f('DL_RENTCODE')}", f_lbl, "0", align="center")
    for k in ("f1m", "f2m"):
        put(ws, f"{DLC[k]}{DL_CODEROW}", "=DL_FREECODE", f_lbl, align="center")
        put(ws, f"{DLC[k]}{DL_ACCTROW}", f"={code_acct_f('DL_FREECODE')}", f_lbl, "0", align="center")
    for j, c in enumerate(DL_CODECOLS):
        fixed = j < len(DL_CODES)
        put(ws, f"{c}{DL_CODEROW}", DL_CODES[j] if fixed else None, f_bold, "@", inp=not fixed, align="center")
        if not fixed:
            dv_code.add(f"{c}{DL_CODEROW}")
        cr = f"{c}${DL_CODEROW}"
        h = put(ws, f"{c}{DL_HDR}", f'=IF({cr}="","(pick a code below)",{cr}&" "&{code_desc_f(cr)})', f_bold, align="center")
        h.alignment = Alignment(horizontal="center", vertical="bottom", wrap_text=True)
        put(ws, f"{c}{DL_ACCTROW}", f"={code_acct_f(cr)}", f_lbl, "0", align="center")
    for k, irow in (("legal", 26), ("arch", 27), ("permit", 28), ("ti", 24), ("commpct", 25), ("emcpct", 25)):
        put(ws, f"{DLC[k]}{DL_ACCTROW}", f"=INPUT!$C${irow}", f_lbl, "0", align="center")
    for col in range(1, DL_LASTCOL + 1):
        ws.cell(DL_ACCTROW, col).border = Border(bottom=thin)
    # deal lines
    fm = {"sf": "#,##0", "start": DATEF, "end": DATEF, "ner": "#,##0.00", "rent": "#,##0.00", "incr": "0.0%", "f1m": "0", "f1n": "0", "f1p": "0%",
          "f2m": "0", "f2n": "0", "f2p": "0%", "legal": "#,##0.00", "arch": "#,##0.00", "permit": "#,##0.00", "ti": "#,##0.00",
          "commpct": "0.00%", "emcpct": "0.00%"}
    for i in range(N_DEALS):
        r = DL_FIRST + i
        put(ws, f"A{r}", i + 1, f_lbl, align="center")
        for k in ["suite", "tenant", "type", "lease", "sf", "start", "end", "ner", "rent", "incr", "f1m", "f1n", "f1p", "f2m", "f2n", "f2p",
                  "legal", "arch", "permit", "ti", "commpct", "emcpct", "notes"]:
            put(ws, f"{DLC[k]}{r}", None, f_norm, fm.get(k), inp=True)
        for c in DL_CODECOLS:
            put(ws, f"{c}{r}", None, f_norm, "#,##0.00", inp=True)
        ws[f"{DLC['suite']}{r}"].number_format = "@"
        ws[f"{DLC['lease']}{r}"].number_format = "@"
        dv_type.add(f"{DLC['type']}{r}")
        dv_lease.add(f"{DLC['lease']}{r}")
        for k in ("incr", "f1p", "f2p", "commpct", "emcpct"):
            dv_pct.add(f"{DLC[k]}{r}")
        for k in ("f1m", "f2m"):
            dv_lm.add(f"{DLC[k]}{r}")
        g, h, f_, ty, ls = (f"${DLC[k]}{r}" for k in ("start", "end", "sf", "type", "lease"))
        put(ws, f"{DLC['term']}{r}", f'=IF(AND(ISNUMBER({g}),ISNUMBER({h})),ROUND(({h}-{g})/365,2),"")', f_norm, "0.00")
        used = (f"COUNTA(${DLC['suite']}{r}:${DLC['end']}{r},${DLC['ner']}{r}:${DLC['f2p']}{r},${DL_CODECOLS[0]}{r}:${DLC['emcpct']}{r})>0")   # not Term (formula)
        exp_ = f"IFERROR(INDEX(TS_EXPIRY,MATCH({ls},TS_LEASE,0)),\"\")"
        ys = "BudgetYear"
        ws[f"{DLC['check']}{r}"] = (
            f'=IF(NOT({used}),"",'
            f'IF(OR(N({f_})<=0,NOT(ISNUMBER({g})),NOT(ISNUMBER({h}))),"CHECK - enter Sq Ft, lease start and lease end",'
            f'IF({h}<{g},"CHECK - lease end is before lease start",'
            f'IF(AND({ty}="Renewal",{ls}=""),"CHECK - pick the MRI lease this renews",'
            f'IF(AND({ls}<>"",COUNTIF(DEAL_LEASE_LIST,{ls})=0),"CHECK - MRI lease not found",'
            f'IF(AND({ty}="Renewal",ISNUMBER({exp_}),{g}<={exp_}),"NOTE - starts before the MRI lease expires: both are budgeted for the overlap",'
            f'IF(OR({h}<DATE({ys},1,1),{g}>DATE({ys},12,31)),"NOTE - no months in the budget year (costs before the start may still fall in it)",'
            f'IF({ty}="","NOTE - no type chosen",""))))))))')
        ws[f"{DLC['check']}{r}"].font = f_bold
    ws.conditional_formatting.add(f"{DLC['check']}{DL_FIRST}:{DLC['check']}{DL_LAST}",
                                  FormulaRule(formula=[f'LEFT(${DLC["check"]}{DL_FIRST},5)="CHECK"'], fill=AMBER))
    name("DEAL_TYPE", f"LEASE_DEALS!${DLC['type']}${DL_FIRST}:${DLC['type']}${DL_LAST}")
    name("DEAL_LEASEID", f"LEASE_DEALS!${DLC['lease']}${DL_FIRST}:${DLC['lease']}${DL_LAST}")
    name("DEAL_CHECK", f"LEASE_DEALS!${DLC['check']}${DL_FIRST}:${DLC['check']}${DL_LAST}")
    ws.freeze_panes = f"D{DL_FIRST}"
    ws.protection.sheet = True
    ws.protection.formatColumns = False
    ws.protection.formatRows = False


def build_deal_schedule(ws):
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDEF", (6, 8, 44, 11, 7, 2)):
        ws.column_dimensions[col].width = w
    for m in range(12):
        ws.column_dimensions[CL(7 + m)].width = 11
    ws.column_dimensions["S"].width = 13
    put(ws, "A1", '=BuildingName&" - DEAL SCHEDULE (calculation, read-only)"', f_title)
    put(ws, "A2", "Monthly amounts for each LEASE_DEALS line. Row IDX = lease month number (1 = first month of rent; 0 = not active). REVENUE sums these rows by role.", f_sub)
    for col, t in {"A": "Deal", "B": "Role", "C": "Line", "D": "GL Acct", "E": "Tag"}.items():
        put(ws, f"{col}5", t, f_bold)
    for m in range(12):
        put(ws, f"{CL(7 + m)}5", f"=DATE(BudgetYear,{m + 1},1)", f_bold, "mmm-yy", align="right")
    put(ws, "S5", "Total", f_bold, align="right")
    for col in range(1, 20):
        ws.cell(5, col).border = Border(bottom=thin)
    LD = "LEASE_DEALS!"
    sv = lambda k: f"{LD}$E${DL_SET[k]}"
    for i in range(N_DEALS):
        d = DL_FIRST + i
        q = lambda k: f"{LD}${DLC[k]}${d}"
        r0 = DS_FIRST + i * DS_PER_DEAL
        put(ws, f"A{r0}", i + 1, f_bold)
        put(ws, f"B{r0}", "IDX", f_note)
        put(ws, f"C{r0}", f'=IF({q("tenant")}="","(deal line {i + 1} - empty)",{q("tenant")}&IF({q("suite")}="",""," (suite "&{q("suite")}&")"))', f_bold)
        put(ws, f"E{r0}", "IDX", f_note)
        g, h, f_ = q("start"), q("end"), q("sf")
        for m in range(12):
            col = CL(7 + m)
            rs = f"DATE(YEAR({g}),MONTH({g})+IF(DAY({g})=1,0,1),1)"
            n = f"((YEAR({col}$5)-YEAR({rs}))*12+MONTH({col}$5)-MONTH({rs})+1)"
            ws[f"{col}{r0}"] = (f'=IF(OR(NOT(ISNUMBER({g})),NOT(ISNUMBER({h})),N({f_})<=0),0,IF(OR({col}$5>{h},{n}<1),0,{n}))')
            ws[f"{col}{r0}"].number_format = "0"
            ws[f"{col}{r0}"].font = f_note
        for j, (role, tag) in enumerate(DS_ROLES, 1):
            r = r0 + j
            put(ws, f"A{r}", i + 1, f_note)
            put(ws, f"B{r}", role, f_note)
            put(ws, f"E{r}", tag, f_note)
            if role == "RENT":
                put(ws, f"C{r}", "  base rent", f_norm)
                put(ws, f"D{r}", f"={code_acct_f('DL_RENTCODE')}", f_norm, "0")
            elif role == "FREE":
                put(ws, f"C{r}", "  free rent (concession)", f_norm)
                put(ws, f"D{r}", f"={code_acct_f('DL_FREECODE')}", f_norm, "0")
            elif (role[0] == "C" and role[1:].isdigit()):
                cc = f"{LD}{DL_CODECOLS[int(role[1:]) - 1]}${DL_CODEROW}"
                put(ws, f"C{r}", f'=IF({cc}="","  (unused charge column)","  "&{cc}&" "&{code_desc_f(cc)})', f_norm)
                put(ws, f"D{r}", f"={code_acct_f(cc)}", f_norm, "0")
            else:
                lbl = {"LEGAL": "legal", "ARCH": "architect", "PERMIT": "permits", "TI": "tenant improvements", "COMM": "broker commission", "EMC": "EMC commission"}[role]
                put(ws, f"C{r}", f"  {lbl}", f_norm)
                put(ws, f"D{r}", f"=INPUT!$C${DS_COST_INPUT_ROW[role]}", f_norm, "0")
            for m in range(12):
                col = CL(7 + m)
                ix = f"{col}${r0}"
                st_m = lambda k_off: f"DATE(YEAR({g}),MONTH({g}){k_off},1)"
                comm_base = f'IF(N({q("ner")})>0,{q("ner")},N({q("rent")}))*N({f_})*({h}-{g})/365'
                if role == "RENT":
                    fml = f'=IF({ix}=0,0,N({f_})*N({q("rent")})/12*(1+N({q("incr")}))^INT(({ix}-1)/12))'
                elif role == "FREE":
                    per = lambda a, b, p: (f'IF(AND(N({q(a)})>0,{ix}>=N({q(a)}),{ix}<N({q(a)})+N({q(b)})),IF({q(p)}="",1,{q(p)}),0)')
                    fml = f"=-{col}{r0 + 1}*({per('f1m', 'f1n', 'f1p')}+{per('f2m', 'f2n', 'f2p')})"
                elif (role[0] == "C" and role[1:].isdigit()):
                    cc_col = DL_CODECOLS[int(role[1:]) - 1]
                    fml = f'=IF({ix}=0,0,N({f_})*N({LD}${cc_col}${d})/12)'
                else:
                    ok = f"AND(ISNUMBER({g}),ISNUMBER({h}))"
                    if role in ("LEGAL", "ARCH", "PERMIT"):
                        rate = {"LEGAL": "legal", "ARCH": "arch", "PERMIT": "permit"}[role]
                        fml = f'=IF(AND({ok},{col}$5={st_m("-N(" + sv("lap") + ")")}),N({f_})*N({q(rate)}),0)'
                    elif role == "TI":
                        fml = f'=IF(AND({ok},{col}$5={st_m("+N(" + sv("ti") + ")")}),N({f_})*N({q("ti")}),0)'
                    elif role == "COMM":
                        fml = (f'=IF({ok},N({q("commpct")})*{comm_base}/2*(({col}$5={st_m("-N(" + sv("comm") + ")")})'
                               f'+({col}$5={st_m("+N(" + f"{LD}$G${DL_SET['comm']}" + ")")})),0)')
                    else:
                        fml = f'=IF(AND({ok},{col}$5={st_m("+N(" + sv("emc") + ")")}),N({q("emcpct")})*{comm_base},0)'
                ws[f"{col}{r}"] = fml
                ws[f"{col}{r}"].number_format = NUM0
            put(ws, f"S{r}", f"=SUM(G{r}:R{r})", f_norm, NUM0)
        for col in range(1, 20):
            ws.cell(r0, col).border = Border(top=thin)
    # hidden list of in-place MRI lease IDs (renewal dropdown)
    lids = list(dict.fromkeys(str(x).strip() for x in te["leasid"]))
    put(ws, f"{DS_LIST_COL}6", "MRI leases", f_lbl)
    for k_, lid in enumerate(lids):
        put(ws, f"{DS_LIST_COL}{7 + k_}", lid, f_norm, "@")
    name("DEAL_LEASE_LIST", f"DEAL_SCHEDULE!${DS_LIST_COL}$7:${DS_LIST_COL}${7 + max(len(lids), 1) - 1}")
    ws.column_dimensions[DS_LIST_COL].hidden = True
    ws.freeze_panes = "G6"
    ws.protection.sheet = True
    ws.protection.formatColumns = False
    ws.protection.formatRows = False


def write_deal_section(r):
    """REVENUE: read-only LEASE DEALS section, one line per role summed over every deal line on DEAL_SCHEDULE (DETAIL reads these like any REVENUE line)."""
    put(ws, f"A{r}", "LEASE DEALS - new leases, renewals and expansions (enter them on LEASE_DEALS; read-only here)", f_bold)
    r += 1
    first = r
    rng = lambda col: f"DEAL_SCHEDULE!{col}${DS_FIRST}:{col}${DS_LAST}"
    for role, tag in DS_ROLES:
        if tag != "INC":
            continue
        if role == "RENT":
            code = "=DL_RENTCODE"
        elif role == "FREE":
            code = "=DL_FREECODE"
        else:
            code = f'=IF(LEASE_DEALS!{DL_CODECOLS[int(role[1:]) - 1]}${DL_CODEROW}="","",LEASE_DEALS!{DL_CODECOLS[int(role[1:]) - 1]}${DL_CODEROW}&"")'
        put(ws, f"B{r}", code)
        put(ws, f"C{r}", desc_f(r))
        put(ws, f"D{r}", acct_f(r))
        put(ws, f"E{r}", "CALC", f_note)
        put(ws, f"U{r}", role, f_note)
        ws[f"U{r}"].number_format = "@"
        for m in range(12):
            col = CL(G0 + m)
            ws[f"{col}{r}"] = f'=SUMIFS({rng(col)},{rng("$B")},"{role}")'
        line_cells(r, "calc")
        r += 1
    last = r - 1
    rt = r
    put(ws, f"C{rt}", "Total lease-deal revenue", f_bold)
    put(ws, f"D{rt}", "TOTAL", f_note)
    for m in range(12):
        col = CL(G0 + m)
        put(ws, f"{col}{rt}", f"=SUM({col}{first}:{col}{last})", f_bold, NUM0)
    put(ws, f"S{rt}", f"=SUM(G{rt}:R{rt})", f_bold, NUM0)
    for col in range(1, 22):
        ws.cell(rt, col).border = Border(top=thin)
    r = rt + 1
    for role, tag in DS_ROLES:
        if tag != "COST":
            continue
        lbl = {"LEGAL": "Legal (lease deals)", "ARCH": "Architecture (lease deals)", "PERMIT": "Permits - Space Alterations (lease deals)",
               "TI": "Tenant Improvements (lease deals)", "COMM": "Commissions - broker (lease deals)", "EMC": "Commissions - EMC (lease deals)"}[role]
        put(ws, f"C{r}", lbl)
        put(ws, f"D{r}", f"=INPUT!$C${DS_COST_INPUT_ROW[role]}", fmt="0")
        put(ws, f"E{r}", "COST", f_note)
        put(ws, f"U{r}", role, f_note)
        ws[f"U{r}"].number_format = "@"
        for m in range(12):
            col = CL(G0 + m)
            ws[f"{col}{r}"] = f'=SUMIFS({rng(col)},{rng("$B")},"{role}")'
        line_cells(r, "calc")
        r += 1
    return r + 1


r = 8
put(ws, f"A{r}", "IN-PLACE LEASES (from MRI)", f_bold)
r += 1
for k in range(1, N_INPLACE + 1):
    r = write_inplace(r, k, BLOCK_LINES[k - 1])
r = write_deal_section(r)
REV_END = r
build_lease_deals(ws_ld)
build_deal_schedule(ws_ds)
print(f"LEASE_DEALS: {N_DEALS} deal lines, charge columns {DL_CODES} + {DL_OTHER} pick-a-code | DEAL_SCHEDULE rows {DS_FIRST}-{DS_LAST}")
assert REV_END < LAST, REV_END
ws.freeze_panes = "D6"
dvl = DataValidation(type="list", formula1="=CC_CODE", allow_blank=True, showErrorMessage=False)
ws.add_data_validation(dvl)
for ref in CODE_CELLS:
    dvl.add(ref)
GREEN = PatternFill("solid", bgColor="E2EFDA", fgColor="E2EFDA")
for _v, _b0, _b1 in GREEN_RULES:
    ws.conditional_formatting.add(f"G{_b0}:R{_b1}", FormulaRule(
        formula=[f'AND($E{_b0}="MRI",$I${_v}<>"Terminate",G$5>$F${_v},G{_b0}<>0)'], fill=GREEN))
ws.conditional_formatting.add(f"G9:S{REV_END}", FormulaRule(formula=['AND(OR($E9="ADJ",$E9="ADD"),G9<>0)'], fill=AMBER))
# ---- DETAIL status per account line: is the account on the MRI Financial Format and does it have a DETAIL section? (V) and the signed amount DETAIL uses (W)
put(ws, "V5", "Detail status", f_bold)
put(ws, "W5", "Detail amount", f_bold, align="right")
ws.column_dimensions["V"].width = 36
ws.column_dimensions["W"].width = 13
for col in (22, 23):
    ws.cell(5, col).border = Border(bottom=thin)
for _r in range(9, REV_END + 1):
    if ws.cell(_r, 5).value in ("MRI", "ADJ", "ADD", "COST", "CALC"):
        _grp = f'IFERROR(INDEX(COA_FMT,MATCH($D{_r},COA_ACCTNO,0))&"","")'
        _ty = f'IFERROR(INDEX(COA_TYPE,MATCH($D{_r},COA_ACCTNO,0))&"","")'
        ws[f"V{_r}"] = (f'=IF(AND(ISNUMBER($D{_r}),$D{_r}>=1000000),IF(COUNTIF(INPUT!$C$24:$C$28,$D{_r})>0,"",IF({_grp}="","IGNORED - NOT ON FINANCIAL FORMAT",'
                        f'IF({_ty}="INCOME","",IF(COUNTIF(BUILT_GROUPS,{_grp})=0,"IGNORED - SECTION NOT BUILT YET","")))),"")')
        ws[f"W{_r}"] = (f'=IF(AND(ISNUMBER($D{_r}),$D{_r}>=1000000),IF($V{_r}<>"",0,IF(OR($E{_r}="COST",{_ty}="INCOME",{_grp}="INCOME"),$S{_r},-$S{_r})),0)')
        ws[f"V{_r}"].font = f_note
        ws[f"W{_r}"].font = f_note
        ws[f"W{_r}"].number_format = NUM0
ws.conditional_formatting.add(f"D9:E{REV_END}", FormulaRule(formula=['LEFT($V9,7)="IGNORED"'], fill=AMBER))
ws.protection.sheet = True
ws.protection.formatColumns = False
ws.protection.formatRows = False

# ------------------------------------------------------------------ INCOME_ACCOUNTS (calc: which INCOME-group accounts had 2026 activity)
ws = ws_ia
for j, h in enumerate(["ACCTNO", "NAME", "FMT_GROUP", "PRIOR-YR ACTUAL (ABS)", "PRIOR-YR BUDGET (ABS)", "QUALIFIES", "RANK", "INCOME LIST #", "DROPDOWN TEXT (User blocks)", "INCOME ACCT"], 1):
    c = ws.cell(1, j, h)
    c.font = f_bold
    c.border = Border(bottom=thin)
    c.alignment = Alignment(wrap_text=True, vertical="bottom")
DA, DB = "DATA_ACTUALS!", "DATA_BUDGET_PY!"
for r in range(2, 2002):
    ws[f"A{r}"] = f'=IF(DATA_COA!C{r}="","",DATA_COA!C{r})'
    ws[f"B{r}"] = f'=IF($A{r}="","",DATA_COA!D{r})'
    ws[f"C{r}"] = f'=IF($A{r}="","",IF(DATA_COA!G{r}="","",DATA_COA!G{r}))'
    ws[f"D{r}"] = (f'=IF($C{r}<>"INCOME",0,SUMPRODUCT(({DA}$C$2:$C$2000="MR"&$A{r})*({DA}$B$2:$B$2000=BudgetYear-1)*ABS({DA}$D$2:$O$2000)))')
    ws[f"E{r}"] = (f'=IF($C{r}<>"INCOME",0,SUMPRODUCT(({DB}$D$2:$D$2000="MR"&$A{r})*({DB}$B$2:$B$2000=BudgetYear-1)*ABS({DB}$E$2:$P$2000)))')
    ws[f"F{r}"] = f'=IF(AND($C{r}="INCOME",$D{r}+$E{r}>0.004),1,0)'
    ws[f"G{r}"] = f'=IF($F{r}=1,COUNTIF($F$2:$F{r},1),"")'
    ws[f"H{r}"] = f'=IF($C{r}="INCOME",COUNTIF($C$2:$C{r},"INCOME"),"")'
    ws[f"J{r}"] = f'=IFERROR(INDEX($A$2:$A$2001,MATCH(ROW()-1,$H$2:$H$2001,0)),"")'
    ws[f"I{r}"] = f'=IF($J{r}="","",TEXT($J{r},"0")&"  "&INDEX($B$2:$B$2001,MATCH($J{r},$A$2:$A$2001,0)))'
    for col in "ABCDEFGHIJ":
        ws[f"{col}{r}"].font = f_norm
    ws[f"D{r}"].number_format = NUM2
    ws[f"E{r}"].number_format = NUM2
for col, w in zip("ABCDEFG", (10, 34, 24, 16, 16, 10, 8)):
    ws.column_dimensions[col].width = w
ws.row_dimensions[1].height = 30
ws.freeze_panes = "A2"
ws["I1"] = ("Calculation sheet (read-only). Lists the chart-of-accounts rows whose Financial Format group is INCOME and that have "
            "prior-year (budget year - 1) actual or budget amounts. OTHER_REVENUE builds one block per qualifying account, in account order.")
ws["I1"].font = f_note
ws.protection.sheet = True
ws.protection.formatColumns = False
ws.protection.formatRows = False

# ------------------------------------------------------------------ OTHER_REVENUE (input: non-tenant revenue, one block per INCOME account)
ws = ws_or
ws.sheet_view.showGridLines = False
N_USER = 5
OR_FIRST, OR_ROWS = 9, 10
OR_LAST = 600


def qualifying_income_accounts():
    """Same rule as INCOME_ACCOUNTS col F: INCOME-group account with prior-year (budget year - 1) actual or budget amounts."""
    py = BUDGET_YEAR - 1
    am, bm = list(act.columns[3:15]), list(bud.columns[4:16])
    a = act[act["Year"] == py]
    b = bud[bud["Year"] == py]
    tot = {}
    for df_, cols in ((a, am), (b, bm)):
        for acct_, v in zip(df_["acctnum"], pd.DataFrame({c: pd.to_numeric(df_[c]) for c in cols}).abs().sum(axis=1)):
            tot[acct_] = tot.get(acct_, 0) + v
    inc = coa[coa["FMT_GROUP"] == "INCOME"]
    return [r.ACCTNO for r in inc.itertuples() if tot.get(r.ACCTNUM, 0) > 0.004]


N_AUTO = len(qualifying_income_accounts())
N_OR = N_AUTO + N_USER
N_INCOME_LIST = int((coa["FMT_GROUP"] == "INCOME").sum())
name("INC_LIST_NUM", f"INCOME_ACCOUNTS!$J$2:$J${1 + N_INCOME_LIST}")
name("INC_LIST_TXT", f"INCOME_ACCOUNTS!$I$2:$I${1 + N_INCOME_LIST}")
assert OR_FIRST + OR_ROWS * N_OR + 2 < OR_LAST, "OTHER_REVENUE needs more than %d rows" % OR_LAST
print("OTHER_REVENUE: %d automatic account blocks + %d user blocks | income accounts in dropdown: %d" % (N_AUTO, N_USER, N_INCOME_LIST))
user_dv = DataValidation(type="list", formula1="=INC_LIST_TXT", allow_blank=True, showErrorMessage=True,
                         errorTitle="Income account", error="Pick an account from the list (income-group accounts only).")
ws.add_data_validation(user_dv)
for col, w in zip("ABCDEF", (10, 2, 46, 11, 8, 2)):
    ws.column_dimensions[col].width = w
for m in range(12):
    ws.column_dimensions[CL(7 + m)].width = 11
ws.column_dimensions["S"].width = 13
ws.column_dimensions["T"].width = 10
ws.column_dimensions["U"].width = 7
put(ws, "A1", '=BuildingName&" - OTHER (NON-TENANT) REVENUE"', f_title)
put(ws, "A2", '="BUDGET FOR THE YEAR "&BudgetYear', f_bold)
put(ws, "A3", "One block per INCOME-group GL account with prior-year actual or budget amounts, then five USER blocks for any other income account. Enter the budget-year amounts (revenue as a positive number) "
              "on the three light-blue lines and describe each one. Tenant-fed accounts (rent, recoveries) show what REVENUE already counts - leave their lines blank.", f_sub)
put(ws, "A4", '="Prior-year forecast = actuals Jan-"&TEXT(DATE(BudgetYear-1,ActualsThru,1),"mmm")&" plus budget for the remaining months (set on INPUT)."', f_sub)
for col, t in {"A": "GL Acct", "C": "Account / Line", "D": "GL", "E": "Source"}.items():
    put(ws, f"{col}5", t, f_bold)
for m in range(12):
    put(ws, f"{CL(7 + m)}5", f"=DATE(BudgetYear,{m + 1},1)", f_bold, "mmm-yy", align="right")
put(ws, "S5", "Total", f_bold, align="right")
put(ws, "T5", "$/SF", f_bold, align="right")
put(ws, "U5", "Block", f_bold)
for col in range(1, 22):
    ws.cell(5, col).border = Border(bottom=thin)
BLK_END = OR_FIRST + OR_ROWS * N_OR - 1
for rr, lbl, tag in ((6, "PRIOR-YEAR FORECAST - ALL ACCOUNTS", "FCST"), (7, "BUDGET-YEAR INPUT - ALL ACCOUNTS (this sheet)", "INPUT")):
    put(ws, f"C{rr}", lbl, f_bold)
    for m in range(12):
        col = CL(7 + m)
        put(ws, f"{col}{rr}", f'=SUMIFS({col}${OR_FIRST}:{col}${OR_LAST},$E${OR_FIRST}:$E${OR_LAST},"{tag}")', f_bold, NUM0)
    put(ws, f"S{rr}", f"=SUM(G{rr}:R{rr})", f_bold, NUM0)
    put(ws, f"T{rr}", f"=IF(BillingSF=0,0,S{rr}/BillingSF)", f_bold, NUM2)
NUM0_BZ = '#,##0;(#,##0);'
NUM2_BZ = '#,##0.00;(#,##0.00);'
white_fill = PatternFill("solid", bgColor="FFFFFF", fgColor="FFFFFF")


def or_calc_cells(r, kind=None, bold=False):
    for m in range(12):
        c = ws.cell(r, 7 + m)
        c.number_format = NUM0
        c.font = f_bold if bold else f_norm
        c.border = Border(bottom=hair)


for k in range(1, N_OR + 1):
    r0 = OR_FIRST + OR_ROWS * (k - 1)
    ra, rb, rf, rr_, i1, i2, i3, rt = r0 + 1, r0 + 2, r0 + 3, r0 + 4, r0 + 5, r0 + 6, r0 + 7, r0 + 8
    A = f"$A${r0}"
    is_user = k > N_AUTO
    if not is_user:
        put(ws, f"U{r0}", k, f_lbl)
        put(ws, f"A{r0}", f'=IFERROR(INDEX(IA_ACCTNO,MATCH($U{r0},IA_RANK,0)),"")', f_bold, "0000\\-000", align="left")
        put(ws, f"C{r0}", f'=IF({A}="","",IFERROR(INDEX(COA_NAME,MATCH({A},COA_ACCTNO,0)),"(not in chart of accounts)"))', f_bold)
    else:
        if k == N_AUTO + 1:
            put(ws, f"A{r0 - 1}", "USER BLOCKS - budget any other income account: pick it in the light-blue cell of a block, then enter amounts on its three light-blue lines.", f_bold)
        put(ws, f"U{r0}", f"U{k - N_AUTO}", f_lbl)
        put(ws, f"A{r0}", f'=IFERROR(VALUE(LEFT($C{r0},7)),"")', f_bold, "0000\\-000", align="left")
        put(ws, f"C{r0}", None, f_bold, inp=True)
        user_dv.add(f"C{r0}")
        put(ws, f"G{r0}", f'=IF($C{r0}="","<- select an income account in the light-blue cell","")', f_note)
        ws[f"T{r0}"] = (f'=IF($C{r0}="","",IF(ISNA(MATCH($A{r0},INC_LIST_NUM,0)),"NOT AN INCOME ACCOUNT",'
                        f'IF(COUNTIF($A${OR_FIRST}:$A${OR_LAST},$A{r0})>1,"DUPLICATE ACCOUNT","")))')
        ws[f"T{r0}"].font = f_bold
    for col in range(1, 22):
        ws.cell(r0, col).border = Border(bottom=thin)
    put(ws, f"C{ra}", f'=IF({A}="","","  "&(BudgetYear-1)&" actual (Jan-"&TEXT(DATE(BudgetYear-1,ActualsThru,1),"mmm")&")")')
    put(ws, f"C{rb}", f'=IF({A}="","",IF(ActualsThru>=12,"  "&(BudgetYear-1)&" budget (none - full year actual)","  "&(BudgetYear-1)&" budget ("&TEXT(DATE(BudgetYear-1,ActualsThru+1,1),"mmm")&"-Dec)"))')
    put(ws, f"C{rf}", f'=IF({A}="","","  "&(BudgetYear-1)&" forecast")', f_bold)
    put(ws, f"C{rr_}", f'=IF({A}="","","  "&BudgetYear&" on REVENUE tab (read-only)")', f_note)
    put(ws, f"C{rt}", f'=IF({A}="","","  "&BudgetYear&" total")', f_bold)
    for rowi, tag in ((ra, "ACT"), (rb, "BUD"), (rf, "FCST"), (rr_, "REV")):
        put(ws, f"E{rowi}", tag, f_note)
    for m in range(12):
        col = CL(7 + m)
        ac = CL(4 + m)       # DATA_ACTUALS month column (D = Jan)
        bc = CL(5 + m)       # DATA_BUDGET_PY month column (E = Jan)
        ws[f"{col}{ra}"] = (f'=IF({A}="","",IF(MONTH({col}$5)<=ActualsThru,-SUMIFS({DA}{ac}$2:{ac}$2000,{DA}$C$2:$C$2000,"MR"&{A},{DA}$B$2:$B$2000,BudgetYear-1),0))')
        ws[f"{col}{rb}"] = (f'=IF({A}="","",IF(MONTH({col}$5)>ActualsThru,-SUMIFS({DB}{bc}$2:{bc}$2000,{DB}$D$2:$D$2000,"MR"&{A},{DB}$B$2:$B$2000,BudgetYear-1),0))')
        ws[f"{col}{rf}"] = f'=IF({A}="","",{col}{ra}+{col}{rb})'
        ws[f"{col}{rr_}"] = f'=IF({A}="","",SUMIFS(REVENUE!{col}$8:{col}$1500,REVENUE!$D$8:$D$1500,{A}))'
        ws[f"{col}{rt}"] = f'=IF({A}="","",{col}{rr_}+SUM({col}{i1}:{col}{i3}))'
    for rowi, bold in ((ra, False), (rb, False), (rf, True), (rr_, False), (rt, True)):
        or_calc_cells(rowi, bold=bold)
        put(ws, f"S{rowi}", f'=IF({A}="","",SUM(G{rowi}:R{rowi}))', f_bold if bold else f_norm, NUM0)
        put(ws, f"T{rowi}", f'=IF({A}="","",IF(BillingSF=0,0,S{rowi}/BillingSF))', f_bold if bold else f_norm, NUM2)
    for col in range(1, 22):
        ws.cell(rt, col).border = Border(top=thin)
    for rowi in (i1, i2, i3):
        put(ws, f"C{rowi}", None, inp=True)
        put(ws, f"D{rowi}", f'=IF({A}="","",{A})', fmt="0")
        put(ws, f"E{rowi}", "INPUT", f_note)
        for m in range(12):
            c = ws.cell(rowi, 7 + m)
            c.fill = INPUT_FILL
            c.protection = UNLOCK
            c.number_format = NUM0
            c.font = f_norm
            c.border = Border(bottom=hair)
        put(ws, f"S{rowi}", f"=SUM(G{rowi}:R{rowi})", fmt=NUM0_BZ)
        put(ws, f"T{rowi}", f"=IF(BillingSF=0,0,S{rowi}/BillingSF)", fmt=NUM2_BZ)
    # unused block: hide the input styling so nothing invites typing where no account exists
    if not is_user:
        ws.conditional_formatting.add(f"C{i1}:R{i3}", FormulaRule(formula=[f'$A${r0}=""'], fill=white_fill))
ws.freeze_panes = "D6"
ws.protection.sheet = True
ws.protection.formatColumns = False
ws.protection.formatRows = False

# ------------------------------------------------------------------ EXPENSE_ACCOUNTS (calc) + the six EXP_ input sheets
EXP_FIRST, EXP_LAST = 9, 1500
# REPAIRS_CAPITAL layout (used by the EXP_ memo lines and by the REPAIRS_CAPITAL sheet)
RC_REP_SLOTS = 40
RC_CAP_SLOTS = 25
CAP_ACCT_DEFAULT = 1345000
RC_M0 = 28                                          # column AB = first "budget by month" column (override cells are P:AA, total AN, check AO, notes AP)
RC_REP_HDR, RC_REP_FIRST = 11, 12
RC_REP_LAST = RC_REP_FIRST + RC_REP_SLOTS - 1       # 51
RC_REP_TOT = RC_REP_LAST + 1                        # 52
RC_CAP_BAND = RC_REP_TOT + 3                        # 55
RC_CAP_ACCT = RC_CAP_BAND + 1                       # 56
RC_CAP_HDR = RC_CAP_ACCT + 2                        # 58
RC_CAP_FIRST = RC_CAP_HDR + 1                       # 59
RC_CAP_LAST = RC_CAP_FIRST + RC_CAP_SLOTS - 1       # 83
RC_CAP_TOT = RC_CAP_LAST + 1                        # 84

ws = ws_ea
for j, h in enumerate(["ACCTNO", "NAME", "FMT_GROUP", "TYPE", "IN EXP SECTION", "PRIOR-YR ACTUAL (ABS)", "PRIOR-YR BUDGET (ABS)", "VENDOR ACTIVITY (ABS)",
                       "QUALIFIES", "BLOCKS ON EXP SHEETS", "NO BLOCK"], 1):
    c = ws.cell(1, j, h)
    c.font = f_bold
    c.border = Border(bottom=thin)
    c.alignment = Alignment(wrap_text=True, vertical="bottom")
DV = "DATA_VENDOR!"
_blocks = "+".join(f"COUNTIF({n}!$A${EXP_FIRST}:$A${EXP_LAST},$A{{r}})" for n in [x[2] for x in EXP_SECTIONS] + ["TAXES", "INSURANCE", "PROF_FEES", "DEBT", "OTHER_PROF_FEES"])
for r in range(2, 2002):
    ws[f"A{r}"] = f'=IF(DATA_COA!C{r}="","",DATA_COA!C{r})'
    ws[f"B{r}"] = f'=IF($A{r}="","",DATA_COA!D{r})'
    ws[f"C{r}"] = f'=IF($A{r}="","",IF(DATA_COA!G{r}="","",DATA_COA!G{r}))'
    ws[f"D{r}"] = f'=IF($A{r}="","",DATA_COA!E{r})'
    ws[f"E{r}"] = (f'=IF(OR(AND($D{r}="EXPENSE",COUNTIF($N$2:$N$10,$C{r})>0),AND($C{r}<>"",$C{r}=$N$11),'
                   f'AND($D{r}="EXPENSE",$C{r}<>"",$C{r}=$N$12,COUNTIF(INPUT!$C$24:$C$28,$A{r})=0)),1,0)')
    ws[f"F{r}"] = f'=IF($E{r}=1,SUMPRODUCT(({DA}$C$2:$C$2000="MR"&$A{r})*({DA}$B$2:$B$2000=BudgetYear-1)*ABS({DA}$D$2:$O$2000)),0)'
    ws[f"G{r}"] = f'=IF($E{r}=1,SUMPRODUCT(({DB}$D$2:$D$2000="MR"&$A{r})*({DB}$B$2:$B$2000=BudgetYear-1)*ABS({DB}$E$2:$P$2000)),0)'
    ws[f"H{r}"] = f'=IF($E{r}=1,SUMPRODUCT(({DV}$D$2:$D$3000="MR"&$A{r})*ABS({DV}$E$2:$P$3000)),0)'
    ws[f"I{r}"] = f'=IF(AND($E{r}=1,$F{r}+$G{r}+$H{r}>0.004),1,0)'
    ws[f"J{r}"] = f'=IF($I{r}=1,{_blocks.format(r=r)},0)'
    ws[f"K{r}"] = f'=IF(AND($I{r}=1,$J{r}=0),1,0)'
    for col in "ABCDEFGHIJK":
        ws[f"{col}{r}"].font = f_norm
    for col in "FGH":
        ws[f"{col}{r}"].number_format = NUM2
ws["N1"] = "EXPENSE SECTIONS (Financial Format groups)"
ws["N1"].font = f_bold
for i_, grp_ in enumerate([x[1] for x in EXP_SECTIONS] + ["TAXES", "INSURANCE", "PROFESSIONAL FEES/G&A"]):
    ws[f"N{2 + i_}"] = grp_
    ws[f"N{2 + i_}"].font = f_norm
ws["N11"] = "DEBT SERVICE"           # any account type (interest expense + mortgage payable)
ws["N11"].font = f_norm
ws["N12"] = "OTHER PROFESSIONAL FEES"  # EXPENSE type; accounts listed on INPUT as leasing costs stay in IMPRVMTS & COMMISSIONS
ws["N12"].font = f_norm
OPF_GROUP = "OTHER PROFESSIONAL FEES"
N_OPF_LIST = int(((coa["FMT_GROUP"] == OPF_GROUP) & (coa["TYPE"] == "EXPENSE")).sum())
for col_, h_ in (("W", "OPF LIST RANK"), ("X", "OPF LIST #"), ("Y", "DROPDOWN TEXT (OTHER_PROF_FEES User blocks)")):
    ws[f"{col_}1"] = h_
    ws[f"{col_}1"].font = f_bold
for r_ in range(2, 2002):
    ws[f"W{r_}"] = f'=IF(AND($C{r_}=$N$12,$D{r_}="EXPENSE",COUNTIF(INPUT!$C$24:$C$28,$A{r_})=0),COUNT($W$1:$W{r_ - 1})+1,"")'
    ws[f"X{r_}"] = f'=IFERROR(INDEX($A$2:$A$2001,MATCH(ROW()-1,$W$2:$W$2001,0)),"")'
    ws[f"Y{r_}"] = f'=IF($X{r_}="","",TEXT($X{r_},"0")&"  "&INDEX($B$2:$B$2001,MATCH($X{r_},$A$2:$A$2001,0)))'
    for col_ in "WXY":
        ws[f"{col_}{r_}"].font = f_norm
ws.column_dimensions["Y"].width = 40
for col_, h_ in (("Z", "CASH ACCT (excl. security deposits)"), ("AA", "CASH BALANCE (start + actual activity thru)")):
    ws[f"{col_}1"] = h_
    ws[f"{col_}1"].font = f_bold
for r_ in range(2, 2002):
    ws[f"Z{r_}"] = f'=IF(AND($C{r_}="CASH BALANCES",ISERROR(SEARCH("security",$B{r_}))),1,0)'
    ws[f"AA{r_}"] = (f'=IF($Z{r_}=1,SUMIFS(DATA_BALANCES!$D$2:$D$500,DATA_BALANCES!$C$2:$C$500,"MR"&$A{r_},DATA_BALANCES!$B$2:$B$500,BudgetYear-1)'
                     f'+SUMPRODUCT((DATA_ACTUALS!$C$2:$C$2000="MR"&$A{r_})*(DATA_ACTUALS!$B$2:$B$2000=BudgetYear-1)*(DATA_ACTUALS!$D$2:$O$2000)*((COLUMN(DATA_ACTUALS!$D$2:$O$2000)-3)<=ActualsThru)),0)')
    ws[f"Z{r_}"].font = f_norm
    ws[f"AA{r_}"].font = f_norm
    ws[f"AA{r_}"].number_format = NUM2
N_RC_LIST = int(((coa["TYPE"] == "EXPENSE") & coa["FMT_GROUP"].isin([x[1] for x in EXP_SECTIONS])).sum())
for col_, h_ in (("AB", "REPAIR LIST RANK"), ("AC", "REPAIR LIST #"), ("AD", "DROPDOWN TEXT (REPAIRS_CAPITAL major repairs: EXPENSE accounts in the six expense sections)")):
    ws[f"{col_}1"] = h_
    ws[f"{col_}1"].font = f_bold
for r_ in range(2, 2002):
    ws[f"AB{r_}"] = f'=IF(AND($D{r_}="EXPENSE",COUNTIF($N$2:$N$7,$C{r_})>0),COUNT($AB$1:$AB{r_ - 1})+1,"")'
    ws[f"AC{r_}"] = f'=IFERROR(INDEX($A$2:$A$2001,MATCH(ROW()-1,$AB$2:$AB$2001,0)),"")'
    ws[f"AD{r_}"] = f'=IF($AC{r_}="","",TEXT($AC{r_},"0")&"  "&INDEX($B$2:$B$2001,MATCH($AC{r_},$A$2:$A$2001,0)))'
    for col_ in ("AB", "AC", "AD"):
        ws[f"{col_}{r_}"].font = f_norm
ws.column_dimensions["AD"].width = 40
name("RC_LIST_NUM", f"EXPENSE_ACCOUNTS!$AC$2:$AC${1 + N_RC_LIST}")
name("RC_LIST_TXT", f"EXPENSE_ACCOUNTS!$AD$2:$AD${1 + N_RC_LIST}")
name("OPF_LIST_NUM", f"EXPENSE_ACCOUNTS!$X$2:$X${1 + N_OPF_LIST}")
name("OPF_LIST_TXT", f"EXPENSE_ACCOUNTS!$Y$2:$Y${1 + N_OPF_LIST}")
ws["O1"] = "FORMAT GROUPS WITH A DETAIL SECTION"
ws["O1"].font = f_bold
for i_, grp_ in enumerate(["INCOME", "UTILITIES", "JANITORIAL & CLEANING", "ELEVATOR", "HVAC & PAYROLL", "LANDSCAPING & GROUNDS", "GENERAL R&M",
                           "TAXES", "INSURANCE", "PROFESSIONAL FEES/G&A", "TENANT IMP. & COMM", "DEBT SERVICE", "OTHER PROFESSIONAL FEES"]):
    ws[f"O{2 + i_}"] = grp_
    ws[f"O{2 + i_}"].font = f_norm
ws.column_dimensions["O"].width = 40
for col, w in zip("ABCDEFGHIJKLMN", (10, 34, 24, 10, 10, 16, 16, 16, 10, 12, 10, 3, 3, 44)):
    ws.column_dimensions[col].width = w
ws.row_dimensions[1].height = 30
ws.freeze_panes = "A2"
ws["Q1"] = ("Calculation sheet (read-only). Lists EXPENSE-type accounts in the six expense sections and the TAXES, INSURANCE and PROFESSIONAL FEES/G&A groups that have prior-year actual, prior-year budget or vendor activity. "
            "Each qualifying account must have a block on an EXP_ sheet, TAXES, INSURANCE, PROF_FEES, OTHER_PROF_FEES (OTHER PROFESSIONAL FEES group, except leasing-cost accounts) or DEBT (DEBT SERVICE group, any account type); the DETAIL status turns to CHECK DATA if one does not (rebuild the workbook).")
ws["Q1"].font = f_note
name("EA_NOBLOCK", "EXPENSE_ACCOUNTS!$K$2:$K$2001")
ws.protection.sheet = True
ws.protection.formatColumns = False
ws.protection.formatRows = False


def build_exp_blocks(sections):
    """{sheet: [(acctno, [(vendid, vendnme, total), ...]), ...]} from the data (accounts and vendors are fixed when the workbook is built)."""
    py = BUDGET_YEAR - 1
    am, bm = list(act.columns[3:15]), list(bud.columns[4:16])
    tot = {}
    for df_, cols, flt in ((act, am, py), (bud, bm, py)):
        d_ = df_[df_["Year"] == flt]
        for a_, v_ in zip(d_["acctnum"], pd.DataFrame({c: pd.to_numeric(d_[c]) for c in cols}).abs().sum(axis=1)):
            tot[a_] = tot.get(a_, 0) + v_
    for a_, v_ in zip(ven["acctnum"], ven[VEN_MONTHS].abs().sum(axis=1)):
        tot[a_] = tot.get(a_, 0) + v_
    out = {}
    for _, grp_, sh_ in sections:
        accs = coa[(coa["FMT_GROUP"] == grp_) & (coa["TYPE"] == "EXPENSE")].sort_values("ACCTNO")
        lst = []
        for r_ in accs.itertuples():
            if tot.get(r_.ACCTNUM, 0) <= 0.004:
                continue
            v_ = ven[ven["acctnum"] == r_.ACCTNUM].copy()
            v_["t"] = v_[VEN_MONTHS].abs().sum(axis=1)
            v_ = v_.sort_values(["t", "vendnme1"], ascending=[False, True])
            lst.append((int(r_.ACCTNO), [(x.vendid, x.vendnme1, x.t) for x in v_.itertuples()]))
        out[sh_] = lst
    return out


INS_SECTION = (11, "INSURANCE", "INSURANCE")
PRO_SECTION = (12, "PROFESSIONAL FEES/G&A", "PROF_FEES")
EXP_BLOCKS = build_exp_blocks(EXP_SECTIONS + [INS_SECTION, PRO_SECTION])


INS_DEFAULT_PCT = {5750000: 0.0215, 5755000: 0.01, 5805000: 0.05}      # legacy 2026 budget assumptions; other accounts default to 0%


def build_exp_sheet(ws, sec_title, blocks, ins=False, pro=False):
    ins = ins or pro
    if ins:
        global ins_dv, ins_pct_dv
        ins_dv = DataValidation(type="list", formula1='"Yes,No"', allow_blank=False, showErrorMessage=True)
        ins_pct_dv = DataValidation(type="decimal", operator="between", formula1="0", formula2="1", allow_blank=False, showErrorMessage=True,
                                    errorTitle="Increase %", error="Enter a percentage between 0% and 100%.")
        ws.add_data_validation(ins_dv)
        ws.add_data_validation(ins_pct_dv)
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDEF", (10, 2, 46, 11, 8, 2)):
        ws.column_dimensions[col].width = w
    for m in range(12):
        ws.column_dimensions[CL(7 + m)].width = 11
    ws.column_dimensions["S"].width = 13
    ws.column_dimensions["T"].width = 22
    ws.column_dimensions["U"].width = 7
    put(ws, "A1", f'=BuildingName&" - {sec_title}"' if ins else f'=BuildingName&" - {sec_title} EXPENSES"', f_title)
    put(ws, "A2", '="BUDGET FOR THE YEAR "&BudgetYear', f_bold)
    if pro:
        put(ws, "A3", "Management fee = fee % x the monthly rent base (panel below). Other accounts: prior-year forecast for each month x (1 + the increase % in the block header). "
                      "Vendor lines show prior-year payments (read-only) with a light-blue adjustment line (+/-) below; the three open lines add amounts. Set 'Use calc' to No to ignore the calculated line and type your own amounts.", f_sub)
    elif ins:
        put(ws, "A3", "One block per GL account with prior-year activity. The calculated line = the prior-year forecast for each month x (1 + the increase % in the block header). "
                      "Vendor lines show prior-year payments (read-only) with a light-blue adjustment line (+/-) below; the three open lines add amounts. Set 'Use calc' to No to ignore the calculated line and type your own amounts.", f_sub)
    else:
        put(ws, "A3", "One block per GL account with prior-year actual, budget or vendor activity. Each vendor shows its prior-year payments by month (read-only) with a light-blue line below "
                      "for the budget-year amounts you expect (expenses as positive numbers). Use the three open light-blue lines for vendors not listed.", f_sub)
    put(ws, "A4", '="Prior-year forecast = actuals Jan-"&TEXT(DATE(BudgetYear-1,ActualsThru,1),"mmm")&" plus budget for the remaining months (set on INPUT). Vendor lines are payments by check period."', f_sub)
    for col, t in {"A": "GL Acct", "C": "Account / Vendor / Line", "D": "GL / Vendor", "E": "Source"}.items():
        put(ws, f"{col}5", t, f_bold)
    for m in range(12):
        put(ws, f"{CL(7 + m)}5", f"=DATE(BudgetYear,{m + 1},1)", f_bold, "mmm-yy", align="right")
    put(ws, "S5", "Total", f_bold, align="right")
    put(ws, "T5", "$/SF", f_bold, align="right")
    put(ws, "U5", "Block", f_bold)
    for col in range(1, 22):
        ws.cell(5, col).border = Border(bottom=thin)
    for rr, lbl, tag in ((6, "PRIOR-YEAR FORECAST - ALL ACCOUNTS", "FCST"), (7, "BUDGET-YEAR TOTAL - ALL ACCOUNTS (this sheet)" if ins else "BUDGET-YEAR INPUT - ALL ACCOUNTS (this sheet)", "FINAL" if ins else "INPUT")):
        put(ws, f"C{rr}", lbl, f_bold)
        for m in range(12):
            col = CL(7 + m)
            put(ws, f"{col}{rr}", f'=SUMIFS({col}${EXP_FIRST}:{col}${EXP_LAST},$E${EXP_FIRST}:$E${EXP_LAST},"{tag}")', f_bold, NUM0)
        put(ws, f"S{rr}", f"=SUM(G{rr}:R{rr})", f_bold, NUM0)
        put(ws, f"T{rr}", f"=IF(BillingSF=0,0,S{rr}/BillingSF)", f_bold, NUM2)

    def calc_cells(r, bold=False):
        for m in range(12):
            c = ws.cell(r, 7 + m)
            c.number_format = NUM0
            c.font = f_bold if bold else f_norm
            c.border = Border(bottom=hair)

    def input_cells(r):
        for m in range(12):
            c = ws.cell(r, 7 + m)
            c.fill = INPUT_FILL
            c.protection = UNLOCK
            c.number_format = NUM0
            c.font = f_norm
            c.border = Border(bottom=hair)
        put(ws, f"S{r}", f"=SUM(G{r}:R{r})", fmt=NUM0_BZ)
        put(ws, f"T{r}", f"=IF(BillingSF=0,0,S{r}/BillingSF)", fmt=NUM2_BZ)

    if pro:
        put(ws, "C9", "MANAGEMENT FEE (5800000)", f_bold)
        for col in range(1, 22):
            ws.cell(9, col).border = Border(bottom=thin)
        put(ws, "C10", "Fee % on base rent (net of concessions)")
        put(ws, "G10", 0.03, f_bold, "0.00%", inp=True, align="right")
        put(ws, "I10", "From the building list in MRI (Allwood 3.0%). Applied to the monthly base rent accounts below.", f_note)
        put(ws, "C11", "Fee % on operating expense reimbursement")
        put(ws, "G11", 0.0, f_bold, "0.00%", inp=True, align="right")
        put(ws, "I11", "From the building list in MRI (Allwood 0.0%).", f_note)
        put(ws, "C12", "Base rent GL accounts (up to 2)")
        put(ws, "G12", 4100000, f_bold, "0", inp=True, align="right")
        put(ws, "H12", 4150000, f_bold, "0", inp=True, align="right")
        put(ws, "I12", "Default = base rent + rent concession. The base is what REVENUE and OTHER_REVENUE show for these accounts, so lease options flow into the fee.", f_note)
        put(ws, "C13", "Operating expense reimbursement GL account")
        put(ws, "G13", 4210000, f_bold, "0", inp=True, align="right")
        put(ws, "I13", "Default = operating expense reimbursement.", f_note)
        pro_pct_dv = DataValidation(type="decimal", operator="between", formula1="0", formula2="1", allow_blank=False, showErrorMessage=True,
                                    errorTitle="Fee %", error="Enter a percentage between 0% and 100%.")
        ws.add_data_validation(pro_pct_dv)
        pro_pct_dv.add("G10")
        pro_pct_dv.add("G11")

    r0 = 16 if pro else EXP_FIRST
    for k, (acctno, vendors) in enumerate(blocks, 1):
        A = f"$A${r0}"
        ra, rb, rf = r0 + 1, r0 + 2, r0 + 3
        put(ws, f"U{r0}", k, f_lbl)
        put(ws, f"A{r0}", acctno, f_bold, "0000\\-000", align="left")
        put(ws, f"C{r0}", f'=IFERROR(INDEX(COA_NAME,MATCH({A},COA_ACCTNO,0)),"(not in chart of accounts)")', f_bold)
        for col in range(1, 22):
            ws.cell(r0, col).border = Border(bottom=thin)
        is_mgmt = pro and acctno == 5800000
        if ins:
            if is_mgmt:
                put(ws, f"G{r0}", "Rate: see panel", f_lbl, align="right")
            else:
                put(ws, f"G{r0}", "Increase %", f_lbl, align="right")
                put(ws, f"H{r0}", INS_DEFAULT_PCT.get(acctno, 0.0), f_bold, "0.00%", inp=True, align="right")
                ins_pct_dv.add(f"H{r0}")
            put(ws, f"I{r0}", "Use calc?", f_lbl, align="right")
            put(ws, f"J{r0}", "Yes", f_bold, inp=True, align="center")
            ins_dv.add(f"J{r0}")
        put(ws, f"C{ra}", f'="  "&(BudgetYear-1)&" actual (Jan-"&TEXT(DATE(BudgetYear-1,ActualsThru,1),"mmm")&")"')
        put(ws, f"C{rb}", f'=IF(ActualsThru>=12,"  "&(BudgetYear-1)&" budget (none - full year actual)","  "&(BudgetYear-1)&" budget ("&TEXT(DATE(BudgetYear-1,ActualsThru+1,1),"mmm")&"-Dec)")')
        put(ws, f"C{rf}", f'="  "&(BudgetYear-1)&" forecast"', f_bold)
        for rowi, tag in ((ra, "ACT"), (rb, "BUD"), (rf, "FCST")):
            put(ws, f"E{rowi}", tag, f_note)
        for m in range(12):
            col = CL(7 + m)
            ac, bc = CL(4 + m), CL(5 + m)
            ws[f"{col}{ra}"] = f'=IF(MONTH({col}$5)<=ActualsThru,SUMIFS({DA}{ac}$2:{ac}$2000,{DA}$C$2:$C$2000,"MR"&{A},{DA}$B$2:$B$2000,BudgetYear-1),0)'
            ws[f"{col}{rb}"] = f'=IF(MONTH({col}$5)>ActualsThru,SUMIFS({DB}{bc}$2:{bc}$2000,{DB}$D$2:$D$2000,"MR"&{A},{DB}$B$2:$B$2000,BudgetYear-1),0)'
            ws[f"{col}{rf}"] = f"={col}{ra}+{col}{rb}"
        for rowi, bold in ((ra, False), (rb, False), (rf, True)):
            calc_cells(rowi, bold)
            put(ws, f"S{rowi}", f"=SUM(G{rowi}:R{rowi})", f_bold if bold else f_norm, NUM0)
            put(ws, f"T{rowi}", f"=IF(BillingSF=0,0,S{rowi}/BillingSF)", f_bold if bold else f_norm, NUM2)
        r = r0 + 4
        rc = None
        if ins:
            rc = r
            put(ws, f"C{rc}", f'="  "&BudgetYear&" fee % x monthly rent base (calculated)"' if is_mgmt else f'="  "&BudgetYear&" prior-year forecast x (1 + increase) (calculated)"', f_norm)
            put(ws, f"E{rc}", "CALC", f_note)
            for m in range(12):
                col = CL(7 + m)
                if is_mgmt:
                    def rent_(cell):
                        return (f'IF({cell}="",0,SUMIFS(REVENUE!{col}$8:{col}$1500,REVENUE!$D$8:$D$1500,{cell})'
                                f'+SUMIFS(OTHER_REVENUE!{col}$9:{col}$600,OTHER_REVENUE!$D$9:$D$600,{cell},OTHER_REVENUE!$E$9:$E$600,"INPUT"))')
                    ws[f"{col}{rc}"] = (f'=IF($J${r0}="No",0,$G$10*({rent_("$G$12")}+{rent_("$H$12")})+$G$11*{rent_("$G$13")})')
                else:
                    ws[f"{col}{rc}"] = f'=IF($J${r0}="No",0,{col}{rf}*(1+$H${r0}))'
            calc_cells(rc)
            put(ws, f"S{rc}", f"=SUM(G{rc}:R{rc})", f_norm, NUM0)
            put(ws, f"T{rc}", f"=IF(BillingSF=0,0,S{rc}/BillingSF)", f_norm, NUM2)
            r += 1
        if PAYROLL_ACCT and acctno == PAYROLL_ACCT and not ins:
            put(ws, f"C{r}", f'="  "&BudgetYear&" payroll from INPUT: annual / 52 x pay weeks + summer help (calculated) - leave the vendor budget lines blank"', f_norm)
            put(ws, f"D{r}", f"={A}", fmt="0")
            put(ws, f"E{r}", "INPUT", f_note)
            for m in range(12):
                col = CL(7 + m)
                mm = f"MONTH({col}$5)"
                ws[f"{col}{r}"] = (f"=N(PayAnnual)/52*IF(COUNTIF(PayWeekMonths,{mm})>0,5,4)+IF(AND({mm}>=N(PaySummerFrom),{mm}<=N(PaySummerTo)),N(PaySummerAmt),0)")
            calc_cells(r)
            put(ws, f"S{r}", f"=SUM(G{r}:R{r})", f_norm, NUM0)
            put(ws, f"T{r}", f"=IF(BillingSF=0,0,S{r}/BillingSF)", f_norm, NUM2)
            r += 1
        for vid, vname, _t in vendors:
            rv, ri = r, r + 1
            put(ws, f"D{rv}", vid, f_note)
            ws[f"D{rv}"].number_format = "@"
            put(ws, f"C{rv}", f'="  "&IFERROR(INDEX(DATA_VENDOR!$C$2:$C$3000,MATCH($D{rv},DATA_VENDOR!$B$2:$B$3000,0)),$D{rv})&" - "&(BudgetYear-1)&" paid"', f_norm)
            put(ws, f"E{rv}", "VEND", f_note)
            for m in range(12):
                col, vc = CL(7 + m), CL(5 + m)
                ws[f"{col}{rv}"] = (f'=SUMIFS(DATA_VENDOR!{vc}$2:{vc}$3000,DATA_VENDOR!$B$2:$B$3000,$D{rv},DATA_VENDOR!$D$2:$D$3000,"MR"&{A})')
            calc_cells(rv)
            put(ws, f"S{rv}", f"=SUM(G{rv}:R{rv})", f_norm, NUM0)
            put(ws, f"T{rv}", f"=IF(BillingSF=0,0,S{rv}/BillingSF)", f_norm, NUM2)
            put(ws, f"C{ri}", f'="    "&BudgetYear&" adjustment (+/-)"' if ins else f'="    "&BudgetYear&" budget"', f_note)
            put(ws, f"D{ri}", f"={A}", fmt="0")
            put(ws, f"E{ri}", "INPUT", f_note)
            input_cells(ri)
            r += 2
        for _ in range(3):
            put(ws, f"C{r}", None, inp=True)
            put(ws, f"D{r}", f"={A}", fmt="0")
            put(ws, f"E{r}", "INPUT", f_note)
            input_cells(r)
            r += 1
        rt = r
        put(ws, f"C{rt}", f'="  "&BudgetYear&" total"', f_bold)
        if ins:
            put(ws, f"D{rt}", f"={A}", f_note, "0")
            put(ws, f"E{rt}", "FINAL", f_note)
        for m in range(12):
            col = CL(7 + m)
            base_ = f"{col}{rc}+" if ins else ""
            put(ws, f"{col}{rt}", f'={base_}SUMIFS({col}{r0 + 1}:{col}{rt - 1},$E{r0 + 1}:$E{rt - 1},"INPUT")', f_bold, NUM0)
        put(ws, f"S{rt}", f"=SUM(G{rt}:R{rt})", f_bold, NUM0)
        put(ws, f"T{rt}", f"=IF(BillingSF=0,0,S{rt}/BillingSF)", f_bold, NUM2)
        for col in range(1, 22):
            ws.cell(rt, col).border = Border(top=thin)
        ro = rt + 1
        rr_ = None if ins else rt + 2          # major repairs memo (six EXP_ sheets only)
        rn = rt + 2 if ins else rt + 3
        put(ws, f"C{ro}", f'="  "&BudgetYear&" tenant-billed offset from REVENUE (read-only)"', f_note)
        put(ws, f"E{ro}", "OFFSET", f_note)
        if rr_:
            put(ws, f"C{rr_}", f'="  "&BudgetYear&" major repairs (REPAIRS_CAPITAL, read-only)"', f_note)
            put(ws, f"E{rr_}", "MEMO", f_note)
        put(ws, f"C{rn}", f'="  "&BudgetYear&" net of tenant-billed offset' + ('' if ins else ' + major repairs') + f' (as on DETAIL)"', f_bold)
        put(ws, f"E{rn}", "NET", f_note)
        for m in range(12):
            col = CL(7 + m)
            ws[f"{col}{ro}"] = f'=-SUMIFS(REVENUE!{col}$8:{col}$1500,REVENUE!$D$8:$D$1500,{A},REVENUE!$V$8:$V$1500,"<>IGNORED*",REVENUE!$E$8:$E$1500,"<>COST")'
            if rr_:
                rcc = CL(RC_M0 + m)
                ws[f"{col}{rr_}"] = (f'=SUMIFS(REPAIRS_CAPITAL!{rcc}${RC_REP_FIRST}:{rcc}${RC_REP_LAST},REPAIRS_CAPITAL!$D${RC_REP_FIRST}:$D${RC_REP_LAST},{A},'
                                     f'REPAIRS_CAPITAL!$E${RC_REP_FIRST}:$E${RC_REP_LAST},"REPAIR")')
            ws[f"{col}{rn}"] = f"={col}{rt}+{col}{ro}" + (f"+{col}{rr_}" if rr_ else "")
        calc_cells(ro)
        if rr_:
            calc_cells(rr_)
        calc_cells(rn, True)
        if rr_:
            put(ws, f"S{rr_}", f"=SUM(G{rr_}:R{rr_})", f_norm, NUM0)
            put(ws, f"T{rr_}", f"=IF(BillingSF=0,0,S{rr_}/BillingSF)", f_norm, NUM2)
        put(ws, f"S{ro}", f"=SUM(G{ro}:R{ro})", f_norm, NUM0)
        put(ws, f"T{ro}", f"=IF(BillingSF=0,0,S{ro}/BillingSF)", f_norm, NUM2)
        put(ws, f"S{rn}", f"=SUM(G{rn}:R{rn})", f_bold, NUM0)
        put(ws, f"T{rn}", f"=IF(BillingSF=0,0,S{rn}/BillingSF)", f_bold, NUM2)
        ws[f"T{r0}"] = (f'=IF(COUNTIF(DATA_VENDOR!$D$2:$D$3000,"MR"&{A})>COUNTIF($E{r0}:$E{rt},"VEND"),"NEW VENDOR - REBUILD","")')
        ws[f"T{r0}"].font = f_bold
        r0 = rt + (4 if ins else 5)
    assert r0 < EXP_LAST, "%s needs more than %d rows" % (sec_title, EXP_LAST)
    ws.freeze_panes = "D6"
    ws.protection.sheet = True
    ws.protection.formatColumns = False
    ws.protection.formatRows = False
    return r0


for _sec, _grp, _sh in EXP_SECTIONS:
    _end = build_exp_sheet(ws_exp[_sh], _grp, EXP_BLOCKS[_sh])
    print(f"{_sh}: {len(EXP_BLOCKS[_sh])} account blocks, {sum(len(v) for _, v in EXP_BLOCKS[_sh])} vendor lines, rows to {_end}")
_end = build_exp_sheet(ws_pro, "PROFESSIONAL FEES/G&A", EXP_BLOCKS["PROF_FEES"], pro=True)
print(f"PROF_FEES: {len(EXP_BLOCKS['PROF_FEES'])} account blocks, {sum(len(v) for _, v in EXP_BLOCKS['PROF_FEES'])} vendor lines, rows to {_end}")
_end = build_exp_sheet(ws_ins, "INSURANCE", EXP_BLOCKS["INSURANCE"], ins=True)
print(f"INSURANCE: {len(EXP_BLOCKS['INSURANCE'])} account blocks, {sum(len(v) for _, v in EXP_BLOCKS['INSURANCE'])} vendor lines, rows to {_end}")



# ------------------------------------------------------------------ OTHER_PROF_FEES (input: OTHER PROFESSIONAL FEES group - forecast + increase %, monthly override, adjustments)
OPF_INC_DEFAULT = 0.01
N_OPF_USER = 5


def opf_accounts():
    """OTHER PROFESSIONAL FEES-group EXPENSE accounts with prior-year actual, budget or vendor activity (leasing-cost accounts on INPUT stay in IMPRVMTS & COMMISSIONS)."""
    py = BUDGET_YEAR - 1
    am, bm = list(act.columns[3:15]), list(bud.columns[4:16])
    tot = {}
    for df_, cols, flt in ((act, am, py), (bud, bm, py)):
        d_ = df_[df_["Year"] == flt]
        for a_, v_ in zip(d_["acctnum"], pd.DataFrame({c: pd.to_numeric(d_[c]) for c in cols}).abs().sum(axis=1)):
            tot[a_] = tot.get(a_, 0) + v_
    for a_, v_ in zip(ven["acctnum"], ven[VEN_MONTHS].abs().sum(axis=1)):
        tot[a_] = tot.get(a_, 0) + v_
    lease = {a_ for _, a_ in LC}
    accs = coa[(coa["FMT_GROUP"] == OPF_GROUP) & (coa["TYPE"] == "EXPENSE")].sort_values("ACCTNO")
    return [int(r_.ACCTNO) for r_ in accs.itertuples() if tot.get(r_.ACCTNUM, 0) > 0.004 and int(r_.ACCTNO) not in lease]


def build_opf_sheet(ws, accts):
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDEF", (10, 2, 54, 11, 8, 2)):
        ws.column_dimensions[col].width = w
    for m in range(12):
        ws.column_dimensions[CL(7 + m)].width = 11
    ws.column_dimensions["S"].width = 13
    ws.column_dimensions["T"].width = 24
    ws.column_dimensions["U"].width = 7
    dv_yn = DataValidation(type="list", formula1='"Yes,No"', allow_blank=False, showErrorMessage=True)
    dv_pct = DataValidation(type="decimal", operator="between", formula1="0", formula2="1", allow_blank=False, showErrorMessage=True,
                            errorTitle="Increase %", error="Enter a percentage between 0% and 100%.")
    dv_user = DataValidation(type="list", formula1="=OPF_LIST_TXT", allow_blank=True, showErrorMessage=True,
                             errorTitle="Account", error="Pick an account from the list (OTHER PROFESSIONAL FEES group; leasing-cost accounts are excluded).")
    for v_ in (dv_yn, dv_pct, dv_user):
        ws.add_data_validation(v_)
    put(ws, "A1", '=BuildingName&" - OTHER PROFESSIONAL FEES"', f_title)
    put(ws, "A2", '="BUDGET FOR THE YEAR "&BudgetYear', f_bold)
    put(ws, "A3", "One block per OTHER PROFESSIONAL FEES GL account with prior-year activity, then five USER blocks for any other account in the group. Calculated line = prior-year forecast for each month x (1 + the increase % in the block header, default 1%). "
                  "Type an amount on the light-blue OVERRIDE line to replace the calculated amount for that month (blank = use calculated; 0 = budget zero). The adjustment and open lines are added to the amount used. "
                  "Use calc = No ignores the calculated line (months without an override become zero).", f_sub)
    put(ws, "A4", '="Prior-year forecast = actuals Jan-"&TEXT(DATE(BudgetYear-1,ActualsThru,1),"mmm")&" plus budget for the remaining months (set on INPUT). Expenses are positive numbers."', f_sub)
    for col, t in {"A": "GL Acct", "C": "Account / Line", "D": "GL", "E": "Source"}.items():
        put(ws, f"{col}5", t, f_bold)
    for m in range(12):
        put(ws, f"{CL(7 + m)}5", f"=DATE(BudgetYear,{m + 1},1)", f_bold, "mmm-yy", align="right")
    put(ws, "S5", "Total", f_bold, align="right")
    put(ws, "T5", "$/SF", f_bold, align="right")
    put(ws, "U5", "Block", f_bold)
    for col in range(1, 22):
        ws.cell(5, col).border = Border(bottom=thin)
    EF, EL = EXP_FIRST, EXP_LAST
    for rr, lbl, tag in ((6, "PRIOR-YEAR FORECAST - ALL ACCOUNTS", "FCST"), (7, "BUDGET-YEAR TOTAL - ALL ACCOUNTS (this sheet)", "FINAL")):
        put(ws, f"C{rr}", lbl, f_bold)
        for m in range(12):
            col = CL(7 + m)
            put(ws, f"{col}{rr}", f'=SUMIFS({col}${EF}:{col}${EL},$E${EF}:$E${EL},"{tag}")', f_bold, NUM0)
        put(ws, f"S{rr}", f"=SUM(G{rr}:R{rr})", f_bold, NUM0)
        put(ws, f"T{rr}", f"=IF(BillingSF=0,0,S{rr}/BillingSF)", f_bold, NUM2)
    r0 = EXP_FIRST
    n_auto = len(accts)
    for k in range(1, n_auto + N_OPF_USER + 1):
        is_user = k > n_auto
        A = f"$A${r0}"
        ra, rb, rf, rc, ro, ru, ri = r0 + 1, r0 + 2, r0 + 3, r0 + 4, r0 + 5, r0 + 6, r0 + 7
        rt = r0 + 11
        W = (lambda e, A=A: f'=IF({A}="","",{e})') if is_user else (lambda e: "=" + e)
        if not is_user:
            put(ws, f"U{r0}", k, f_lbl)
            put(ws, f"A{r0}", accts[k - 1], f_bold, "0000\\-000", align="left")
            put(ws, f"C{r0}", f'=IFERROR(INDEX(COA_NAME,MATCH({A},COA_ACCTNO,0)),"(not in chart of accounts)")', f_bold)
        else:
            if k == n_auto + 1:
                put(ws, f"A{r0 - 1}", "USER BLOCKS - budget any other OTHER PROFESSIONAL FEES account: pick it in the light-blue cell of a block, then use its lines.", f_bold)
            put(ws, f"U{r0}", f"U{k - n_auto}", f_lbl)
            put(ws, f"A{r0}", f'=IFERROR(VALUE(LEFT($C{r0},7)),"")', f_bold, "0000\\-000", align="left")
            put(ws, f"C{r0}", None, f_bold, inp=True)
            dv_user.add(f"C{r0}")
            put(ws, f"L{r0}", f'=IF($C{r0}="","<- select an account in the light-blue cell","")', f_note)
            ws[f"T{r0}"] = (f'=IF($C{r0}="","",IF(ISNA(MATCH($A{r0},OPF_LIST_NUM,0)),"NOT AN OTHER PROF FEE ACCT",'
                            f'IF(COUNTIF($A${EF}:$A${EL},$A{r0})>1,"DUPLICATE ACCOUNT","")))')
            ws[f"T{r0}"].font = f_bold
        for col in range(1, 22):
            ws.cell(r0, col).border = Border(bottom=thin)
        put(ws, f"G{r0}", "Increase %", f_lbl, align="right")
        put(ws, f"H{r0}", OPF_INC_DEFAULT, f_bold, "0.00%", inp=True, align="right")
        dv_pct.add(f"H{r0}")
        put(ws, f"I{r0}", "Use calc?", f_lbl, align="right")
        put(ws, f"J{r0}", "Yes", f_bold, inp=True, align="center")
        dv_yn.add(f"J{r0}")
        lbl = lambda e: f'=IF({A}="","",{e})' if is_user else "=" + e
        put(ws, f"C{ra}", lbl('"  "&(BudgetYear-1)&" actual (Jan-"&TEXT(DATE(BudgetYear-1,ActualsThru,1),"mmm")&")"'))
        put(ws, f"C{rb}", lbl('IF(ActualsThru>=12,"  "&(BudgetYear-1)&" budget (none - full year actual)","  "&(BudgetYear-1)&" budget ("&TEXT(DATE(BudgetYear-1,ActualsThru+1,1),"mmm")&"-Dec)")'))
        put(ws, f"C{rf}", lbl('"  "&(BudgetYear-1)&" forecast"'), f_bold)
        put(ws, f"C{rc}", lbl('"  "&BudgetYear&" prior-year forecast x (1 + increase) (calculated)"'))
        put(ws, f"C{ro}", lbl('"    "&BudgetYear&" OVERRIDE by month (blank = use calculated; 0 = zero)"'), f_bold)
        put(ws, f"C{ru}", lbl('"  "&BudgetYear&" amount used (override, else calculated)"'), f_bold)
        put(ws, f"C{ri}", lbl('"    "&BudgetYear&" adjustment (+/-)"'), f_note)
        for rowi, tag in ((ra, "ACT"), (rb, "BUD"), (rf, "FCST"), (rc, "CALC"), (ro, "OVR"), (ru, "USED")):
            put(ws, f"E{rowi}", tag, f_note)
        for m in range(12):
            col, ac, bc = CL(7 + m), CL(4 + m), CL(5 + m)
            ws[f"{col}{ra}"] = W(f'IF(MONTH({col}$5)<=ActualsThru,SUMIFS({DA}{ac}$2:{ac}$2000,{DA}$C$2:$C$2000,"MR"&{A},{DA}$B$2:$B$2000,BudgetYear-1),0)')
            ws[f"{col}{rb}"] = W(f'IF(MONTH({col}$5)>ActualsThru,SUMIFS({DB}{bc}$2:{bc}$2000,{DB}$D$2:$D$2000,"MR"&{A},{DB}$B$2:$B$2000,BudgetYear-1),0)')
            ws[f"{col}{rf}"] = W(f"{col}{ra}+{col}{rb}")
            ws[f"{col}{rc}"] = W(f'IF($J${r0}="No",0,{col}{rf}*(1+$H${r0}))')
            ws[f"{col}{ru}"] = W(f"IF(ISNUMBER({col}{ro}),{col}{ro},{col}{rc})")
        for rowi, bold in ((ra, False), (rb, False), (rf, True), (rc, False), (ru, True)):
            for m in range(12):
                c = ws.cell(rowi, 7 + m)
                c.number_format = NUM0
                c.font = f_bold if bold else f_norm
                c.border = Border(bottom=hair)
            put(ws, f"S{rowi}", W(f"SUM(G{rowi}:R{rowi})"), f_bold if bold else f_norm, NUM0)
            put(ws, f"T{rowi}", W(f"IF(BillingSF=0,0,S{rowi}/BillingSF)"), f_bold if bold else f_norm, NUM2)
        in_rows = [ri, ri + 1, ri + 2, ri + 3]
        for x in in_rows[1:]:
            put(ws, f"C{x}", None, inp=True)
        for x in in_rows:
            put(ws, f"D{x}", W(A) if is_user else f"={A}", fmt="0")
            put(ws, f"E{x}", "INPUT", f_note)
        for x in in_rows + [ro]:
            for m in range(12):
                c = ws.cell(x, 7 + m)
                c.fill = INPUT_FILL
                c.protection = UNLOCK
                c.number_format = NUM0
                c.font = f_norm
                c.border = Border(bottom=hair)
            put(ws, f"S{x}", f"=SUM(G{x}:R{x})", fmt=NUM0_BZ)
            put(ws, f"T{x}", f"=IF(BillingSF=0,0,S{x}/BillingSF)", fmt=NUM2_BZ)
        put(ws, f"C{rt}", lbl('"  "&BudgetYear&" total"'), f_bold)
        put(ws, f"D{rt}", W(A) if is_user else f"={A}", f_note, "0")
        put(ws, f"E{rt}", "FINAL", f_note)
        for m in range(12):
            col = CL(7 + m)
            put(ws, f"{col}{rt}", W(f'{col}{ru}+SUMIFS({col}{ri}:{col}{rt - 1},$E{ri}:$E{rt - 1},"INPUT")'), f_bold, NUM0)
        put(ws, f"S{rt}", W(f"SUM(G{rt}:R{rt})"), f_bold, NUM0)
        put(ws, f"T{rt}", W(f"IF(BillingSF=0,0,S{rt}/BillingSF)"), f_bold, NUM2)
        for col in range(1, 22):
            ws.cell(rt, col).border = Border(top=thin)
        r0 = rt + 3
    assert r0 < EXP_LAST, "OTHER_PROF_FEES needs more than %d rows" % EXP_LAST
    ws.freeze_panes = "D6"
    ws.protection.sheet = True
    ws.protection.formatColumns = False
    ws.protection.formatRows = False
    return r0


OPF_ACCTS = opf_accounts()
_end = build_opf_sheet(ws_opf, OPF_ACCTS)
print(f"OTHER_PROF_FEES: {len(OPF_ACCTS)} account blocks + {N_OPF_USER} user blocks {OPF_ACCTS}, rows to {_end}; dropdown accounts {N_OPF_LIST}")


# ------------------------------------------------------------------ REPAIRS_CAPITAL (input: major repairs by GL account + capital improvements -> DETAIL)
def build_rc_sheet(ws):
    ws.sheet_view.showGridLines = False
    for col, w in {"A": 6, "B": 8, "C": 42, "D": 10, "E": 9, "F": 36, "G": 7, "H": 13, "I": 12, "J": 12, "K": 15, "L": 13, "M": 8, "N": 8, "O": 13,
                   "AN": 12, "AO": 60, "AP": 40}.items():
        ws.column_dimensions[col].width = w
    for m in range(12):
        ws.column_dimensions[CL(16 + m)].width = 10
        ws.column_dimensions[CL(RC_M0 + m)].width = 10
    dv_yn = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True, showErrorMessage=True, errorTitle="Include?", error="Choose Yes or No.")
    dv_basis = DataValidation(type="list", formula1='"Quote,Estimate,Allowance"', allow_blank=True, showErrorMessage=True, errorTitle="Basis", error="Choose Quote, Estimate or Allowance.")
    dv_prio = DataValidation(type="list", formula1='"1,2,3,Next year"', allow_blank=True, showErrorMessage=True, errorTitle="Priority", error="Choose 1, 2, 3 or Next year.")
    dv_year = DataValidation(type="whole", operator="between", formula1="1990", formula2="2100", allow_blank=True, showErrorMessage=True,
                             errorTitle="Year", error="Enter a four-digit year.")
    dv_amt = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True, showErrorMessage=True,
                            errorTitle="Amount", error="Enter a positive amount (expenses are positive numbers).")
    dv_m = DataValidation(type="whole", operator="between", formula1="1", formula2="12", allow_blank=True, showErrorMessage=True,
                          errorTitle="Month", error="Enter a month number from 1 to 12.")
    dv_acct = DataValidation(type="list", formula1="=RC_LIST_TXT", allow_blank=True, showErrorMessage=True, errorTitle="GL account",
                             error="Pick an expense account from the list (the six expense sections: utilities, janitorial, elevator, HVAC & payroll, landscaping, general R&M).")
    dv_cap = DataValidation(type="whole", operator="between", formula1="1000000", formula2="9999999", allow_blank=True, showErrorMessage=True,
                            errorTitle="GL account", error="Enter a 7-digit GL account number.")
    for v_ in (dv_yn, dv_basis, dv_prio, dv_year, dv_amt, dv_m, dv_acct, dv_cap):
        ws.add_data_validation(v_)
    put(ws, "A1", '=BuildingName&" - REPAIRS & CAPITAL"', f_title)
    put(ws, "A2", '="BUDGET FOR THE YEAR "&BudgetYear', f_bold)
    put(ws, "A3", "One line per project. MAJOR REPAIRS post to the GL account you pick and add to that account's expense on DETAIL; CAPITAL IMPROVEMENTS all post to one capital GL account "
                  "(default 1345000) under IMPRVMTS & COMMISSIONS. Enter expenses as positive amounts. Set Include? to No to keep a line on the sheet but leave it out of the budget.", f_sub)
    put(ws, "A4", "Budget-year amount: leave blank to use total cost less amount already spent, or type the amount for this budget year (0 is allowed). It is spread evenly from the start month "
                  "over the number of months (any remainder goes in the last month).", f_sub)
    put(ws, "A5", "For uneven timing type amounts in the light-blue month cells (P to AA). If ANY month cell has a number, those 12 cells ARE the schedule (a blank month = $0) and the start month "
                  "and months are ignored. Tenant improvements, commissions, legal, architectural and permits stay on the REVENUE lease cost lines.", f_sub)
    for rr, lbl, src in ((7, "MAJOR REPAIRS - budget-year total by month", RC_REP_TOT), (8, "CAPITAL IMPROVEMENTS - budget-year total by month", RC_CAP_TOT)):
        put(ws, f"C{rr}", lbl, f_bold)
        for m in range(12):
            col = CL(RC_M0 + m)
            put(ws, f"{col}{rr}", f"={col}{src}", f_bold, NUM0)
        put(ws, f"AN{rr}", f"=AN{src}", f_bold, NUM0)

    def header(row, band, prio):
        put(ws, f"A{row - 1}", band, f_bold)
        put(ws, f"P{row - 1}", "OPTIONAL: type amounts by month (blank = spread evenly from the start month)", f_bold)
        put(ws, f"{CL(RC_M0)}{row - 1}", "BUDGET BY MONTH (calculated)", f_bold)
        heads = {"A": "Line", "B": "Include?", "C": "Description", "D": "GL #", "E": "Source", "F": "GL account" if not prio else "GL account (set above)", "G": "Year",
                 "H": "Priority" if prio else "Quote / Estimate / Allowance", "I": "Total cost", "J": "Spent in prior years",
                 "K": "Budget-year amount (blank = cost less spent)", "L": "Budget-year amount used", "M": "Start month (1-12)", "N": "Months",
                 "O": "Remaining after budget year", "AN": "Total", "AO": "Line check", "AP": "Notes / comments"}
        for col, t in heads.items():
            c = put(ws, f"{col}{row}", t, f_bold)
            c.alignment = Alignment(wrap_text=True, horizontal="center" if col not in ("C", "F", "AO", "AP") else "left", vertical="bottom")
        for m in range(12):
            for base in (16, RC_M0):
                c = put(ws, f"{CL(base + m)}{row}", f"=DATE(BudgetYear,{m + 1},1)", f_bold, "mmm-yy", align="right")
                c.alignment = Alignment(horizontal="right", vertical="bottom")
        for col in range(1, 43):
            ws.cell(row, col).border = Border(bottom=thin)
        ws.row_dimensions[row].height = 54

    def line(r, k, cap):
        first, last = (RC_CAP_FIRST, RC_CAP_LAST) if cap else (RC_REP_FIRST, RC_REP_LAST)
        ov = f"$P{r}:$AA{r}"
        put(ws, f"A{r}", k, f_lbl, align="center")
        put(ws, f"B{r}", "Yes", f_norm, inp=True, align="center")
        dv_yn.add(f"B{r}")
        put(ws, f"C{r}", None, f_norm, inp=True)
        if cap:
            put(ws, f"D{r}", f'=IF(ISNUMBER(MATCH($D${RC_CAP_ACCT},COA_ACCTNO,0)),$D${RC_CAP_ACCT},"")', f_note, "0")
            put(ws, f"E{r}", "FINAL", f_note)
            put(ws, f"F{r}", f"=$F${RC_CAP_ACCT}", f_note)
        else:
            put(ws, f"D{r}", f'=IFERROR(VALUE(LEFT($F{r},7)),"")', f_note, "0")
            put(ws, f"E{r}", f'=IF($D{r}="","",IF(ISNA(MATCH($D{r},RC_LIST_NUM,0)),"IGNORED","REPAIR"))', f_note)
            put(ws, f"F{r}", None, f_norm, inp=True)
            dv_acct.add(f"F{r}")
        put(ws, f"G{r}", None, f_norm, "0", inp=True, align="center")
        dv_year.add(f"G{r}")
        put(ws, f"H{r}", None, f_norm, inp=True, align="center")
        (dv_prio if cap else dv_basis).add(f"H{r}")
        for col in "IJK":
            put(ws, f"{col}{r}", None, f_norm, "#,##0", inp=True, align="right")
            dv_amt.add(f"{col}{r}")
        put(ws, f"L{r}", f'=IF($B{r}="No",0,IF(ISNUMBER($K{r}),$K{r},IF(ISNUMBER($I{r}),MAX(0,$I{r}-N($J{r})),IF(COUNT({ov})>0,SUM({ov}),0))))', f_norm, NUM0)
        put(ws, f"M{r}", None, f_norm, "0", inp=True, align="center")
        dv_m.add(f"M{r}")
        put(ws, f"N{r}", 1, f_norm, "0", inp=True, align="center")
        dv_m.add(f"N{r}")
        put(ws, f"O{r}", f'=IF(OR($B{r}="No",NOT(ISNUMBER($I{r}))),"",$I{r}-N($J{r})-$L{r})', f_norm, NUM0)
        n_ = f"MAX(1,N($N{r}))"
        for m in range(1, 13):
            c = ws.cell(r, 15 + m)
            c.fill = INPUT_FILL
            c.protection = UNLOCK
            c.number_format = "#,##0"
            c.font = f_norm
            c.border = Border(bottom=hair)
            dv_amt.add(c.coordinate)
            put(ws, f"{CL(RC_M0 + m - 1)}{r}",
                f'=IF($B{r}="No",0,IF(COUNT({ov})>0,N({CL(15 + m)}{r}),IF(AND(ISNUMBER($M{r}),$M{r}>=1,$M{r}<=12),'
                f'IF(OR({m}<$M{r},{m}>$M{r}+{n_}-1),0,IF({m}=$M{r}+{n_}-1,$L{r}-ROUND($L{r}/{n_},2)*({n_}-1),ROUND($L{r}/{n_},2))),0)))', f_norm, NUM0)
        put(ws, f"AN{r}", f"=SUM({CL(RC_M0)}{r}:{CL(RC_M0 + 11)}{r})", f_bold, NUM0)
        act = f"OR(ISNUMBER($I{r}),ISNUMBER($K{r}),COUNT({ov})>0)"
        dup = (f"AND($L{r}>0,ISNUMBER($M{r}),COUNTIFS($L${RC_REP_FIRST}:$L${RC_REP_LAST},$L{r},$M${RC_REP_FIRST}:$M${RC_REP_LAST},$M{r})"
               f"+COUNTIFS($L${RC_CAP_FIRST}:$L${RC_CAP_LAST},$L{r},$M${RC_CAP_FIRST}:$M${RC_CAP_LAST},$M{r})>1)")
        acct_checks = "" if cap else (f'IF($D{r}="","CHECK - SELECT A GL ACCOUNT",IF($E{r}="IGNORED","CHECK - NOT A BUDGETABLE EXPENSE ACCOUNT",')
        acct_close = "" if cap else "))"
        ws[f"AO{r}"] = (f'=IFERROR(IF(NOT({act}),"",IF($B{r}="No","NOTE - excluded (Include = No)",{acct_checks}'
                        f'IF(AND(COUNT({ov})=0,OR(NOT(ISNUMBER($M{r})),$M{r}<1,$M{r}>12,$M{r}<>INT($M{r}))),"CHECK - ENTER A START MONTH (1-12) OR MONTHLY AMOUNTS",'
                        f'IF(AND(COUNT({ov})>0,ABS(SUM({ov})-$L{r})>0.5),"CHECK - MONTHLY AMOUNTS DO NOT ADD TO THE BUDGET-YEAR AMOUNT",'
                        f'IF(AND(COUNT({ov})=0,$M{r}+MAX(1,N($N{r}))-1>12),"NOTE - SPREAD RUNS PAST DEC (later months are not budgeted)",'
                        f'IF(AND(ISNUMBER($I{r}),N($O{r})<-0.5),"NOTE - BUDGET-YEAR AMOUNT EXCEEDS THE REMAINING COST",'
                        f'IF({dup},"NOTE - POSSIBLE DUPLICATE: another line has the same amount and start month",""))))){acct_close})),"CHECK - ENTRY ERROR")')
        ws[f"AO{r}"].font = f_bold
        put(ws, f"AP{r}", None, f_norm, inp=True)
        for col in range(1, 43):
            if ws.cell(r, col).border.bottom.style is None:
                ws.cell(r, col).border = Border(bottom=hair)

    def totals(r, first, last, tag):
        put(ws, f"C{r}", "TOTAL MAJOR REPAIRS (lines that count)" if tag else "TOTAL CAPITAL IMPROVEMENTS", f_bold)
        for col in "IJO":
            put(ws, f"{col}{r}", f"=SUM({col}{first}:{col}{last})", f_bold, NUM0)
        put(ws, f"L{r}", f'=SUMIFS(L{first}:L{last},$E{first}:$E{last},"{tag or "FINAL"}")', f_bold, NUM0)
        for m in range(12):
            col = CL(RC_M0 + m)
            put(ws, f"{col}{r}", f'=SUMIFS({col}{first}:{col}{last},$E{first}:$E{last},"{tag or "FINAL"}")', f_bold, NUM0)
        put(ws, f"AN{r}", f"=SUM({CL(RC_M0)}{r}:{CL(RC_M0 + 11)}{r})", f_bold, NUM0)
        for col in range(1, 43):
            ws.cell(r, col).border = Border(top=thin)

    header(RC_REP_HDR, "MAJOR REPAIRS - post to the GL account picked on each line and add to that account on DETAIL", False)
    for k in range(1, RC_REP_SLOTS + 1):
        line(RC_REP_FIRST + k - 1, k, False)
    totals(RC_REP_TOT, RC_REP_FIRST, RC_REP_LAST, "REPAIR")
    put(ws, f"A{RC_CAP_BAND}", "CAPITAL IMPROVEMENTS - every line posts to the one GL account below (DETAIL: IMPRVMTS & COMMISSIONS)", f_bold)
    put(ws, f"C{RC_CAP_ACCT}", "GL account for all capital lines:", f_bold)
    put(ws, f"D{RC_CAP_ACCT}", CAP_ACCT_DEFAULT, f_bold, "0", inp=True, align="center")
    dv_cap.add(f"D{RC_CAP_ACCT}")
    put(ws, f"F{RC_CAP_ACCT}", f'=IFERROR(INDEX(COA_NAME,MATCH($D${RC_CAP_ACCT},COA_ACCTNO,0)),"(not in chart of accounts)")', f_bold)
    ws[f"AO{RC_CAP_ACCT}"] = f'=IF(ISNA(MATCH($D${RC_CAP_ACCT},COA_ACCTNO,0)),"CHECK - CAPITAL GL ACCOUNT IS NOT IN THE CHART OF ACCOUNTS","")'
    ws[f"AO{RC_CAP_ACCT}"].font = f_bold
    header(RC_CAP_HDR, "", True)
    ws[f"A{RC_CAP_HDR - 1}"] = None
    for k in range(1, RC_CAP_SLOTS + 1):
        line(RC_CAP_FIRST + k - 1, k, True)
    totals(RC_CAP_TOT, RC_CAP_FIRST, RC_CAP_LAST, None)
    ws.conditional_formatting.add(f"AO{RC_REP_FIRST}:AO{RC_CAP_ACCT}", FormulaRule(formula=[f'LEFT($AO{RC_REP_FIRST},5)="CHECK"'], fill=AMBER))
    ws.conditional_formatting.add(f"AO{RC_CAP_FIRST}:AO{RC_CAP_LAST}", FormulaRule(formula=[f'LEFT($AO{RC_CAP_FIRST},5)="CHECK"'], fill=AMBER))
    ws.conditional_formatting.add(f"C{RC_REP_FIRST}:F{RC_REP_LAST}", FormulaRule(formula=[f'$E{RC_REP_FIRST}="IGNORED"'], fill=AMBER))
    ws.freeze_panes = "D1"
    ws.protection.sheet = True
    ws.protection.formatColumns = False
    ws.protection.formatRows = False


build_rc_sheet(ws_cap)
print(f"REPAIRS_CAPITAL: {RC_REP_SLOTS} major repair lines (rows {RC_REP_FIRST}-{RC_REP_LAST}), {RC_CAP_SLOTS} capital lines (rows {RC_CAP_FIRST}-{RC_CAP_LAST})")

# ------------------------------------------------------------------ DEBT (input: primary mortgage roll-forward + DEBT SERVICE account blocks)
def debt_defaults():
    """Accounts in the DEBT SERVICE format group with activity, plus defaults derived from the data (typed into the DEBT input cells)."""
    py = BUDGET_YEAR - 1
    am, bm = list(act.columns[3:15]), list(bud.columns[4:16])
    tot = {}
    for df_, cols, flt in ((act, am, py), (bud, bm, py)):
        d_ = df_[df_["Year"] == flt]
        for a_, v_ in zip(d_["acctnum"], pd.DataFrame({c: pd.to_numeric(d_[c]) for c in cols}).abs().sum(axis=1)):
            tot[a_] = tot.get(a_, 0) + v_
    for a_, v_ in zip(ven["acctnum"], ven[VEN_MONTHS].abs().sum(axis=1)):
        tot[a_] = tot.get(a_, 0) + v_
    accs = coa[coa["FMT_GROUP"] == "DEBT SERVICE"].sort_values("ACCTNO")
    accs = [r_ for r_ in accs.itertuples() if tot.get(r_.ACCTNUM, 0) > 0.004]
    pick = lambda ty: max((r_ for r_ in accs if r_.TYPE == ty), key=lambda r_: tot[r_.ACCTNUM], default=None)
    p_, i_ = pick("LIABILITY"), pick("EXPENSE")
    d = dict(accts=[int(r_.ACCTNO) for r_ in accs], prin=int(p_.ACCTNO) if p_ else None, intr=int(i_.ACCTNO) if i_ else None,
             rate=0.0, lender="")
    if p_ is not None and i_ is not None:
        a_py = act[act["Year"] == py]
        mon = lambda acc, k: float(pd.to_numeric(a_py[a_py["acctnum"] == acc][am[k]]).sum())
        start = float(pd.to_numeric(bal[bal["acctnum"] == p_.ACCTNUM]["starting_balance"]).sum())
        m_ = ACTUALS_THRU
        open_ = -(start + sum(mon(p_.ACCTNUM, k) for k in range(m_ - 1)))
        if open_ > 0 and mon(i_.ACCTNUM, m_ - 1) > 0:
            d["rate"] = round(mon(i_.ACCTNUM, m_ - 1) / (open_ * 30 / 360), 5)
        v_ = ven[ven["acctnum"] == p_.ACCTNUM].copy()
        if not v_.empty:
            v_["t"] = v_[VEN_MONTHS].abs().sum(axis=1)
            d["lender"] = str(v_.sort_values("t", ascending=False).iloc[0]["vendnme1"])
    return d


DEBT_DEF = debt_defaults()
print("DEBT defaults:", DEBT_DEF)


def build_debt_sheet(ws, D):
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDEF", (10, 2, 50, 11, 8, 2)):
        ws.column_dimensions[col].width = w
    for m in range(12):
        ws.column_dimensions[CL(7 + m)].width = 14.5
    ws.column_dimensions["S"].width = 14
    ws.column_dimensions["T"].width = 10
    ws.column_dimensions["U"].width = 7
    dv_yn = DataValidation(type="list", formula1='"Yes,No"', allow_blank=False, showErrorMessage=True)
    dv_pct = DataValidation(type="decimal", operator="between", formula1="0", formula2="1", allow_blank=False, showErrorMessage=True,
                            errorTitle="Percentage", error="Enter a percentage between 0% and 100%.")
    dv_rate = DataValidation(type="decimal", operator="between", formula1="0", formula2="1", allow_blank=False, showErrorMessage=True,
                             errorTitle="Interest rate", error="Enter the annual rate as a percentage between 0% and 100% (e.g. 3.72%).")
    dv_basis = DataValidation(type="list", formula1='"30/360,Actual/360,Actual/365"', allow_blank=False, showErrorMessage=True)
    dv_acct = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True, showErrorMessage=True,
                             errorTitle="GL account", error="Enter a 7-digit GL account number.")
    dv_amt = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True, showErrorMessage=True,
                            errorTitle="Amount", error="Enter a positive amount, or leave blank.")
    for v_ in (dv_yn, dv_pct, dv_rate, dv_basis, dv_acct, dv_amt):
        ws.add_data_validation(v_)
    put(ws, "A1", '=BuildingName&" - DEBT SERVICE"', f_title)
    put(ws, "A2", '=\"BUDGET FOR THE YEAR \"&BudgetYear', f_bold)
    put(ws, "A3", "Primary mortgage: interest = opening balance x rate x days / day-count basis; principal = monthly payment - interest. The opening balance for the budget year is the prior-year ending balance "
                  "(MRI starting balance less principal paid: actual through the actuals-through month, then calculated). Each GL account below shows its prior-year activity, the calculated line, "
                  "an adjustment line and three open lines. Set 'Use calc' to No to ignore the calculated line and type your own amounts.", f_sub)
    put(ws, "A4", '="Prior-year forecast = actuals Jan-"&TEXT(DATE(BudgetYear-1,ActualsThru,1),"mmm")&" plus budget for the remaining months (set on INPUT)."', f_sub)
    for col, t in {"A": "GL Acct", "C": "Account / Line", "D": "GL", "E": "Source"}.items():
        put(ws, f"{col}5", t, f_bold)
    for m in range(12):
        put(ws, f"{CL(7 + m)}5", f"=DATE(BudgetYear,{m + 1},1)", f_bold, "mmm-yy", align="right")
    put(ws, "S5", "Total", f_bold, align="right")
    put(ws, "T5", "$/SF", f_bold, align="right")
    put(ws, "U5", "Block", f_bold)
    for col in range(1, 22):
        ws.cell(5, col).border = Border(bottom=thin)
    EF, EL = EXP_FIRST, EXP_LAST
    for rr, lbl, tag in ((6, "PRIOR-YEAR FORECAST - ALL ACCOUNTS", "FCST"), (7, "BUDGET-YEAR TOTAL - ALL ACCOUNTS (this sheet)", "FINAL")):
        put(ws, f"C{rr}", lbl, f_bold)
        for m in range(12):
            col = CL(7 + m)
            put(ws, f"{col}{rr}", f'=SUMIFS({col}${EF}:{col}${EL},$E${EF}:$E${EL},"{tag}")', f_bold, NUM0)
        put(ws, f"S{rr}", f"=SUM(G{rr}:R{rr})", f_bold, NUM0)
        put(ws, f"T{rr}", f"=IF(BillingSF=0,0,S{rr}/BillingSF)", f_bold, NUM2)

    # ---- panel (rows 9-20)
    put(ws, "C9", "PRIMARY MORTGAGE (inputs)", f_bold)
    for col in range(1, 22):
        ws.cell(9, col).border = Border(bottom=thin)
    put(ws, "C10", "Lender")
    put(ws, "G10", D["lender"] or None, f_bold, inp=True, align="left")
    put(ws, "C11", "Interest expense GL account")
    put(ws, "G11", D["intr"], f_bold, "0", inp=True, align="right")
    put(ws, "I11", "Default = the DEBT SERVICE interest account with the most prior-year activity. The calculated interest goes to this account.", f_note)
    put(ws, "C12", "Mortgage principal (liability) GL account")
    put(ws, "G12", D["prin"], f_bold, "0", inp=True, align="right")
    put(ws, "I12", "Default = the DEBT SERVICE liability account with the most prior-year activity. Its MRI starting balance is the loan balance; the calculated principal goes to this account.", f_note)
    put(ws, "C13", "Interest rate (annual)")
    put(ws, "G13", D["rate"], f_bold, "0.000%", inp=True, align="right")
    put(ws, "I13", "Typed input (MRI does not hold the rate). Default = the rate implied by the latest MRI actual month (next line); change it when the lender resets the rate.", f_note)
    put(ws, "C14", "  Rate implied by the latest MRI actual month (read-only)")
    put(ws, "C15", "Interest day-count basis")
    put(ws, "G15", "30/360", f_bold, inp=True, align="right")
    put(ws, "I15", "30/360 = 30 days over 360 (legacy). Actual/360 and Actual/365 use the calendar days of each month.", f_note)
    put(ws, "C16", "Monthly principal + interest payment - override")
    put(ws, "G16", None, f_bold, NUM2, inp=True, align="right")
    put(ws, "I16", "Leave blank to use the latest MRI actual payment below. Type a payment when the lender changes it.", f_note)
    put(ws, "C17", "  Latest MRI actual payment (read-only)")
    put(ws, "C18", "Monthly payment used", f_bold)
    put(ws, "G18", "=IF(ISNUMBER($G$16),$G$16,$G$17)", f_bold, NUM2, align="right")
    put(ws, "C19", "Loan maturity date (optional)")
    put(ws, "G19", None, f_bold, DATEF, inp=True, align="right")
    put(ws, "I19", '=IF(ISNUMBER($G$19),IF(YEAR($G$19)<=BudgetYear,"LOAN MATURES IN OR BEFORE THE BUDGET YEAR - review balloon payment / refinance",""),"")', f_bold)
    put(ws, "C20", '="Opening balance 1/1/"&BudgetYear&" per lender statement - override (optional)"')
    put(ws, "G20", None, f_bold, NUM2, inp=True, align="right")
    put(ws, "I20", '="Leave blank: the opening balance is the "&(BudgetYear-1)&" ending balance below. Type the lender\'s 12/31 balance if it differs."', f_note)
    dv_acct.add("G11")
    dv_acct.add("G12")
    dv_rate.add("G13")
    dv_basis.add("G15")
    dv_amt.add("G16")
    dv_amt.add("G20")
    DAYDEN = 'IF($G$15="Actual/365",365,360)'
    DAB = lambda c: f"DATA_ACTUALS!{c}$2:{c}$2000"
    # ---- roll-forward (rows 22-35)
    put(ws, "C22", "MORTGAGE ROLL-FORWARD (read-only)", f_bold)
    for col in range(1, 22):
        ws.cell(22, col).border = Border(bottom=thin)
    put(ws, "C23", '=(BudgetYear-1)&" (actual through "&TEXT(DATE(BudgetYear-1,ActualsThru,1),"mmm")&", then calculated)"', f_bold)
    put(ws, "C30", '=BudgetYear&" budget (calculated)"', f_bold)
    for rr, lbl in ((24, "    Interest days"), (25, "    Opening balance"), (26, "    Interest"), (27, "    Principal"), (28, "    Ending balance"),
                    (31, "    Interest days"), (32, "    Opening balance"), (33, "    Interest"), (34, "    Principal"), (35, "    Ending balance")):
        put(ws, f"C{rr}", lbl, f_bold if rr in (28, 35, 25, 32) else f_norm)
    for m in range(12):
        col, ac = CL(7 + m), CL(4 + m)                    # sheet month column; DATA_ACTUALS month column
        prev = CL(6 + m)
        put(ws, f"{col}23", f"=DATE(BudgetYear-1,{m + 1},1)", f_bold, "mmm-yy", align="right")
        put(ws, f"{col}30", f"=DATE(BudgetYear,{m + 1},1)", f_bold, "mmm-yy", align="right")
        put(ws, f"{col}24", f'=IF($G$15="30/360",30,DAY(DATE(BudgetYear-1,{m + 2},0)))', f_norm, "0")
        put(ws, f"{col}31", f'=IF($G$15="30/360",30,DAY(DATE(BudgetYear,{m + 2},0)))', f_norm, "0")
        if m == 0:
            put(ws, f"{col}25", '=-SUMIFS(DATA_BALANCES!$D$2:$D$500,DATA_BALANCES!$C$2:$C$500,"MR"&$G$12,DATA_BALANCES!$B$2:$B$500,BudgetYear-1)', f_bold, NUM2)
            put(ws, f"{col}32", f"=IF(ISNUMBER($G$20),$G$20,$R$28)", f_bold, NUM2)
        else:
            put(ws, f"{col}25", f"={prev}28", f_bold, NUM2)
            put(ws, f"{col}32", f"={prev}35", f_bold, NUM2)
        act_ = lambda acc: f'SUMIFS({DAB(ac)},DATA_ACTUALS!$C$2:$C$2000,"MR"&{acc},DATA_ACTUALS!$B$2:$B$2000,BudgetYear-1)'
        put(ws, f"{col}26", f"=IF(MONTH({col}$23)<=ActualsThru,{act_('$G$11')},ROUND({col}25*$G$13*{col}24/{DAYDEN},2))", f_norm, NUM2)
        put(ws, f"{col}27", f"=IF(MONTH({col}$23)<=ActualsThru,{act_('$G$12')},MAX(0,MIN($G$18-{col}26,{col}25)))", f_norm, NUM2)
        put(ws, f"{col}28", f"={col}25-{col}27", f_bold, NUM2)
        put(ws, f"{col}33", f"=ROUND({col}32*$G$13*{col}31/{DAYDEN},2)", f_norm, NUM2)
        put(ws, f"{col}34", f"=MAX(0,MIN($G$18-{col}33,{col}32))", f_norm, NUM2)
        put(ws, f"{col}35", f"={col}32-{col}34", f_bold, NUM2)
    for rr in (26, 27, 33, 34):
        put(ws, f"S{rr}", f"=SUM(G{rr}:R{rr})", f_bold, NUM2)
    put(ws, "T35", '=IF(ISNUMBER($G$20),IF(ABS($G$20-$R$28)>=0.005,"Override differs from "&(BudgetYear-1)&" ending by "&TEXT($G$20-$R$28,"#,##0.00"),""),"")', f_note)
    # read-only panel lines that read the roll-forward / MRI data
    put(ws, "G14", f"=IF(OR(ActualsThru<1,ActualsThru>12),0,IF(INDEX($G$25:$R$25,ActualsThru)<=0,0,ROUND(INDEX($G$26:$R$26,ActualsThru)/(INDEX($G$25:$R$25,ActualsThru)*INDEX($G$24:$R$24,ActualsThru)/{DAYDEN}),5)))", f_norm, "0.000%", align="right")
    pay = ('SUMPRODUCT((DATA_ACTUALS!$C$2:$C$2000="MR"&{a})*(DATA_ACTUALS!$B$2:$B$2000=BudgetYear-1)*(DATA_ACTUALS!$D$2:$O$2000)*((COLUMN(DATA_ACTUALS!$D$2:$O$2000)-3)=ActualsThru))')
    put(ws, "G17", "=" + pay.format(a="$G$11") + "+" + pay.format(a="$G$12"), f_norm, NUM2, align="right")
    put(ws, "I14", "Interest for the latest actual month / (opening balance x days / basis). Compare with the rate typed above.", f_note)
    put(ws, "I17", "Principal + interest paid in the actuals-through month (INPUT).", f_note)
    # ---- account blocks
    r0 = 37
    for k, acctno in enumerate(D["accts"], 1):
        A = f"$A${r0}"
        ra, rb, rf, rc, ri = r0 + 1, r0 + 2, r0 + 3, r0 + 4, r0 + 5
        put(ws, f"U{r0}", k, f_lbl)
        put(ws, f"A{r0}", acctno, f_bold, "0000\\-000", align="left")
        put(ws, f"C{r0}", f'=IFERROR(INDEX(COA_NAME,MATCH({A},COA_ACCTNO,0)),"(not in chart of accounts)")', f_bold)
        for col in range(1, 22):
            ws.cell(r0, col).border = Border(bottom=thin)
        put(ws, f"G{r0}", f'=IF(OR({A}=$G$11,{A}=$G$12),"Calc: see panel","Increase %")', f_lbl, align="right")
        put(ws, f"H{r0}", 0.0, f_bold, "0.00%", inp=True, align="right")
        dv_pct.add(f"H{r0}")
        put(ws, f"I{r0}", "Use calc?", f_lbl, align="right")
        put(ws, f"J{r0}", "Yes", f_bold, inp=True, align="center")
        dv_yn.add(f"J{r0}")
        put(ws, f"C{ra}", f'="  "&(BudgetYear-1)&" actual (Jan-"&TEXT(DATE(BudgetYear-1,ActualsThru,1),"mmm")&")"')
        put(ws, f"C{rb}", f'=IF(ActualsThru>=12,"  "&(BudgetYear-1)&" budget (none - full year actual)","  "&(BudgetYear-1)&" budget ("&TEXT(DATE(BudgetYear-1,ActualsThru+1,1),"mmm")&"-Dec)")')
        put(ws, f"C{rf}", f'="  "&(BudgetYear-1)&" forecast"', f_bold)
        for rowi, tag in ((ra, "ACT"), (rb, "BUD"), (rf, "FCST")):
            put(ws, f"E{rowi}", tag, f_note)
        put(ws, f"C{rc}", f'="  "&BudgetYear&IF(OR({A}=$G$11,{A}=$G$12)," from the mortgage roll-forward (calculated)"," prior-year forecast x (1 + increase) (calculated)")')
        put(ws, f"E{rc}", "CALC", f_note)
        for m in range(12):
            col, ac, bc = CL(7 + m), CL(4 + m), CL(5 + m)
            ws[f"{col}{ra}"] = f'=IF(MONTH({col}$5)<=ActualsThru,SUMIFS({DA}{ac}$2:{ac}$2000,{DA}$C$2:$C$2000,"MR"&{A},{DA}$B$2:$B$2000,BudgetYear-1),0)'
            ws[f"{col}{rb}"] = f'=IF(MONTH({col}$5)>ActualsThru,SUMIFS({DB}{bc}$2:{bc}$2000,{DB}$D$2:$D$2000,"MR"&{A},{DB}$B$2:$B$2000,BudgetYear-1),0)'
            ws[f"{col}{rf}"] = f"={col}{ra}+{col}{rb}"
            ws[f"{col}{rc}"] = f'=IF($J${r0}="No",0,IF({A}=$G$11,{col}$33,IF({A}=$G$12,{col}$34,{col}{rf}*(1+$H${r0}))))'
        for rowi, bold in ((ra, False), (rb, False), (rf, True), (rc, False)):
            for m in range(12):
                c = ws.cell(rowi, 7 + m)
                c.number_format = NUM0
                c.font = f_bold if bold else f_norm
                c.border = Border(bottom=hair)
            put(ws, f"S{rowi}", f"=SUM(G{rowi}:R{rowi})", f_bold if bold else f_norm, NUM0)
            put(ws, f"T{rowi}", f"=IF(BillingSF=0,0,S{rowi}/BillingSF)", f_bold if bold else f_norm, NUM2)
        r = ri
        put(ws, f"C{r}", f'="    "&BudgetYear&" adjustment (+/-)"', f_note)
        put(ws, f"D{r}", f"={A}", fmt="0")
        put(ws, f"E{r}", "INPUT", f_note)
        in_rows = [r] + [r + 1, r + 2, r + 3]
        for x in in_rows[1:]:
            put(ws, f"C{x}", None, inp=True)
            put(ws, f"D{x}", f"={A}", fmt="0")
            put(ws, f"E{x}", "INPUT", f_note)
        for x in in_rows:
            for m in range(12):
                c = ws.cell(x, 7 + m)
                c.fill = INPUT_FILL
                c.protection = UNLOCK
                c.number_format = NUM0
                c.font = f_norm
                c.border = Border(bottom=hair)
            put(ws, f"S{x}", f"=SUM(G{x}:R{x})", fmt=NUM0_BZ)
            put(ws, f"T{x}", f"=IF(BillingSF=0,0,S{x}/BillingSF)", fmt=NUM2_BZ)
        rt = r + 4
        put(ws, f"C{rt}", f'="  "&BudgetYear&" total"', f_bold)
        put(ws, f"D{rt}", f"={A}", f_note, "0")
        put(ws, f"E{rt}", "FINAL", f_note)
        for m in range(12):
            col = CL(7 + m)
            put(ws, f"{col}{rt}", f'={col}{rc}+SUMIFS({col}{r0 + 1}:{col}{rt - 1},$E{r0 + 1}:$E{rt - 1},"INPUT")', f_bold, NUM0)
        put(ws, f"S{rt}", f"=SUM(G{rt}:R{rt})", f_bold, NUM0)
        put(ws, f"T{rt}", f"=IF(BillingSF=0,0,S{rt}/BillingSF)", f_bold, NUM2)
        for col in range(1, 22):
            ws.cell(rt, col).border = Border(top=thin)
        r0 = rt + 3
    assert r0 < EXP_LAST, "DEBT needs more than %d rows" % EXP_LAST
    ws.freeze_panes = "D6"
    ws.protection.sheet = True
    ws.protection.formatColumns = False
    ws.protection.formatRows = False
    return r0


_end = build_debt_sheet(ws_debt, DEBT_DEF)
print(f"DEBT: {len(DEBT_DEF['accts'])} account blocks, rows to {_end}")

# ------------------------------------------------------------------ TAXES (input: real estate tax, escrow, other tax accounts)
TAX_GROUP = "TAXES"
TAX_INC_DEFAULT = 0.025          # default Increase % for every tax account except the main one (was the BIA increase)
TAX_PANEL_END = 18
TAX_BLOCK_ROWS = 11              # header, ACT, BUD, FCST, CALC, 3 INPUT, FINAL, OFFSET, NET
TAX_LIST_COL, TAX_FCST_COL, TAX_LIST_ROW = "W", "X", 7   # hidden helper list for the INPUT dropdown (account, prior-year forecast)


def tax_accounts():
    """TAXES-group accounts with prior-year actual, budget or vendor activity (fixed when the workbook is built).
    Returns (expense-region accounts in account order, escrow account or None, default main tax account or None)."""
    py = BUDGET_YEAR - 1
    am, bm = list(act.columns[3:15]), list(bud.columns[4:16])
    tot = {}
    for df_, cols, flt in ((act, am, py), (bud, bm, py)):
        d_ = df_[df_["Year"] == flt]
        for a_, v_ in zip(d_["acctnum"], pd.DataFrame({c: pd.to_numeric(d_[c]) for c in cols}).abs().sum(axis=1)):
            tot[a_] = tot.get(a_, 0) + v_
    for a_, v_ in zip(ven["acctnum"], ven[VEN_MONTHS].abs().sum(axis=1)):
        tot[a_] = tot.get(a_, 0) + v_
    accs = [r_ for r_ in coa[coa["FMT_GROUP"] == TAX_GROUP].sort_values("ACCTNO").itertuples() if tot.get(r_.ACCTNUM, 0) > 0.004]
    non_exp = [r_ for r_ in accs if r_.TYPE != "EXPENSE"]
    esc = max(non_exp, key=lambda r_: tot[r_.ACCTNUM]) if non_exp else None          # escrow = the balance-sheet account with most activity
    exp_ = [r_ for r_ in accs if r_.TYPE == "EXPENSE"]
    main = next((r_ for r_ in exp_ if int(r_.ACCTNO) == 5710000), None) or (max(exp_, key=lambda r_: tot[r_.ACCTNUM]) if exp_ else None)
    region = [(int(r_.ACCTNO), r_.TYPE == "EXPENSE") for r_ in accs if esc is None or r_.ACCTNUM != esc.ACCTNUM]
    return region, (int(esc.ACCTNO) if esc is not None else None), (int(main.ACCTNO) if main is not None else None)


def build_tax_sheet(ws):
    region, esc_acct, main_def = tax_accounts()
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDEF", (10, 2, 46, 11, 8, 2)):
        ws.column_dimensions[col].width = w
    for m in range(12):
        ws.column_dimensions[CL(7 + m)].width = 11
    ws.column_dimensions["S"].width = 13
    ws.column_dimensions["T"].width = 10
    ws.column_dimensions["U"].width = 7
    put(ws, "A1", '=BuildingName&" - TAXES"', f_title)
    put(ws, "A2", '="BUDGET FOR THE YEAR "&BudgetYear', f_bold)
    put(ws, "A3", "Real estate tax payments, escrow and other tax accounts. Light-blue cells are inputs. Each account shows its prior-year actual/budget/forecast, a calculated line "
                  "(main tax account: the quarterly payments below; escrow: deposit less tax paid; every other account: prior-year forecast x (1 + Increase %), switched off with Use calc = No) "
                  "and three light-blue lines for amounts you add. Expenses are positive numbers. The main real estate tax account is chosen on INPUT.", f_sub)
    put(ws, "A4", '="Prior-year forecast = actuals Jan-"&TEXT(DATE(BudgetYear-1,ActualsThru,1),"mmm")&" plus budget for the remaining months (set on INPUT)."', f_sub)
    for col, t in {"A": "GL Acct", "C": "Account / Line", "D": "GL", "E": "Source"}.items():
        put(ws, f"{col}5", t, f_bold)
    for m in range(12):
        put(ws, f"{CL(7 + m)}5", f"=DATE(BudgetYear,{m + 1},1)", f_bold, "mmm-yy", align="right")
    put(ws, "S5", "Total", f_bold, align="right")
    put(ws, "T5", "$/SF", f_bold, align="right")
    for col in range(1, 22):
        ws.cell(5, col).border = Border(bottom=thin)
    # ---- block layout: every account except escrow first (account order), escrow LAST so its formula can read the
    #      other blocks without a range that contains itself (Excel flags such ranges as circular)
    order = region + ([(esc_acct, False)] if esc_acct else [])
    pos, r = {}, TAX_PANEL_END + 2
    for a_, is_exp in order:
        pos[a_] = dict(r0=r, ra=r + 1, rb=r + 2, rf=r + 3, rc=r + 4, i1=r + 5, rt=r + 8, exp=is_exp, esc=(a_ == esc_acct))
        r += TAX_BLOCK_ROWS + 1
    top = TAX_PANEL_END + 2
    bot = (pos[region[-1][0]]["rt"] + 2) if region else top          # last row of the non-escrow region
    # ---- helper list (hidden column W:X): expense accounts that can be the main tax account, with their prior-year forecast
    exp_list = [a_ for a_, e_ in region if e_]
    put(ws, f"{TAX_LIST_COL}{TAX_LIST_ROW - 1}", "Main-account list", f_lbl)
    for k_, a_ in enumerate(exp_list):
        put(ws, f"{TAX_LIST_COL}{TAX_LIST_ROW + k_}", a_, f_norm, "0")
        put(ws, f"{TAX_FCST_COL}{TAX_LIST_ROW + k_}", f"=S{pos[a_]['rf']}", f_norm, NUM2)
    last_ = TAX_LIST_ROW + max(len(exp_list), 1) - 1
    name("TAX_MAIN_LIST", f"TAXES!${TAX_LIST_COL}${TAX_LIST_ROW}:${TAX_LIST_COL}${last_}")
    ws.column_dimensions[TAX_LIST_COL].hidden = True
    ws.column_dimensions[TAX_FCST_COL].hidden = True
    MAIN_FCST = f"IFERROR(INDEX(${TAX_FCST_COL}${TAX_LIST_ROW}:${TAX_FCST_COL}${last_},MATCH(TaxMain,${TAX_LIST_COL}${TAX_LIST_ROW}:${TAX_LIST_COL}${last_},0)),0)"
    # ---- input panel
    put(ws, "C6", "REAL ESTATE TAX ESTIMATE AND PAYMENTS", f_bold)
    put(ws, "I6", '="Main real estate tax account (set on INPUT): "&IF(COUNTIF(TAX_MAIN_LIST,TaxMain)=0,"NONE - choose one on INPUT",TaxMain&" "&IFERROR(INDEX(COA_NAME,MATCH(TaxMain,COA_ACCTNO,0)),""))', f_bold)
    for col in range(1, 22):
        ws.cell(6, col).border = Border(bottom=thin)
    rows = [
        (7, "Prior-year assessed value (optional)", None, "#,##0", True, "Optional. With the tax rate below it gives the prior-year tax; leave both blank to use the prior-year GL forecast instead."),
        (8, "Prior-year tax rate (optional)", None, "0.00000%", True, "Tax rate on the prior-year bill (e.g. 2.531%)."),
        (9, "Prior-year annual tax (used)", f"=IF(AND(N(G7)>0,N(G8)>0),G7*G8,{MAIN_FCST})", NUM2, False, "Assessed value x rate when both are entered, otherwise the prior-year forecast of the main tax account (actuals + budget for the rest of the year)."),
        (10, "Assumed increase for the budget year", 0.035, "0.0%", True, "Assumption - change as needed."),
        (11, "Budget-year annual tax", "=G9*(1+G10)", NUM2, False, "Prior-year tax x (1 + increase)."),
    ]
    for rr, lbl, val, fmt, is_inp, note in rows:
        put(ws, f"C{rr}", lbl, f_bold if rr in (9, 11) else f_norm)
        put(ws, f"G{rr}", val, f_bold if not is_inp else f_norm, fmt, inp=is_inp, align="right")
        put(ws, f"I{rr}", note, f_note)
    pay = [(12, "1st payment (prior-year tax / 4)", 2, "=G9/4"), (13, "2nd payment (prior-year tax / 4)", 5, "=G9/4"),
           (14, "3rd payment (true-up: remaining tax / 2)", 8, "=(G11-H12-H13)/2"), (15, "4th payment (true-up: remaining tax / 2)", 11, "=(G11-H12-H13)/2")]
    put(ws, "G6", "Pay month", f_lbl, align="right")
    put(ws, "H6", "Amount", f_lbl, align="right")
    dv_m = DataValidation(type="whole", operator="between", formula1="1", formula2="12", allow_blank=False, showErrorMessage=True,
                          errorTitle="Payment month", error="Enter a month number from 1 to 12.")
    dv_pct = DataValidation(type="decimal", operator="between", formula1="-1", formula2="1", allow_blank=False, showErrorMessage=True,
                            errorTitle="Increase %", error="Enter a percentage between -100% and 100%.")
    dv_yn = DataValidation(type="list", formula1='"Yes,No"', allow_blank=False, showErrorMessage=True)
    for v_ in (dv_m, dv_pct, dv_yn):
        ws.add_data_validation(v_)
    for rr, lbl, mo, fml in pay:
        put(ws, f"C{rr}", lbl)
        put(ws, f"G{rr}", mo, f_norm, "0", inp=True, align="right")
        put(ws, f"H{rr}", fml, f_norm, NUM2, align="right")
        dv_m.add(f"G{rr}")
    put(ws, "I12", "Month numbers: 1 = Jan ... 12 = Dec. Payments are posted to the main real estate tax account (INPUT).", f_note)
    put(ws, "C16", "Monthly escrow deposit", f_norm)
    if esc_acct:
        put(ws, "G16", f"=MAX(0,MAX(G{pos[esc_acct]['ra']}:R{pos[esc_acct]['ra']}))", f_norm, NUM2, inp=True, align="right")
        put(ws, "I16", f"Default = the largest monthly posting to escrow account {esc_acct} this year. Type over it with the amount from the last mortgage statement; 0 = taxes are not paid from escrow.", f_note)
    else:
        put(ws, "G16", 0, f_norm, NUM2, inp=True, align="right")
        put(ws, "I16", "This building has no real estate tax escrow account with activity, so no escrow line is calculated.", f_note)
    put(ws, "C17", "Default increase for the other tax accounts", f_norm)
    put(ws, "G17", TAX_INC_DEFAULT, f_norm, "0.0%", inp=True, align="right")
    dv_pct.add("G17")
    put(ws, "I17", "Copied into the Increase % of every account below except the main tax account; type over it there to change one account.", f_note)
    put(ws, "C18", "Total real estate tax (calculated payments)", f_bold)
    put(ws, "G18", "=SUM(H12:H15)", f_bold, NUM2, align="right")
    put(ws, "I18", "Check: equals the budget-year annual tax.", f_note)

    def calc_cells(r, bold=False):
        for m in range(12):
            c = ws.cell(r, 7 + m)
            c.number_format = NUM0
            c.font = f_bold if bold else f_norm
            c.border = Border(bottom=hair)

    for a_, _ in order:
        p = pos[a_]; r0, ra, rb, rf, rc, i1, rt = p["r0"], p["ra"], p["rb"], p["rf"], p["rc"], p["i1"], p["rt"]
        A = f"$A${r0}"
        put(ws, f"A{r0}", a_, f_bold, "0000\\-000", align="left")
        put(ws, f"C{r0}", f'=IFERROR(INDEX(COA_NAME,MATCH({A},COA_ACCTNO,0)),"(not in chart of accounts)")', f_bold)
        for col in range(1, 22):
            ws.cell(r0, col).border = Border(bottom=thin)
        if not p["esc"]:
            put(ws, f"G{r0}", "Increase %", f_lbl, align="right")
            put(ws, f"H{r0}", "=$G$17", f_bold, "0.00%", inp=True, align="right")
            dv_pct.add(f"H{r0}")
            put(ws, f"I{r0}", "Use calc?", f_lbl, align="right")
            put(ws, f"J{r0}", "Yes", f_bold, inp=True, align="center")
            dv_yn.add(f"J{r0}")
            put(ws, f"K{r0}", f'=IF({A}=TaxMain,"MAIN TAX ACCOUNT - calculated line = quarterly payments above (Increase % and Use calc not used)","")', f_note)
        put(ws, f"C{ra}", f'="  "&(BudgetYear-1)&" actual (Jan-"&TEXT(DATE(BudgetYear-1,ActualsThru,1),"mmm")&")"')
        put(ws, f"C{rb}", f'=IF(ActualsThru>=12,"  "&(BudgetYear-1)&" budget (none - full year actual)","  "&(BudgetYear-1)&" budget ("&TEXT(DATE(BudgetYear-1,ActualsThru+1,1),"mmm")&"-Dec)")')
        put(ws, f"C{rf}", f'="  "&(BudgetYear-1)&" forecast"', f_bold)
        for rowi, tag in ((ra, "ACT"), (rb, "BUD"), (rf, "FCST"), (rc, "CALC")):
            put(ws, f"E{rowi}", tag, f_note)
        if p["esc"]:
            put(ws, f"C{rc}", '="  "&BudgetYear&" escrow deposit less tax paid from escrow (calculated)"', f_norm)
        else:
            put(ws, f"C{rc}", f'=IF({A}=TaxMain,"  "&BudgetYear&" quarterly payments (calculated - main tax account)",'
                              f'"  "&BudgetYear&" prior-year forecast x (1 + increase) (calculated)")', f_norm)
        for m in range(12):
            col = CL(7 + m); ac, bc = CL(4 + m), CL(5 + m)
            ws[f"{col}{ra}"] = f'=IF(MONTH({col}$5)<=ActualsThru,SUMIFS({DA}{ac}$2:{ac}$2000,{DA}$C$2:$C$2000,"MR"&{A},{DA}$B$2:$B$2000,BudgetYear-1),0)'
            ws[f"{col}{rb}"] = f'=IF(MONTH({col}$5)>ActualsThru,SUMIFS({DB}{bc}$2:{bc}$2000,{DB}$D$2:$D$2000,"MR"&{A},{DB}$B$2:$B$2000,BudgetYear-1),0)'
            ws[f"{col}{rf}"] = f"={col}{ra}+{col}{rb}"
            if p["esc"]:
                main_m = f'SUMIFS({col}${top}:{col}${bot},$D${top}:$D${bot},TaxMain,$E${top}:$E${bot},"FINAL")'
                ws[f"{col}{rc}"] = f"=IF($G$16=0,0,$G$16-{main_m})"
            else:
                paym = "+".join(f"(MONTH({col}$5)=$G${q})*$H${q}" for q in (12, 13, 14, 15))
                ws[f"{col}{rc}"] = f'=IF({A}=TaxMain,{paym},IF($J${r0}="Yes",{col}{rf}*(1+$H${r0}),0))'
        for rowi, bold in ((ra, False), (rb, False), (rf, True), (rc, False)):
            calc_cells(rowi, bold)
            put(ws, f"S{rowi}", f"=SUM(G{rowi}:R{rowi})", f_bold if bold else f_norm, NUM0)
            put(ws, f"T{rowi}", f"=IF(BillingSF=0,0,S{rowi}/BillingSF)", f_bold if bold else f_norm, NUM2)
        for k_ in range(3):
            rr = i1 + k_
            put(ws, f"C{rr}", None, inp=True)
            put(ws, f"D{rr}", f"={A}", fmt="0")
            put(ws, f"E{rr}", "INPUT", f_note)
            for m in range(12):
                c = ws.cell(rr, 7 + m)
                c.fill = INPUT_FILL; c.protection = UNLOCK; c.number_format = NUM0; c.font = f_norm; c.border = Border(bottom=hair)
            put(ws, f"S{rr}", f"=SUM(G{rr}:R{rr})", fmt=NUM0_BZ)
            put(ws, f"T{rr}", f"=IF(BillingSF=0,0,S{rr}/BillingSF)", fmt=NUM2_BZ)
        put(ws, f"C{rt}", f'="  "&BudgetYear&" total"', f_bold)
        put(ws, f"D{rt}", f"={A}", f_note, "0")
        put(ws, f"E{rt}", "FINAL", f_note)
        for m in range(12):
            col = CL(7 + m)
            put(ws, f"{col}{rt}", f"={col}{rc}+SUM({col}{i1}:{col}{i1 + 2})", f_bold, NUM0)
        put(ws, f"S{rt}", f"=SUM(G{rt}:R{rt})", f_bold, NUM0)
        put(ws, f"T{rt}", f"=IF(BillingSF=0,0,S{rt}/BillingSF)", f_bold, NUM2)
        for col in range(1, 22):
            ws.cell(rt, col).border = Border(top=thin)
        ro, rn = rt + 1, rt + 2
        put(ws, f"C{ro}", f'="  "&BudgetYear&" tenant-billed offset from REVENUE (read-only)"', f_note)
        put(ws, f"E{ro}", "OFFSET", f_note)
        put(ws, f"C{rn}", f'="  "&BudgetYear&" net of tenant-billed offset (as shown on DETAIL)"', f_bold)
        put(ws, f"E{rn}", "NET", f_note)
        for m in range(12):
            col = CL(7 + m)
            ws[f"{col}{ro}"] = f'=-SUMIFS(REVENUE!{col}$8:{col}$1500,REVENUE!$D$8:$D$1500,{A},REVENUE!$V$8:$V$1500,"<>IGNORED*",REVENUE!$E$8:$E$1500,"<>COST")'
            ws[f"{col}{rn}"] = f"={col}{rt}+{col}{ro}"
        calc_cells(ro)
        calc_cells(rn, True)
        put(ws, f"S{ro}", f"=SUM(G{ro}:R{ro})", f_norm, NUM0)
        put(ws, f"T{ro}", f"=IF(BillingSF=0,0,S{ro}/BillingSF)", f_norm, NUM2)
        put(ws, f"S{rn}", f"=SUM(G{rn}:R{rn})", f_bold, NUM0)
        put(ws, f"T{rn}", f"=IF(BillingSF=0,0,S{rn}/BillingSF)", f_bold, NUM2)
    ws.freeze_panes = "D6"
    ws.protection.sheet = True
    ws.protection.formatColumns = False
    ws.protection.formatRows = False
    return pos, main_def, esc_acct


TAX_POS, TAX_MAIN_DEFAULT, TAX_ESCROW = build_tax_sheet(ws_tax)
print("TAXES: %d account blocks %s, main tax account %s, escrow %s, rows to %d" % (
    len(TAX_POS), sorted(TAX_POS), TAX_MAIN_DEFAULT, TAX_ESCROW, max([p["rt"] for p in TAX_POS.values()] + [20]) + 2))

# INPUT: main real estate tax account (rows 63-64; legend moved to row 66)
put(ws_in, "B63", "REAL ESTATE TAXES", f_bold)
for c_ in "BCDEF":
    ws_in[f"{c_}63"].border = Border(bottom=thin)
put(ws_in, "B64", "Main real estate tax account", f_bold)
put(ws_in, "C64", TAX_MAIN_DEFAULT, f_bold, "0", inp=True, align="right")
put(ws_in, "D64", '=IF(C64="","",IFERROR(INDEX(COA_NAME,MATCH(C64,COA_ACCTNO,0)),""))', f_norm)
put(ws_in, "E65", "The account that receives the quarterly real estate tax payments calculated on TAXES (assessed value, rate, increase, pay months); "
                  "escrow is calculated against it. Every other tax account is budgeted at its prior-year forecast x (1 + Increase %). "
                  "Default = 5710000 Real Estate Taxes, or the tax expense account with the most activity.", f_note)
dv_tax = DataValidation(type="list", formula1="=TAX_MAIN_LIST", allow_blank=True, showErrorMessage=True,
                        errorTitle="Main tax account", error="Pick a tax expense account that has a block on TAXES.")
ws_in.add_data_validation(dv_tax)
dv_tax.add("C64")
name("TaxMain", "INPUT!$C$64")

# ------------------------------------------------------------------ DETAIL / QTRLY skeletons (macro builds the content)
for w, widths in ((ws_det, (11, 46, 4) + (11,) * 12 + (13, 9, 12, 12, 12, 12)),
                  (ws_pf, (10, 44, 13, 13, 13, 13, 11, 11, 11, 11)),
                  (ws_q, (11, 46, 4, 13, 13, 13, 13, 14, 10))):
    w.sheet_view.showGridLines = False
    for i, wd in enumerate(widths, 1):
        w.column_dimensions[CL(i)].width = wd
    w.freeze_panes = "C6" if w is ws_pf else "D6"
    put(w, "A1", "Not built yet - click Build Schedules on the INPUT tab.", f_bold)

# ------------------------------------------------------------------ TEST FILL (scratch copy only)
if TEST:                                       # scratch copy: one deal line filled in on LEASE_DEALS
    _r = DL_FIRST
    for _k, _v in (("suite", "1700"), ("tenant", "TEST TENANT"), ("type", "New lease"), ("sf", 3000), ("start", dt.datetime(BUDGET_YEAR, 3, 1)),
                   ("end", dt.datetime(BUDGET_YEAR + 2, 2, 28)), ("rent", 30), ("incr", 0.03), ("f1m", 1), ("f1n", 2)):
        ws_ld[f"{DLC[_k]}{_r}"] = _v

# ------------------------------------------------------------------ COVER (building, budget year, photo, linked index of all worksheets)
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.drawing.image import Image as XLImage
import io
NAVY, MID, LIGHT = "1F497D", "4F81BD", "DCE6F1"
SHEET_INFO = {
    "INPUT": ("Budget year, building, billing SF, lease options and the Build Schedules button", "Input"),
    "REVENUE": ("Tenant revenue by lease: MRI charges, your adjustments and leasing costs; lease deals summarised from LEASE_DEALS", "Input"),
    "LEASE_DEALS": ("One line per new lease, renewal or expansion: rent, free rent, this building's charge codes, leasing costs and commissions", "Input"),
    "OTHER_REVENUE": ("Non-tenant income by GL account, plus blocks for any other income account", "Input"),
    "EXP_UTILITIES": ("Utilities expenses by GL account and vendor", "Input"),
    "EXP_JANITORIAL": ("Janitorial and cleaning expenses by GL account and vendor", "Input"),
    "EXP_ELEVATOR": ("Elevator expenses by GL account and vendor", "Input"),
    "EXP_HVAC": ("HVAC and payroll expenses by GL account and vendor", "Input"),
    "EXP_LANDSCAPING": ("Landscaping and grounds expenses by GL account and vendor", "Input"),
    "EXP_GENERAL_RM": ("General repairs and maintenance by GL account and vendor", "Input"),
    "TAXES": ("Real estate taxes, escrow deposits and assessments", "Input"),
    "INSURANCE": ("Insurance and insurance consulting", "Input"),
    "PROF_FEES": ("Management fee, accounting fees, bank charges and other professional fees/G&A", "Input"),
    "OTHER_PROF_FEES": ("Other professional fees by GL account: prior-year forecast + increase %, monthly overrides and adjustments", "Input"),
    "REPAIRS_CAPITAL": ("Major repairs by GL account and capital improvements: cost, start month or monthly amounts", "Input"),
    "DEBT": ("Primary mortgage roll-forward (interest and principal) and debt service accounts", "Input"),
    "DETAIL": ("Budget by GL account and month, with section totals and N.O.I.", "Schedule"),
    "PROFORMA": ("Budget year vs prior-year forecast and prior-year budget by GL account: variance, % and $/SF change", "Schedule"),
    "QTRLY SUM": ("Budget by GL account and quarter", "Schedule"),
    "TENANT_SCHEDULE": ("Calculation: MRI tenant charges by month for the budget year", "MRI data / calculation"),
    "DEAL_SCHEDULE": ("Calculation: monthly amounts for each LEASE_DEALS line", "MRI data / calculation"),
    "INCOME_ACCOUNTS": ("Calculation: income accounts that get a block on OTHER_REVENUE", "MRI data / calculation"),
    "EXPENSE_ACCOUNTS": ("Calculation: expense accounts that need a block on the input sheets", "MRI data / calculation"),
    "DATA_COA": ("MRI chart of accounts with Financial Format groups", "MRI data / calculation"),
    "DATA_CHARGECODES": ("MRI charge codes and the income accounts they post to", "MRI data / calculation"),
    "DATA_BUDGET_PY": ("MRI prior-year budget for this building", "MRI data / calculation"),
    "DATA_ACTUALS": ("MRI monthly actuals for this building", "MRI data / calculation"),
    "DATA_BALANCES": ("MRI opening and rolled-forward balances", "MRI data / calculation"),
    "DATA_SUITES": ("MRI suites and square footage", "MRI data / calculation"),
    "DATA_TENANT_CHARGES": ("MRI tenant charge schedule", "MRI data / calculation"),
    "DATA_VENDOR": ("MRI vendor payments by GL account", "MRI data / calculation"),
}
GROUPS = [("INPUT SHEETS", "Input", ORANGE, "orange tabs - the only sheets you type on"),
          ("BUDGET SCHEDULES", "Schedule", PURPLE, "purple tabs - built by the Build Schedules button on INPUT"),
          ("MRI DATA AND CALCULATIONS", "MRI data / calculation", BLUE, "blue tabs - read-only")]
ws = ws_cov
ws.sheet_view.showGridLines = False
ws.sheet_properties.tabColor = NAVY
for col, w in zip("ABCDEFGHIJKL", (3, 13, 13, 13, 13, 13, 13, 3, 24, 24, 3, 3)):
    ws.column_dimensions[col].width = w
fill = lambda hexc: PatternFill("solid", fgColor=hexc)
ws.row_dimensions[1].height = 8
for r_ in (2, 3):
    ws.row_dimensions[r_].height = 24
ws.row_dimensions[4].height = 24
ws.merge_cells("B2:J3")
ws.merge_cells("B4:J4")
put(ws, "B2", "=BuildingName", Font(name=FONT, size=30, bold=True, color="FFFFFF"), align="left")
put(ws, "B4", '="BUDGET YEAR "&BudgetYear', Font(name=FONT, size=16, bold=True, color="FFFFFF"), align="left")
for r_ in (2, 3):
    for c_ in range(2, 11):
        ws.cell(r_, c_).fill = fill(NAVY)
for c_ in range(2, 11):
    ws.cell(4, c_).fill = fill(MID)
ws["B2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws["B4"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
# ---- photo
img_path = ARGS.image
if not img_path:
    for d_ in (UP.rstrip("/"), HERE):
        for ext_ in ("jpg", "jpeg", "png"):
            p_ = os.path.join(d_, f"building_{BLDG_ID}_image.{ext_}")
            if os.path.exists(p_):
                img_path = p_
                break
        if img_path:
            break
PHOTO_ROW, PHOTO_ROWS, PHOTO_W = 6, 19, 482
if img_path:
    from PIL import Image as PILImage
    im_ = PILImage.open(img_path).convert("RGB")
    wpx = min(1100, im_.width)
    im_ = im_.resize((wpx, int(im_.height * wpx / im_.width)))
    buf_ = io.BytesIO()
    im_.save(buf_, format="JPEG", quality=88)
    buf_.seek(0)
    xi = XLImage(buf_)
    xi.width = PHOTO_W
    xi.height = int(PHOTO_W * im_.height / im_.width)
    ws.add_image(xi, f"B{PHOTO_ROW}")
    PHOTO_ROWS = max(PHOTO_ROWS, int(xi.height / 20) + 1)
    print("cover photo:", os.path.basename(img_path), "->", xi.width, "x", xi.height)
else:
    ws.merge_cells(start_row=PHOTO_ROW, start_column=2, end_row=PHOTO_ROW + PHOTO_ROWS - 1, end_column=7)
    put(ws, f"B{PHOTO_ROW}", "BUILDING PHOTO\n(Review > Unprotect Sheet, then Insert > Pictures. Or name a file building_%s_image.jpg and rebuild.)" % BLDG_ID,
        Font(name=FONT, size=11, italic=True, color="7F7F7F"), align="center")
    ws[f"B{PHOTO_ROW}"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r_ in range(PHOTO_ROW, PHOTO_ROW + PHOTO_ROWS):
        for c_ in range(2, 8):
            ws.cell(r_, c_).fill = fill("F2F2F2")
# ---- building facts (right of the photo)
panel = [("Building ID", "=INPUT!C8", "@"), ("Billing SF", "=BillingSF", "#,##0"), ("Schedule status", "=DETAIL!B3", None), ("Schedules last built", '=IF(DETAIL!B4="","",DETAIL!B4)', "yyyy-mm-dd hh:mm")]
rr_ = PHOTO_ROW
put(ws, f"I{rr_}", "BUILDING", Font(name=FONT, size=10, bold=True, color="FFFFFF"))
for c_ in (9, 10):
    ws.cell(rr_, c_).fill = fill(NAVY)
rr_ += 1
for lbl_, fml_, fmt_ in panel:
    put(ws, f"I{rr_}", lbl_, Font(name=FONT, size=9, italic=True, color="7F7F7F"))
    put(ws, f"I{rr_ + 1}", fml_, Font(name=FONT, size=12, bold=True, color=NAVY), fmt_, align="left")
    ws.merge_cells(f"I{rr_ + 1}:J{rr_ + 1}")
    ws[f"I{rr_ + 1}"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws.row_dimensions[rr_ + 1].height = 20 if lbl_ != "Schedule status" else 32
    rr_ += 3
# ---- worksheet index with working links
r = PHOTO_ROW + PHOTO_ROWS + 1
put(ws, f"B{r}", "WORKSHEETS  (click a name to go to that sheet)", Font(name=FONT, size=13, bold=True, color=NAVY))
for c_ in range(2, 11):
    ws.cell(r, c_).border = Border(bottom=Side(style="medium", color=NAVY))
r += 2
link_font = Font(name=FONT, size=11, underline="single", color="0563C1", bold=True)
shown = set()
for title_, kind_, colr_, note_ in GROUPS:
    names_ = [n_ for n_ in wb.sheetnames if n_ != "COVER" and SHEET_INFO.get(n_, ("", "MRI data / calculation"))[1] == kind_]
    if not names_:
        continue
    for c_ in range(2, 11):
        ws.cell(r, c_).fill = fill(colr_)
    put(ws, f"B{r}", title_, Font(name=FONT, size=11, bold=True, color="FFFFFF"))
    put(ws, f"E{r}", note_, Font(name=FONT, size=9, italic=True, color="FFFFFF"))
    r += 1
    for k_, n_ in enumerate(names_):
        desc_ = SHEET_INFO.get(n_, ("", ""))[0]
        c1 = ws.cell(r, 2, n_)
        c1.font = link_font
        c1.hyperlink = Hyperlink(ref=c1.coordinate, location=f"'{n_}'!A1", display=n_)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        put(ws, f"E{r}", desc_, Font(name=FONT, size=10, color="404040"))
        for c_ in range(2, 11):
            ws.cell(r, c_).fill = fill(LIGHT if k_ % 2 == 0 else "FFFFFF")
            ws.cell(r, c_).border = Border(bottom=Side(style="hair", color="BFBFBF"))
        ws.row_dimensions[r].height = 18
        shown.add(n_)
        r += 1
    r += 1
put(ws, f"B{r}", "Tabs: orange = input | purple = budget schedules | blue = MRI reference data. Click Build Schedules on INPUT after changing accounts or MRI data.", f_note)
ws.page_setup.orientation = "portrait"
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 1
ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
ws.protection.sheet = True
ws.protection.objects = False
ws.protection.formatColumns = False
ws.protection.formatRows = False
wb.active = 0

wb.calculation.fullCalcOnLoad = True

# ------------------------------------------------------------------ compact UI: group the reference / vendor-paid rows on the section sheets (Excel row outline)
def add_outline(ws):
    """Group (collapse by default) the 'prior-year' reference rows of every block so each vendor shows one line: the ACT / BUD / FCST rows plus the first VEND row
    sit above the budget (INPUT) row, which becomes the visible summary row; every further VEND row is grouped above its own budget row. Rows are tagged in column E.
    The budget row label becomes '<vendor> - <year> budget' (or '... adjustment (+/-)'). Nothing else moves, and hidden rows still count in every formula."""
    tag = {r: ws.cell(r, 5).value for r in range(1, ws.max_row + 1)}
    grouped, sums = set(), set()
    r = 1
    while r <= ws.max_row:
        if tag.get(r) == "ACT" and tag.get(r + 1) == "BUD" and tag.get(r + 2) == "FCST":
            last = r + 2
            if tag.get(r + 3) == "VEND":
                last = r + 3
            grouped.update(range(r, last + 1))
            sums.add(last + 1)
            r = last + 1
            continue
        if tag.get(r) == "VEND" and r not in grouped:
            grouped.add(r)
            sums.add(r + 1)
        r += 1
    for r in sorted(grouped):
        d = ws.row_dimensions[r]
        d.outlineLevel = 1
        d.hidden = True
    for r in sums:
        ws.row_dimensions[r].collapsed = True
    # vendor name on the budget row that follows a VEND row
    for r in range(1, ws.max_row):
        if tag.get(r) == "VEND" and tag.get(r + 1) == "INPUT":
            c = ws.cell(r + 1, 3)
            old = c.value if isinstance(c.value, str) else ""
            what = "adjustment (+/-)" if "adjustment" in old else "budget"
            c.value = (f'="  "&IFERROR(INDEX(DATA_VENDOR!$C$2:$C$3000,MATCH($D{r},DATA_VENDOR!$B$2:$B$3000,0)),$D{r})&" - "&BudgetYear&" {what}"')
            c.font = f_norm
    ws.sheet_properties.outlinePr.summaryBelow = True
    return len(grouped)


OUTLINE_SHEETS = list(ws_exp.values()) + [ws_tax, ws_ins, ws_pro, ws_opf, ws_debt]
print("outline groups (rows grouped):", {w.title: add_outline(w) for w in OUTLINE_SHEETS})

wb.save(OUT)
print("saved", OUT, "| revenue rows", REV_END, "| code input cells", len(CODE_CELLS))
