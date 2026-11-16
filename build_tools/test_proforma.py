# PROFORMA: every account's prior-year forecast and full-year budget vs an independent calculation from the raw MRI extracts,
# budget-year column vs DETAIL, variance signs (+ = favourable) and the change columns, summary lines and row coverage.
# Reads the FINISHED workbook (cached values written by the LibreOffice build); no LibreOffice needed:
#   python3 test_proforma.py FILE.xlsm
import sys, json, os
import openpyxl, pandas as pd
from paths import MRI_DIR, HERE

MON = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
inp = wb["INPUT"]
BLDG, YEAR, THRU, SF = str(inp["C8"].value).strip(), int(inp["C3"].value), int(inp["C15"].value), float(inp["C5"].value)
PY = YEAR - 1


def load(*names):
    f = next(os.path.join(MRI_DIR, n) for n in names if os.path.exists(os.path.join(MRI_DIR, n)))
    d = pd.read_excel(f, sheet_name="MRIQRY", dtype=str)
    d.columns = [c.strip() for c in d.columns]
    d = d[(d.entityid.str.strip() == BLDG) & (d.Year.str.strip() == str(PY))].copy()
    d["a"] = d.acctnum.str.strip()
    for m in MON:
        d[m] = pd.to_numeric(d[m]).fillna(0)
    return d


act = pd.concat([load("FY_2025_ACTUALS.xlsx", "2025_actuals.xlsx"), load("Q1-Q3_2026_ACTUALS.xlsx", "Q1-Q3_ACTUALS.xlsx")])
bud = load("FY_2026_BUDGETS.xlsx", "FY_BUDGET_1.xlsx")
pf, dt = wb["PROFORMA"], wb["DETAIL"]
fails = []
check = lambda ok, msg: (None if ok else fails.append(msg))

print("0) status:", pf["B3"].value, "| DETAIL:", dt["B3"].value)
check(str(pf["B3"].value).startswith(("OK", "CHECK DATA")), "PROFORMA status")
check(wb.sheetnames.index("PROFORMA") == wb.sheetnames.index("DETAIL") + 1, "PROFORMA must follow DETAIL")

det = {int(r[0]): r[15] for r in dt.iter_rows(min_row=6, values_only=True) if isinstance(r[0], (int, float)) and r[0] >= 1e6}
rows, sec, n = [], None, 0
for r in pf.iter_rows(min_row=6, max_col=10, values_only=True):
    if isinstance(r[0], str) and r[2] is None and r[1] is None:
        sec = r[0]
        continue
    rows.append((sec, r))
    if not isinstance(r[0], (int, float)):
        continue
    a, n = int(r[0]), n + 1
    s, key = (-1 if sec == "INCOME" else 1), f"MR{int(r[0])}"
    fc = act[act.a == key][MON[:THRU]].to_numpy().sum() + bud[bud.a == key][MON[THRU:]].to_numpy().sum()
    bb = bud[bud.a == key][MON].to_numpy().sum()
    check(abs(r[3] - s * fc) < 0.01, f"{a} forecast {r[3]} expect {s * fc}")
    check(abs(r[4] - s * bb) < 0.01, f"{a} PY budget {r[4]} expect {s * bb}")
    check(abs(r[2] - det.get(a, 0)) < 0.01, f"{a} budget {r[2]} DETAIL {det.get(a, 0)}")
    var = (r[3] - r[4]) if sec == "INCOME" else (r[4] - r[3])
    check(abs(r[5] - var) < 0.01, f"{a} variance sign")
    g = "n/a" if r[3] == 0 else (r[2] - r[3]) / abs(r[3])
    check(g == r[6] if g == "n/a" else abs(r[6] - g) < 1e-9, f"{a} % change vs forecast")
    check(abs(r[7] - (r[2] - r[3]) / SF) < 1e-6, f"{a} $/SF change vs forecast")
    i_ = "n/a" if r[4] == 0 else (r[2] - r[4]) / abs(r[4])
    check(i_ == r[8] if i_ == "n/a" else abs(r[8] - i_) < 1e-9, f"{a} % change vs budget")
    check(abs(r[9] - (r[2] - r[4]) / SF) < 1e-6, f"{a} $/SF change vs budget")
print(f"1) {n} accounts: prior-year forecast / budget vs raw MRI, budget year vs DETAIL, variance and change columns")

# coverage: every DETAIL account is on PROFORMA; prior-year accounts left out belong only to groups after NET CASH FLOW
on = {int(r[0]) for _, r in rows if isinstance(r[0], (int, float))}
check(set(det) <= on, f"DETAIL accounts missing: {sorted(set(det) - on)}")
fg = {int(k): v.get("group") for k, v in json.load(open(os.path.join(HERE, "fin_groups.json")))["accounts"].items()}
py_accts = {int(a[2:]) for d in (act, bud) for a, v in zip(d.a, d[MON].abs().sum(axis=1)) if v > 0.004}
left = sorted(a for a in py_accts - on if fg.get(a) not in ("DISTRIBUTIONS/(CONTRIBUTIONS)", "CASH BALANCES", "OTHER DISB/(RECEIPTS)", None))
check(not left, f"prior-year accounts with a DETAIL section but not on PROFORMA: {left}")
print("2) coverage: DETAIL accounts all present; prior-year accounts left out only in cash/distribution/other-disb groups or not on the format")

# summary lines: income-like lines favour forecast > budget, cost lines the opposite; NET CASH FLOW is the last line
lab = lambda r: r[0] if isinstance(r[0], str) else r[1]
summ = {lab(r): r for _, r in rows if isinstance(lab(r), str)}
for lbl_, inc in (("TOTAL INCOME", True), ("TOTAL EXPENSES", False), ("TOTAL OPERATING EXPENSES", False),
                  ("N.O.I. BEFORE OTHER PROFESSIONAL FEES & DEBT SERVICE", True), ("NET CASH FLOW", True)):
    r = summ.get(lbl_)
    check(r is not None, f"missing line {lbl_}")
    if r:
        check(abs(r[5] - ((r[3] - r[4]) if inc else (r[4] - r[3]))) < 0.01, f"{lbl_} variance sign")
noi, toe, ti = summ["N.O.I. BEFORE OTHER PROFESSIONAL FEES & DEBT SERVICE"], summ["TOTAL OPERATING EXPENSES"], summ["TOTAL INCOME"]
for c in (2, 3, 4):
    check(abs(noi[c] - (ti[c] - toe[c])) < 0.01, f"N.O.I. column {c}")
dsum = {(r[0] if isinstance(r[0], str) else r[1]): r[15] for r in dt.iter_rows(min_row=6, values_only=True) if isinstance(r[0] if isinstance(r[0], str) else r[1], str)}
check(abs(summ["NET CASH FLOW"][2] - dsum["NET CASH FLOW"]) < 0.01, "NET CASH FLOW budget vs DETAIL")
texts = [str(lab(r)) for _, r in rows if lab(r)]
check(not any("PROJECTED" in t or "DISTRIBUTIONS" in t or "RESERVE" in t for t in texts), "cash/distribution lines must not appear")
print("3) summary lines: N.O.I. = income - operating expenses in all three columns; NET CASH FLOW budget = DETAIL; no cash-balance lines")
print("   TOTAL INCOME", [round(ti[c]) for c in (2, 3, 4, 5)], "| N.O.I.", [round(noi[c]) for c in (2, 3, 4, 5)],
      "| NET CASH FLOW", [round(summ["NET CASH FLOW"][c]) for c in (2, 3, 4, 5)])

print("OVERALL:", "ALL PASS" if not fails else f"{len(fails)} FAIL")
for f in fails[:20]:
    print("  FAIL", f)
sys.exit(1 if fails else 0)
