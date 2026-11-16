# Compact UI: on the 11 section sheets the ACT / BUD / FCST rows and the VEND rows are grouped (outline level 1, hidden = collapsed) and their budget (summary) row
# carries the collapsed flag and the vendor name; every row the user types in or that shows a total stays visible. Reads the finished .xlsm (no LibreOffice needed).
import sys, openpyxl
F = sys.argv[1]
wb = openpyxl.load_workbook(F, keep_vba=True); wv = openpyxl.load_workbook(F, data_only=True, keep_vba=True)
SHEETS = ["EXP_UTILITIES", "EXP_JANITORIAL", "EXP_ELEVATOR", "EXP_HVAC", "EXP_LANDSCAPING", "EXP_GENERAL_RM", "TAXES", "INSURANCE", "PROF_FEES", "DEBT", "OTHER_PROF_FEES"]
ok = True
for sn in SHEETS:
    ws, wsv = wb[sn], wv[sn]
    ref = {"ACT", "BUD", "FCST", "VEND"}
    n_ref = n_vis_bad = n_grp_bad = n_sum = n_sum_bad = n_lbl = n_lbl_bad = 0
    for r in range(1, ws.max_row + 1):
        t = ws.cell(r, 5).value; d = ws.row_dimensions[r]
        if t in ref:
            n_ref += 1
            if not (d.outlineLevel == 1 and d.hidden): n_grp_bad += 1
        else:
            if d.outlineLevel or d.hidden: n_vis_bad += 1
        if t == "VEND":
            nxt = ws.row_dimensions[r + 1]
            lbl = wsv.cell(r + 1, 3).value
            n_lbl += 1
            if not (isinstance(lbl, str) and "- 2027" in lbl and lbl.strip().split(" - 2027")[0] and ws.cell(r + 1, 5).value == "INPUT"): n_lbl_bad += 1
    sums = [r for r in range(1, ws.max_row + 1) if ws.row_dimensions[r].collapsed]
    # every collapsed (summary) row must directly follow a grouped row and be visible
    for r in sums:
        n_sum += 1
        if not (ws.row_dimensions[r - 1].outlineLevel == 1 and not ws.row_dimensions[r].hidden): n_sum_bad += 1
    good = n_ref > 0 and n_vis_bad == 0 and n_grp_bad == 0 and n_sum_bad == 0 and n_lbl_bad == 0 and n_sum > 0
    ok &= good
    print(f"{sn:16s} reference rows {n_ref:3d} grouped+hidden (bad {n_grp_bad}) | other rows hidden/grouped (bad {n_vis_bad}) | summary rows {n_sum} (bad {n_sum_bad}) | vendor budget labels {n_lbl} (bad {n_lbl_bad}) -> {'OK' if good else 'FAIL'}")
ws = wv["EXP_UTILITIES"]
print("EXP_UTILITIES C14 reads:", repr(ws["C14"].value), "| row 14 hidden:", wb["EXP_UTILITIES"].row_dimensions[14].hidden, "| rows 10-13 hidden:", [wb["EXP_UTILITIES"].row_dimensions[r].hidden for r in range(10, 14)])
print("OVERALL:", "ALL PASS" if ok else "FAILURES ABOVE")
