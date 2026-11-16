# OTHER_REVENUE User blocks: pick any income account, duplicates, invalid picks, orphan amounts.
import sys
from lo_driver import LO
from make_xlsm import embed_vba, cell
lo = LO()
try:
    doc = lo.load(sys.argv[1]); embed_vba(lo, doc)
    run = lambda: (lo.run_macro(doc, "VBAProject.modBuildAll.BuildAllSchedulesQuiet"), doc.calculateAll())
    S = lambda: cell(doc, "DETAIL", "B3").String
    orv = doc.Sheets.getByName("OTHER_REVENUE"); O = lambda ref: orv.getCellRangeByName(ref)
    det = doc.Sheets.getByName("DETAIL")
    run()
    arr = orv.getCellRangeByPosition(0, 0, 20, 400).getDataArray()
    heads = [(i + 1, r) for i, r in enumerate(arr) if isinstance(r[20], (float, str)) and r[20] not in ("", "Block") and i >= 8 and (isinstance(r[20], float) or str(r[20]).startswith("U"))]
    auto = [(i, r) for i, r in heads if isinstance(r[20], float)]; user = [(i, r) for i, r in heads if isinstance(r[20], str)]
    print("0) built:", S(), "| automatic blocks:", len(auto), [int(r[0]) for _, r in auto], "| user blocks:", len(user), "at rows", [i for i, _ in user])
    print("   first user block banner (row above):", repr(orv.getCellRangeByName(f"A{user[0][0]-1}").String[:60]))
    print("   sheet rows used:", max(i for i, r in enumerate(arr, 1) if any(str(c) != "" for c in r)))
    def ti():
        a = det.getCellRangeByPosition(0, 0, 20, 80).getDataArray()
        return [r[15] for r in a if r[1] == "TOTAL INCOME"][0]
    base = ti(); u1 = user[0][0]; cols = "GHIJKLMNOPQR"
    # 1) pick Gain on Disposal, enter 1,000/month on line 1
    O(f"C{u1}").setString("4910000  Gain on Disposal"); O(f"C{u1+5}").setString("Sale of building")
    for c in cols: O(f"{c}{u1+5}").setValue(1000)
    doc.calculateAll()
    print("1) picked 4910000 ->", "A:", O(f"A{u1}").Value, "| prior-yr forecast row:", O(f"S{u1+3}").String or O(f"S{u1+3}").Value, "| block total:", O(f"S{u1+8}").Value, "| flag:", repr(O(f"T{u1}").String), "| status:", S())
    run()
    arrd = det.getCellRangeByPosition(0, 0, 20, 60).getDataArray()
    r4910 = [r for r in arrd if r[0] == 4910000]
    print("2) after rebuild  ->", S(), "| DETAIL has 4910000 row:", bool(r4910), "total", r4910[0][15] if r4910 else None, "| TOTAL INCOME increase", round(ti() - base), "(expect 12,000)")
    # 3) second user block duplicates an automatic account
    u2 = user[1][0]; O(f"C{u2}").setString("4430000  Interest Income"); doc.calculateAll()
    print("3) duplicate of automatic block 4430000 ->", repr(O(f"T{u2}").String), "|", S())
    O(f"C{u2}").setString("4910000  Gain on Disposal"); doc.calculateAll()
    print("   duplicate of user block 1 ->", repr(O(f"T{u2}").String), "/", repr(O(f"T{u1}").String), "|", S())
    O(f"C{u2}").setString(""); doc.calculateAll(); print("   cleared ->", S())
    # 4) non-income account typed (bypassing the dropdown)
    O(f"C{u2}").setString("5100000  Something"); doc.calculateAll(); print("4) non-income pick ->", repr(O(f"T{u2}").String), "|", S())
    O(f"C{u2}").setString(""); doc.calculateAll()
    # 5) amount with no account selected
    u3 = user[2][0]; O(f"G{u3+5}").setValue(250); doc.calculateAll(); print("5) amount in user block with no account ->", S())
    O(f"G{u3+5}").setValue(0); doc.calculateAll(); print("   cleared ->", S())
    # 6) QTRLY follows
    q = doc.Sheets.getByName("QTRLY SUM").getCellRangeByPosition(0, 0, 8, 60).getDataArray()
    print("6) QTRLY 4910000 total:", [r[7] for r in q if r[0] == 4910000], "| status:", cell(doc, "QTRLY SUM", "B3").String)
    doc.close(True)
finally:
    lo.close()
