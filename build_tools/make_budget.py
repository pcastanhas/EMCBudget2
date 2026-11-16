"""One command: MRI extracts -> finished budget workbook for one building, with the automated quality gates.

    python3 make_budget.py "WEST MOUNT PLEASANT" [--year 2027] [--src DIR] [--out-dir DIR]
                           [--billing-sf N] [--actuals-thru N] [--image FILE] [--version X.Y]
                           [--suffix TEST] [--keep-work] [--skip-reopen]

Steps: build_v2.py (base .xlsx) -> final_build.py (LibreOffice: embed VBA, run Build Schedules, add button,
export DETAIL font colours) -> postprocess.py (tab colours, button macro, font colours, collapsed outline flags)
-> gates -> E<id>_<BUILDING_NAME>_<year>_Budget[_<suffix>]_v<X.Y>.xlsm in --out-dir.

Automated gates (PROJECT_INSTRUCTIONS "Quality gates"):
  G1  zero formula errors in the saved values of every sheet
  G2  DETAIL status (B3) reported; "OK" expected, "CHECK DATA" reported as a warning (data issue, not a build failure)
  G3  embedded VBA identical to VBA_modules/*.bas, no VBASupport line, vba_lint.py clean on both
  G4  worksheet/workbook/styles XML valid against sml.xsd (skipped with a warning if the schema is not installed);
      exactly one <tabColor> per sheet
  G5  reopen the finished file, import its VBA and rerun the build (final_reopen.py)
  G6  tab order and colours, every sheet protected, Build Schedules button assigned
Not automated: the Allwood-specific test_*.py scripts (run them separately for template changes), rendering/visual
review, and anything that needs Excel. Exit code 0 = all gates passed (warnings allowed), 1 = a gate failed.
"""
import argparse, glob, json, os, re, shutil, subprocess, sys, tempfile, zipfile

from paths import HERE, REPO, VBA_DIR

ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
ap.add_argument("building", help="MRI building ID (e.g. 05) or building name / unique part of it")
ap.add_argument("--year", type=int, default=2027, help="budget year (default 2027)")
ap.add_argument("--src", default=None, help="MRI extract folder (default: $BUDGET_SRC, else <repo>/MRI_Files_<year>)")
ap.add_argument("--out-dir", default=os.path.join(REPO, "output"), help="where the finished workbook goes (default <repo>/output, not committed)")
ap.add_argument("--billing-sf", type=float, default=None, help="billing SF for INPUT!C5 (default: MRI leasable SF)")
ap.add_argument("--actuals-thru", type=int, default=None, help="actuals-through month for INPUT!C15 (default: last month with activity)")
ap.add_argument("--image", default=None, help="cover photo (default: building_<id>_image.jpg/.png in the source folder)")
ap.add_argument("--version", default=None, help="template version for the file name (default: from the workbook in <repo>/workbook)")
ap.add_argument("--suffix", default="", help="optional tag in the file name, e.g. TEST")
ap.add_argument("--keep-work", action="store_true", help="keep the intermediate files (base .xlsx, LibreOffice .xlsm, logs)")
ap.add_argument("--skip-reopen", action="store_true", help="skip gate G5 (saves about a minute)")
A = ap.parse_args()

SCHEMA = "/mnt/skills/public/xlsx/scripts/office/schemas/ISO-IEC29500-4_2016/sml.xsd"
NSM = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
ORANGE, PURPLE, BLUE, NAVY = "FFFFA500", "FF7030A0", "FF0070C0", "FF1F497D"
EXPECTED_TABS = (["COVER", "INPUT", "REVENUE", "OTHER_REVENUE", "EXP_UTILITIES", "EXP_JANITORIAL", "EXP_ELEVATOR", "EXP_HVAC",
                  "EXP_LANDSCAPING", "EXP_GENERAL_RM", "TAXES", "INSURANCE", "PROF_FEES", "OTHER_PROF_FEES", "DEBT",
                  "REPAIRS_CAPITAL", "DETAIL", "PROFORMA", "QTRLY SUM", "TENANT_SCHEDULE", "INCOME_ACCOUNTS", "EXPENSE_ACCOUNTS",
                  "DATA_COA", "DATA_CHARGECODES", "DATA_BUDGET_PY", "DATA_ACTUALS", "DATA_BALANCES", "DATA_SUITES",
                  "DATA_TENANT_CHARGES", "DATA_VENDOR"])
TAB_COLOUR = {n: (NAVY if i == 0 else ORANGE if i <= 15 else PURPLE if i <= 18 else BLUE) for i, n in enumerate(EXPECTED_TABS)}

results = []          # (gate, "PASS"/"WARN"/"FAIL"/"SKIP", detail)
def gate(name, status, detail=""):
    results.append((name, status, detail))
    print(f"  [{status}] {name}" + (f" - {detail}" if detail else ""), flush=True)


def run(step, args, log, timeout):
    """Run one pipeline script from build_tools/; stop the whole build if it fails."""
    print(f"-> {step}", flush=True)
    with open(log, "w") as fh:
        try:
            rc = subprocess.run([sys.executable] + args, cwd=HERE, stdout=fh, stderr=subprocess.STDOUT, timeout=timeout).returncode
        except subprocess.TimeoutExpired:
            rc = "timeout"
    out = open(log).read()
    if rc != 0:
        print(out[-3000:])
        sys.exit(f"STOPPED: {step} failed ({rc}). Log: {log}")
    return out


def lo_reset(*dirs):
    """Stale soffice processes or lock files make LibreOffice load() return None (continuation §10)."""
    subprocess.run(["pkill", "-x", "soffice.bin"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    for d in dirs:
        for f in glob.glob(os.path.join(d, ".~lock.*")):
            os.remove(f)


def default_version():
    vs = [re.search(r"_v(\d+\.\d+)\.xlsm$", f) for f in os.listdir(os.path.join(REPO, "workbook"))]
    vs = sorted((tuple(map(int, m.group(1).split("."))) for m in vs if m), reverse=True)
    return "%d.%d" % vs[0] if vs else "0.0"


def sheet_parts(z):
    wbx = z.read("xl/workbook.xml").decode()
    rels = z.read("xl/_rels/workbook.xml.rels").decode()
    rid2t = dict(re.findall(r'<Relationship Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    rid2t.update({a: b for b, a in re.findall(r'Target="([^"]+)"[^>]*Id="([^"]+)"', rels)})
    return [(n.replace("&amp;", "&"), "xl/" + rid2t[r].lstrip("/").replace("xl/", "", 1))
            for n, r in re.findall(r'<sheet name="([^"]+)"[^>]*r:id="([^"]+)"', wbx)]


def norm_vba(text):
    return [l.rstrip() for l in text.replace("\r", "").split("\n") if l.strip() and not l.startswith("Attribute ")]


# ------------------------------------------------------------------ build
src = A.src or os.environ.get("BUDGET_SRC") or os.path.join(REPO, f"MRI_Files_{A.year}")
if not os.path.isdir(src):
    sys.exit(f"MRI source folder not found: {src}")
os.makedirs(A.out_dir, exist_ok=True)
work = tempfile.mkdtemp(prefix="make_budget_", dir=A.out_dir)
base, lo_file, pp_file = (os.path.join(work, f) for f in ("base.xlsx", "lo.xlsm", "post.xlsm"))
colors = os.path.join(work, "colors_detail.json")

cmd = ["build_v2.py", base, "--building", A.building, "--year", str(A.year), "--src", src]
for flag, val in (("--billing-sf", A.billing_sf), ("--actuals-thru", A.actuals_thru), ("--image", A.image)):
    if val is not None:
        cmd += [flag, str(val)]
out = run("build base workbook (build_v2.py)", cmd, os.path.join(work, "1_build.log"), 400)
m = re.search(r"^building: (\S+) (.+?) \| budget year", out, re.M)
bldg_id, bldg_name = m.group(1), m.group(2).strip()
print(f"   building {bldg_id} {bldg_name}, budget year {A.year}")

lo_reset(work, HERE)
out = run("LibreOffice: embed VBA, Build Schedules, button (final_build.py)",
          ["final_build.py", base, lo_file, colors], os.path.join(work, "2_final_build.log"), 500)
lo_reset(work, HERE)
run("post-process (postprocess.py)", ["postprocess.py", lo_file, pp_file, colors], os.path.join(work, "3_postprocess.log"), 300)

version = A.version or default_version()
safe_name = re.sub(r"[^A-Z0-9]+", "_", bldg_name.upper()).strip("_")
final_name = f"E{bldg_id}_{safe_name}_{A.year}_Budget{'_' + A.suffix if A.suffix else ''}_v{version}.xlsm"
final = os.path.join(A.out_dir, final_name)
shutil.copyfile(pp_file, final)

# ------------------------------------------------------------------ gates
print("-> quality gates")
import openpyxl
wb = openpyxl.load_workbook(final, data_only=True, keep_vba=True)

# G1 formula errors (cached values written by LibreOffice after the build)
errs = []
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            v = c.value
            if isinstance(v, str) and (re.match(r"^#(REF!|VALUE!|DIV/0!|N/A|NAME\?|NUM!|NULL!)", v) or v.startswith("Err:")):
                errs.append(f"{ws.title}!{c.coordinate}={v}")
gate("G1 zero formula errors", "PASS" if not errs else "FAIL", f"{len(errs)} error cells: {errs[:8]}" if errs else "")

# G2 DETAIL status
status = str(wb["DETAIL"]["B3"].value or "")
gate("G2 DETAIL status", "PASS" if status.startswith("OK") else "WARN" if status.startswith("CHECK DATA") else "FAIL", status)

# G6 tab order, colours, protection, button
names = wb.sheetnames
gate("G6 tab order", "PASS" if names == EXPECTED_TABS else "FAIL",
     "" if names == EXPECTED_TABS else f"got {names}")
z = zipfile.ZipFile(final)
parts = sheet_parts(z)
bad_col, unprot, tabcount = [], [], []
for n, p in parts:
    x = z.read(p).decode("utf-8")
    k = x.count("<tabColor")
    if k != 1:
        tabcount.append(f"{n}:{k}")
    mcol = re.search(r'<tabColor rgb="([0-9A-F]{8})"', x)
    if n in TAB_COLOUR and (not mcol or mcol.group(1) != TAB_COLOUR[n]):
        bad_col.append(n)
    if not re.search(r'<sheetProtection[^>]*sheet="(true|1)"', x):
        unprot.append(n)
gate("G6 tab colours", "PASS" if not bad_col else "FAIL", f"wrong colour: {bad_col}" if bad_col else "")
gate("G6 every sheet protected", "PASS" if not unprot else "FAIL", f"unprotected: {unprot}" if unprot else "")
btn = [n for n in z.namelist() if re.match(r"xl/drawings/drawing\d+\.xml$", n) and b"BuildSchedulesButton" in z.read(n)]
assigned = bool(btn) and b'macro="[0]!BuildAllSchedules"' in z.read(btn[0])
gate("G6 Build Schedules button assigned", "PASS" if assigned else "FAIL", btn[0] if btn else "button drawing not found")

# G4 schema + one tabColor per sheet
gate("G4 exactly one <tabColor> per sheet", "PASS" if not tabcount else "FAIL", f"{tabcount}" if tabcount else "")
if os.path.exists(SCHEMA):
    from lxml import etree
    schema = etree.XMLSchema(etree.parse(SCHEMA))
    invalid = []
    for p in [p for _, p in parts] + ["xl/workbook.xml", "xl/styles.xml"]:
        if not schema.validate(etree.fromstring(z.read(p))):
            invalid.append(f"{p}: {schema.error_log.last_error}")
    gate("G4 XML valid against sml.xsd (worksheets, workbook, styles)", "PASS" if not invalid else "FAIL",
         "; ".join(invalid[:3]))
else:
    gate("G4 XML schema validation", "SKIP", f"schema not installed ({SCHEMA})")

# G3 VBA identical to source, no VBASupport, lint clean on both
try:
    from oletools.olevba import VBA_Parser
    emb_dir = os.path.join(work, "vba_embedded"); os.makedirs(emb_dir, exist_ok=True)
    found, diffs, vbasup = set(), [], []
    for (_, _, name, code) in VBA_Parser(final).extract_macros():
        name = name.replace(".bas", "")
        srcf = os.path.join(VBA_DIR, f"{name}.bas")
        if not os.path.exists(srcf):
            continue
        found.add(name)
        open(os.path.join(emb_dir, f"{name}.bas"), "w").write(code)
        if "VBASupport" in code:
            vbasup.append(name)
        if norm_vba(code) != norm_vba(open(srcf).read()):
            diffs.append(name)
    expected = {os.path.basename(f)[:-4] for f in glob.glob(os.path.join(VBA_DIR, "*.bas"))}
    missing = sorted(expected - found)
    ok = not (diffs or vbasup or missing)
    gate("G3 embedded VBA identical to VBA_modules", "PASS" if ok else "FAIL",
         f"differ: {diffs} missing: {missing} VBASupport: {vbasup}" if not ok else f"{len(found)} modules")
    lint = []
    for d in (VBA_DIR, emb_dir):
        r = subprocess.run([sys.executable, "vba_lint.py"] + sorted(glob.glob(os.path.join(d, "*.bas"))),
                           cwd=HERE, capture_output=True, text=True)
        if "VBA lint: clean" not in r.stdout:
            lint.append(f"{d}: {r.stdout.strip()[-300:]}")
    gate("G3 vba_lint.py clean (source and embedded)", "PASS" if not lint else "FAIL", "; ".join(lint))
except ImportError:
    gate("G3 embedded VBA check", "SKIP", "oletools not installed (pip install oletools)")

# G5 reopen + rebuild from the embedded VBA
if A.skip_reopen:
    gate("G5 reopen and rebuild", "SKIP", "--skip-reopen")
else:
    lo_reset(A.out_dir, work, HERE)
    reopen_copy = os.path.join(work, "reopen.xlsm")
    shutil.copyfile(final, reopen_copy)       # never let the gate touch the delivered file
    print("-> reopen and rebuild (final_reopen.py)", flush=True)
    try:
        r = subprocess.run([sys.executable, "final_reopen.py", reopen_copy], cwd=HERE, capture_output=True, text=True, timeout=400)
        line = next((l for l in r.stdout.splitlines() if l.startswith("rebuild from imported VBA")), "")
        st = line.split("->", 1)[1].split("|")[0].strip() if line else ""
        good = r.returncode == 0 and st and st == status
        gate("G5 reopen and rebuild", "PASS" if good else "FAIL",
             (line or (r.stdout + r.stderr)[-500:]) + ("" if good else f" (expected status: {status})"))
    except subprocess.TimeoutExpired:
        gate("G5 reopen and rebuild", "FAIL", "timeout")
    lo_reset(A.out_dir, work, HERE)

# ------------------------------------------------------------------ summary
failed = [r for r in results if r[1] == "FAIL"]
if A.keep_work:
    print("work files kept in", work)
else:
    shutil.rmtree(work, ignore_errors=True)
print()
print("=" * 78)
print(f"{'FAILED' if failed else 'DONE'}: {final}")
print(f"building {bldg_id} {bldg_name} | year {A.year} | version v{version} | status: {status}")
print(f"gates: {sum(r[1] == 'PASS' for r in results)} pass, {sum(r[1] == 'WARN' for r in results)} warn, "
      f"{len(failed)} fail, {sum(r[1] == 'SKIP' for r in results)} skipped")
print("Not tested here: Excel itself. Open the file in Excel, click Build Schedules and report any message verbatim.")
sys.exit(1 if failed else 0)
