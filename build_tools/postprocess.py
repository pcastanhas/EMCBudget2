import zipfile, re, sys, json, copy
from lxml import etree
from xml.dom import minidom
src, dst = sys.argv[1], sys.argv[2]

NSM = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
def q(t): return "{%s}%s" % (NSM, t)

def mark_collapsed(sheet_xml):
    """LibreOffice writes collapsed="false" on the visible summary row under a collapsed row group. Excel then shows the group button as expanded although the rows are hidden:
    set collapsed="true" on the row that follows each run of hidden outline rows."""
    hid = {int(m.group(1)) for m in re.finditer(r'<row r="(\d+)"[^>]*hidden="true"[^>]*outlineLevel="[1-7]"', sheet_xml)}
    ends = {r + 1 for r in hid if (r + 1) not in hid}
    def fix(m):
        return m.group(0).replace('collapsed="false"', 'collapsed="true"') if int(m.group(1)) in ends else m.group(0)
    return re.sub(r'<row r="(\d+)"[^>]*>', fix, sheet_xml)


def patch_colors(styles_b, sheet_b, colors):
    """Give each listed cell an explicit font colour (clone font + cellXf, never edit shared ones)."""
    st = etree.fromstring(styles_b); sh = etree.fromstring(sheet_b)
    fonts = st.find(q("fonts")); xfs = st.find(q("cellXfs")); fc, xc = {}, {}
    def new_font(fid, rgb):
        k = (fid, rgb)
        if k not in fc:
            f = copy.deepcopy(fonts[fid]); col = f.find(q("color"))
            if col is None:
                col = etree.Element(q("color")); sz = f.find(q("sz"))
                if sz is not None: sz.addnext(col)
                else: f.insert(0, col)
            for a in list(col.attrib): del col.attrib[a]
            col.set("rgb", "FF" + rgb)
            fonts.append(f); fonts.set("count", str(len(fonts))); fc[k] = len(fonts) - 1
        return fc[k]
    def new_xf(xid, rgb):
        k = (xid, rgb)
        if k not in xc:
            x = copy.deepcopy(xfs[xid]); x.set("fontId", str(new_font(int(x.get("fontId")), rgb))); x.set("applyFont", "true")
            xfs.append(x); xfs.set("count", str(len(xfs))); xc[k] = len(xfs) - 1
        return xc[k]
    n = 0
    for c in sh.iter(q("c")):
        rgb = colors.get(c.get("r"))
        if rgb:
            c.set("s", str(new_xf(int(c.get("s", "0")), rgb))); n += 1
    ser = lambda t: etree.tostring(t, xml_declaration=True, encoding="UTF-8", standalone=True)
    return ser(st), ser(sh), n

zin = zipfile.ZipFile(src)
# map sheet name -> part
wbx = zin.read("xl/workbook.xml").decode()
rels = zin.read("xl/_rels/workbook.xml.rels").decode()
rid2t = dict(re.findall(r'<Relationship Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
rid2t.update({a: b for b, a in re.findall(r'Target="([^"]+)"[^>]*Id="([^"]+)"', rels)})
sheets = re.findall(r'<sheet name="([^"]+)"[^>]*r:id="([^"]+)"', wbx)
part = {n.replace("&amp;", "&"): "xl/" + rid2t[r] for n, r in sheets}
print("sheet parts:", part)
colors = {"COVER": "FF1F497D", "INPUT": "FFFFA500", "REVENUE": "FFFFA500", "LEASE_DEALS": "FFFFA500", "OTHER_REVENUE": "FFFFA500", "EXP_UTILITIES": "FFFFA500", "EXP_JANITORIAL": "FFFFA500", "EXP_ELEVATOR": "FFFFA500", "EXP_HVAC": "FFFFA500", "EXP_LANDSCAPING": "FFFFA500", "EXP_GENERAL_RM": "FFFFA500", "TAXES": "FFFFA500", "INSURANCE": "FFFFA500", "PROF_FEES": "FFFFA500", "DEBT": "FFFFA500", "OTHER_PROF_FEES": "FFFFA500", "REPAIRS_CAPITAL": "FFFFA500", "INCOME_ACCOUNTS": "FF0070C0", "EXPENSE_ACCOUNTS": "FF0070C0", "DETAIL": "FF7030A0", "PROFORMA": "FF7030A0", "QTRLY SUM": "FF7030A0", "TENANT_SCHEDULE": "FF0070C0", "DEAL_SCHEDULE": "FF0070C0"}
for _n in ("DATA_COA", "DATA_CHARGECODES", "DATA_BUDGET_PY", "DATA_ACTUALS", "DATA_BALANCES", "DATA_SUITES", "DATA_TENANT_CHARGES", "DATA_VENDOR"):
    colors[_n] = "FF0070C0"          # protected sheets lose their tab colour in the LibreOffice export, so set all of them
by_part = {part[n]: c for n, c in colors.items()}
colors_file = sys.argv[3] if len(sys.argv) > 3 else None
patched_styles, patched_sheets = None, {}
if colors_file:
    cmap = json.load(open(colors_file))
    if cmap and not isinstance(next(iter(cmap.values())), dict):
        cmap = {"DETAIL": cmap}                     # old format: DETAIL cells only
    styles_b = zin.read("xl/styles.xml")
    for sh_, cells_ in cmap.items():
        if sh_ in part:
            styles_b, patched_sheets[part[sh_]], ncol = patch_colors(styles_b, zin.read(part[sh_]), cells_)
            print(f"{sh_} font colours written:", ncol)
    patched_styles = styles_b
btn_part = next((n for n in zin.namelist() if re.match(r"xl/drawings/drawing\d+\.xml$", n) and b"BuildSchedulesButton" in zin.read(n)), None)
print("button drawing:", btn_part)
zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
for item in zin.infolist():
    data = zin.read(item.filename)
    if patched_styles is not None and item.filename == "xl/styles.xml":
        data = patched_styles
    if item.filename in patched_sheets:
        data = patched_sheets[item.filename]
    if item.filename in by_part:
        s = data.decode("utf-8")
        assert '<sheetPr filterMode="false">' in s, item.filename
        if "<tabColor" in s:      # LibreOffice kept it on this sheet: set the colour, never add a second element
            s = re.sub(r'<tabColor[^>]*/>', '<tabColor rgb="%s"/>' % by_part[item.filename], s, count=1)
        else:
            s = s.replace('<sheetPr filterMode="false">', '<sheetPr filterMode="false"><tabColor rgb="%s"/>' % by_part[item.filename], 1)
        s = mark_collapsed(s)
        data = s.encode("utf-8")
    if item.filename == btn_part:
        s = data.decode("utf-8")
        s = s.replace("<xdr:sp>", '<xdr:sp macro="[0]!BuildAllSchedules" textlink="">', 1)
        s = s.replace('<xdr:cNvPr id="0" name="BuildSchedulesButton">', '<xdr:cNvPr id="2" name="BuildSchedulesButton">', 1)
        s = s.replace("<xdr:clientData/>", '<xdr:clientData fLocksWithSheet="0"/>', 1)
        s = s.replace('<a:ln w="0">', '<a:ln w="12700">', 1)
        s = s.replace(' spc="-1"', "")
        minidom.parseString(s.encode("utf-8"))   # well-formedness
        data = s.encode("utf-8")
    zout.writestr(item, data)
zout.close()
print("written", dst)
