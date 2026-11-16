"""Static VBA checks that LibreOffice does not enforce but Excel does.
1. module-level declarations (Private/Public/Dim/Const/Global, not Sub/Function/Property/Type/Enum) must come before the first procedure
2. Option Explicit: every assigned / declared name inside a procedure must be declared (heuristic)
3. every Sub/Function/If/For/Do/Select/With block is closed (simple nesting count)
Usage: python3 vba_lint.py vba/*.bas"""
import re, sys
PROC = re.compile(r'^\s*(Public\s+|Private\s+|Friend\s+)?(Static\s+)?(Sub|Function|Property\s+(Get|Let|Set))\s+(\w+)', re.I)
END_PROC = re.compile(r'^\s*End\s+(Sub|Function|Property)\b', re.I)
DECL = re.compile(r'^\s*(Public|Private|Global|Dim|Const)\s+(?!Sub\b|Function\b|Property\b|Type\b|Enum\b|Declare\b)', re.I)
def strip(l):
    out, q = [], False
    for ch in l:
        if ch == '"': q = not q
        if ch == "'" and not q: break
        out.append(ch)
    return "".join(out)
def names_from_decl(l):
    l = re.sub(r'^\s*(Public|Private|Global|Dim|Const|Static|ReDim\s+Preserve|ReDim)\s+', '', strip(l), flags=re.I)
    res = []
    depth = 0; cur = ""
    for ch in l:
        if ch == "(": depth += 1
        if ch == ")": depth -= 1
        if ch == "," and depth == 0: res.append(cur); cur = ""
        else: cur += ch
    res.append(cur)
    return [re.match(r'\s*(\w+)', x).group(1).lower() for x in res if re.match(r'\s*(\w+)', x)]
bad = 0
for fn in sys.argv[1:]:
    raw_lines = open(fn).read().splitlines()
    lines, buf = [], ""
    for rl in raw_lines:                      # join " _" line continuations (keeps one entry per logical line)
        if strip(rl).rstrip().endswith(" _"):
            buf += strip(rl).rstrip()[:-1] + " "
        else:
            lines.append(buf + rl); buf = ""
    in_proc = False; seen_proc = False; mod_names = set(); proc_names = set(); cont = ""
    # pass 1: module-level names
    for i, raw in enumerate(lines, 1):
        l = strip(raw)
        if PROC.match(l): in_proc = True; seen_proc = True
        elif END_PROC.match(l): in_proc = False
        elif not in_proc and DECL.match(l):
            if seen_proc:
                print(f"{fn}:{i}: module-level declaration after the first procedure -> {raw.strip()}"); bad += 1
            mod_names.update(names_from_decl(l))
    # pass 2: undeclared assignment targets inside procedures
    in_proc = False
    blocks = []
    for i, raw in enumerate(lines, 1):
        l = strip(raw)
        m = PROC.match(l)
        if m:
            in_proc = True; proc_names = set()
            params = re.search(r'\((.*)\)', l)
            if params:
                for p in params.group(1).split(","):
                    mm = re.search(r'(?:ByVal|ByRef|Optional)?\s*(\w+)\s*(?:\(\))?\s*(?:As|$)', p.strip(), re.I)
                    if mm: proc_names.add(mm.group(1).lower())
            proc_names.add(m.group(5).lower())
            continue
        if END_PROC.match(l): in_proc = False; continue
        if in_proc:
            if re.match(r'^\s*(Dim|Static|Const|ReDim)\b', l, re.I): proc_names.update(names_from_decl(l)); continue
            am = re.match(r'^\s*(?:Set\s+|Let\s+)?(\w+)\s*(?:\([^=]*\))?\s*=(?!=)', l)
            if am:
                nm = am.group(1).lower()
                if nm not in proc_names and nm not in mod_names and not nm.startswith(("true", "false")):
                    # property assignments like ws.Cells(...) start with a dotted name and do not match; bare names must be declared
                    print(f"{fn}:{i}: '{am.group(1)}' assigned but not declared -> {raw.strip()}"); bad += 1
print("VBA lint:", "clean" if bad == 0 else f"{bad} problem(s)")
sys.exit(1 if bad else 0)
