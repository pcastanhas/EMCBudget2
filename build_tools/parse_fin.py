"""Parse the MRI Financial Format preview (EMCSTD_V2) into an account -> group mapping.

Report structure (per the PDF):
  TI = title (starts a group), RA/SM = account ranges (+ included, - excluded),
  SU/TO = subtotal / total that rolls up the lines above it, BL = blank.
Net membership of a line = included accounts minus excluded accounts.
"""
import re, json, sys, os
from collections import defaultdict, OrderedDict

SRC = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), "fin_layout.txt")
OUT = sys.argv[2] if len(sys.argv) > 2 else os.path.join(os.path.dirname(os.path.abspath(__file__)), "fin_groups.json")

HDR = re.compile(r"^\s{0,8}(\d{1,3})\s{2,}(BL|TI|RA|SU|TO|SM)\b(.*)$")
ACC = re.compile(r"^\s+([+-])\s+(\d{4})-(\d{3})\s+(.*?)\s*$")
SKIP = ("Database:", "Report ID:", "Report Fmt:", "Format Type", "Sort", "Seq.", "Title", "Type")


def clean_opt(s):
    """drop a trailing single-letter option code ('INCOME S' -> 'INCOME'; 'P' -> '')"""
    s = s.strip()
    if re.fullmatch(r"[A-Z]", s):
        return ""
    return re.sub(r"\s+[A-Z]$", "", s)


lines = []          # ordered dicts: seq, type, label, plus, minus
cur = None
for raw in open(SRC, encoding="utf-8", errors="replace"):
    raw = raw.rstrip("\n")
    m = HDR.match(raw)
    if m and not raw.lstrip().startswith(("+", "-")):
        seq, typ, rest = int(m.group(1)), m.group(2), m.group(3).strip()
        label = ""
        if typ == "TI":
            label = clean_opt(re.sub(r"^Title\s*", "", rest))
        elif typ == "SU":
            mm = re.match(r"Sub-total\s+\d+\s+(?:(?:Debit|Credit)\s+)?(.*)$", rest)
            label = clean_opt(mm.group(1)) if mm else ""
        elif typ == "TO":
            mm = re.match(r"Total\s+(?:Debit|Credit)\s+(.*?)(?:\s+Sub-totals.*|\s+=.*)?$", rest)
            label = clean_opt(mm.group(1)) if mm else ""
        elif typ == "SM":
            mm = re.match(r"Summary\s+(?:(?:Debit|Credit)\s+)?(.*?)(?:\s{2,}[A-Z\^]?\s*Include.*|\s+Include.*)?$", rest)
            label = clean_opt(mm.group(1)) if mm else ""
        cur = OrderedDict(seq=seq, type=typ, label=label, plus=[], minus=[])
        lines.append(cur)
        continue
    a = ACC.match(raw)
    if a and cur is not None:
        acct = int(a.group(2) + a.group(3))
        (cur["plus"] if a.group(1) == "+" else cur["minus"]).append(acct)

# fix extraction glue seen in the PDF ("TOTALTAXES")
for L in lines:
    if L["label"] == "TOTALTAXES":
        L["label"] = "TOTAL TAXES"

by_seq = {L["seq"]: L for L in lines}
seqs = sorted(by_seq)
MAIN_END = 121            # main body: income statement + cash-flow lines (through the capital/other lines)
SUBGROUPS = {"UTILITIES", "JANITORIAL & CLEANING", "ELEVATOR", "HVAC & PAYROLL", "LANDSCAPING & GROUNDS", "GENERAL R&M"}

# governing title for each line (main body)
title = ""
for L in lines:
    if L["seq"] > MAIN_END:
        break
    if L["type"] == "TI" and L["label"]:
        title = L["label"]
    L["title"] = title


def next_subtotal(seq):
    for s in seqs:
        if s > seq and by_seq[s]["type"] in ("SU", "TO") and by_seq[s]["label"]:
            return by_seq[s]["label"]
    return ""


members = defaultdict(list)     # acct -> [(seq, source)]
for L in lines:
    if L["type"] not in ("RA", "SM"):
        continue
    net = set(L["plus"]) - set(L["minus"])
    L["net"] = sorted(net)
    for acct in net:
        members[acct].append(L["seq"])

SUMMARY_MAP = {          # supplemental block (lines 137-149): group of each summary line
    137: "UTILITIES", 138: "UTILITIES", 139: "UTILITIES", 142: "GENERAL R&M",
    143: "PROFESSIONAL FEES/G&A", 144: "PROFESSIONAL FEES/G&A", 145: "PROFESSIONAL FEES/G&A",
    148: "INSURANCE", 149: "TAXES"}

result = {}
conflicts = {}
for acct, seq_list in members.items():
    main = [s for s in seq_list if s <= MAIN_END]
    cash = [s for s in seq_list if 122 <= s <= 132]
    summ = [s for s in seq_list if s in SUMMARY_MAP]
    if main:
        s = main[0]
        L = by_seq[s]
        g = L["title"]
        note = ""
        if len(main) > 1:
            conflicts[acct] = main
            others = [f"line {o} ({by_seq[o]['title']})" for o in main[1:]]
            note = "Also included by " + ", ".join(others) + "; first line used"
        result[acct] = dict(group=g, parent="EXPENSES" if g in SUBGROUPS else "", subtotal=next_subtotal(s), line=s, source="main", note=note)
    elif cash:
        s = cash[0]
        result[acct] = dict(group="CASH BALANCES", parent="", subtotal="", line=s, source="cash", note=f"Cash block only (line {s})")
    elif summ:
        s = summ[0]
        result[acct] = dict(group=SUMMARY_MAP[s], parent="EXPENSES" if SUMMARY_MAP[s] in SUBGROUPS else "", subtotal="", line=s, source="summary", note=f"Summary block only (line {s}: {by_seq[s]['label'] or 'summary'})")

json.dump({"accounts": {str(k): v for k, v in result.items()},
           "conflicts": {str(k): v for k, v in conflicts.items()},
           "lines": [{k: L[k] for k in ("seq", "type", "label")} | {"title": L.get("title", ""), "n": len(L.get("net", []))} for L in lines]},
          open(OUT, "w"), indent=1)

if __name__ == "__main__":
    print(f"parsed {len(lines)} format lines | {len(result)} accounts mapped | {len(conflicts)} accounts included by 2+ main lines")
    print("\nFormat lines (seq type label | net members):")
    for L in lines:
        if L["type"] in ("TI", "SU", "TO") and L["label"]:
            print(f"  {L['seq']:>3} {L['type']} {L['label']}")
    print("\nGroup sizes (main body):")
    from collections import Counter
    c = Counter((v["group"]) for v in result.values())
    for g, n in c.most_common():
        print(f"  {n:5}  {g}")
