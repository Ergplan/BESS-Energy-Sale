"""Step 1 · Read the IEX 15-minute xlsx files in the project folder and write prices.json.
Pure Python (no openpyxl needed): python3 source/extract_prices.py"""
import zipfile, re, json, os, collections
HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(HERE)
FILES = {'DAM': ('DAM_15min_2025-09-01_to_2026-09-23.xlsx', 'H'),
         'GDAM': ('GDAM_15min_2025-09-01_to_2026-09-23.xlsx', 'L'),
         'RTM': ('RTM_15min_2025-09-01_to_2026-09-23.xlsx', 'I')}   # column holding MCP (₹/MWh)
rowre = re.compile(r'<row r="(\d+)"[^>]*>(.*?)</row>', re.S)
cellre = re.compile(r'<c r="([A-Z]+)\d+"[^>]*>(?:<v>(.*?)</v>)?</c>', re.S)
data = {}
for m, (f, col) in FILES.items():
    xml = zipfile.ZipFile(os.path.join(ROOT, f)).read('xl/worksheets/sheet1.xml').decode()
    d = collections.defaultdict(dict)
    for rn, body in rowre.findall(xml):
        if rn == '1': continue
        c = dict(cellre.findall(body)); tb = c.get('B', '').replace(' ', '')
        try: h, mi = tb.split('-')[0].split(':'); blk = int(h)*4 + int(mi)//15
        except Exception: continue
        if c.get(col) is not None: d[c.get('A')][blk] = float(c[col])
    data[m] = d
common = sorted(set(data['DAM']) & set(data['GDAM']) & set(data['RTM']))
out = {'dates': [], 'DAM': [], 'GDAM': [], 'RTM': []}
for day in common:
    rows = {}
    for m in ('DAM', 'GDAM', 'RTM'):
        dd = data[m][day]   # fill any missing block with the nearest available block
        rows[m] = [round(dd[b]) if b in dd else round(dd[min(dd, key=lambda k: abs(k-b))]) for b in range(96)]
    out['dates'].append(day)
    for m in rows: out[m].append(rows[m])
json.dump(out, open(os.path.join(HERE, 'prices.json'), 'w'), separators=(',', ':'))
print(f"prices.json: {len(out['dates'])} days, {out['dates'][0]} to {out['dates'][-1]}")
