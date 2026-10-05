"""Step 5 · Store calculated values in BESS_IEX_Financial_Model.xlsx.

openpyxl saves formulas without results, so previewers that don't calculate (macOS Quick Look,
Google Drive / Gmail previews, phone viewers) show empty cells. This recalculates the workbook with
the `formulas` engine and writes each result next to its formula. The formulas stay live: Excel,
Numbers and Google Sheets still recalculate on open.

python3.12 source/add_cached_values.py      (needs: pip install formulas)
"""
import os, re, zipfile, shutil, tempfile, math
from xml.sax.saxutils import escape
import formulas

HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(HERE)
XLSX = os.path.join(ROOT, 'BESS_IEX_Financial_Model.xlsx')
NAME = os.path.basename(XLSX)

sol = formulas.ExcelModel().loads(XLSX).finish().calculate()
vals = {}   # (SHEET, CELL) -> value
for k, v in sol.items():
    m = re.match(r"'\[" + re.escape(NAME) + r"\](.+)'!([A-Z]+[0-9]+)$", k)
    if m and hasattr(v, 'value'):
        vals[(m.group(1), m.group(2))] = v.value[0][0]

with zipfile.ZipFile(XLSX) as z:
    wb = z.read('xl/workbook.xml').decode()
    rels = z.read('xl/_rels/workbook.xml.rels').decode()
    sheets = re.findall(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wb)
    target = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels) + [(b, a) for a, b in re.findall(r'Target="([^"]+)"[^>]*Id="([^"]+)"', rels)])
    files = {('xl/' + target[rid].lstrip('/').removeprefix('xl/')): name for name, rid in sheets}
    contents = {n: z.read(n) for n in z.namelist()}

filled = missing = 0
cell_re = re.compile(r'<c r="([A-Z]+[0-9]+)"((?: [a-z]+="[^"]*")*)><f>(.*?)</f><v></v></c>', re.S)
for path, sheet in files.items():
    xml = contents[path].decode()
    def fill(m):
        global filled, missing
        ref, attrs, f = m.group(1), re.sub(r' t="[^"]*"', '', m.group(2)), m.group(3)
        v = vals.get((sheet.upper(), ref))
        if v is None or type(v).__name__ in ('XlError', 'Empty') or (isinstance(v, float) and math.isnan(v)):
            missing += 1; return m.group(0)
        filled += 1
        if isinstance(v, bool) or type(v).__name__ == 'bool_':
            return f'<c r="{ref}"{attrs} t="b"><f>{f}</f><v>{int(bool(v))}</v></c>'
        if isinstance(v, str):
            return f'<c r="{ref}"{attrs} t="str"><f>{f}</f><v>{escape(v)}</v></c>'
        return f'<c r="{ref}"{attrs}><f>{f}</f><v>{float(v)!r}</v></c>'
    contents[path] = cell_re.sub(fill, xml).encode()

tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx').name
with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as z:
    for n, data in contents.items(): z.writestr(n, data)
shutil.move(tmp, XLSX)
print(f'cached values written: {filled} formula cells filled, {missing} left blank')
