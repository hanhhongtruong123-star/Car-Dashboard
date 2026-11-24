"""Build the dashboard into one self-contained HTML file.

    python3 build.py                                  # sample data -> dist/index.html
    python3 build.py --data data/private              # your own Excel files (keep them out of git)
    python3 build.py --out dist/dashboard.html --cdn  # load Chart.js / SheetJS from cdnjs instead of inlining

Every sheet of every .xlsx in --data is embedded as built-in data; uploads in the page itself still work.
Needs: Python 3 + openpyxl, and `npm install` once (Chart.js and SheetJS are read from node_modules)."""
import argparse, datetime, glob, json, os, openpyxl

ROOT = os.path.dirname(os.path.abspath(__file__))
ap = argparse.ArgumentParser()
ap.add_argument('--data', default=os.path.join(ROOT, 'data', 'sample'))
ap.add_argument('--out', default=os.path.join(ROOT, 'dist', 'index.html'))
ap.add_argument('--cdn', action='store_true', help='load libraries from cdnjs (smaller file, needs internet when opened)')
a = ap.parse_args()

EPOCH = datetime.datetime(1899, 12, 30)
def cell(v):  # Excel values -> the same numbers SheetJS gives in the browser
    if isinstance(v, datetime.datetime): return (v - EPOCH).total_seconds() / 86400
    if isinstance(v, datetime.time): return (v.hour * 3600 + v.minute * 60 + v.second) / 86400
    if isinstance(v, datetime.timedelta): return v.total_seconds() / 86400
    if isinstance(v, str) and v.startswith('='): return None  # formula columns are recomputed in JS
    return v

files = sorted(glob.glob(os.path.join(a.data, '*.xlsx')))
if not files: raise SystemExit(f'No .xlsx files in {a.data}')
sheets = []
for f in files:
    grab_file = 'grab' in os.path.basename(f).lower()
    for ws in openpyxl.load_workbook(f).worksheets:
        name = ws.title if not grab_file or 'grab' in ws.title.lower() else 'Grab ' + ws.title  # Grab sheets are recognised by name
        sheets.append({'name': name, 'aoa': [[cell(v) for v in r] for r in ws.iter_rows(values_only=True) if any(x is not None for x in r)]})
label = ', '.join(os.path.splitext(os.path.basename(f))[0] for f in files)
seed = json.dumps({'label': label, 'sheets': sheets}, ensure_ascii=False, separators=(',', ':')).replace('<', '\\u003c')

lib = lambda p: open(os.path.join(ROOT, 'node_modules', p), encoding='utf-8').read().replace('<!--', '<\\!--')
src = lambda p: open(os.path.join(ROOT, 'src', p), encoding='utf-8').read()
html = src('template.html')
if a.cdn:
    html = html.replace('<script>/*__CHARTJS__*/</script>', '<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>')
    html = html.replace('<script>/*__XLSX__*/</script>', '<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>')
else:
    html = html.replace('/*__CHARTJS__*/', lib('chart.js/dist/chart.umd.js')).replace('/*__XLSX__*/', lib('xlsx/dist/xlsx.full.min.js'))
html = html.replace('/*__SEED__*/', seed).replace('/*__APP__*/', src('app.js'))
os.makedirs(os.path.dirname(os.path.abspath(a.out)), exist_ok=True)
open(a.out, 'w', encoding='utf-8').write(html)
print(f'written {a.out} ({len(html)//1024} KB) from {len(files)} file(s), {len(sheets)} sheet(s)')
