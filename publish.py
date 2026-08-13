"""
publish.py — Generate a static index.html with KPI data baked in.

Run:  python publish.py
Then: push index.html to GitHub → visible at GitHub Pages URL.
"""

import json
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

import kpi_engine

BASE_DIR   = Path(__file__).parent
EXCEL_PATH = BASE_DIR / "Impact KPI.xlsx"
TMPL_PATH  = BASE_DIR / "dashboard.html"
OUT_PATH   = BASE_DIR / "index.html"

print("Computing KPIs from Excel...")
data = kpi_engine.compute_kpis(str(EXCEL_PATH))
kpi_json = kpi_engine.kpi_data_to_json(data)
print(f"  Timestamp : {kpi_json['timestamp']}")
print(f"  Prior month: {kpi_json.get('prior_ts','')}")

# Inline data as a JS constant + replace fetch('/data') with static load
data_js = json.dumps(kpi_json, ensure_ascii=False, indent=2)

html = TMPL_PATH.read_text(encoding="utf-8")

# Remove the upload & export buttons (no server in static mode)
html = html.replace(
    '<input type="file" id="fileInput" accept=".xlsx" style="display:none">',
    ''
)
html = html.replace(
    '''    <button class="btn-upload" id="uploadBtn" onclick="document.getElementById('fileInput').click()">
      <span class="icon">&#128640;</span>
      <span id="uploadLabel">Upload New Data + Refresh</span>
      <span class="spinner" id="spinner"></span>
    </button>
    <button class="btn-export" id="exportBtn" onclick="exportPpt()">
      <span class="icon">&#128196;</span>
      <span>Export (PPT)</span>
    </button>''',
    '<span style="font-size:11px;opacity:.7;">View only — contact owner to refresh data</span>'
)

# Replace the loadData() fetch with a static inline version
static_load = f"""
// ── Static data (baked in at publish time) ────────────────────────────────
const _STATIC_DATA = {data_js};

function loadData() {{
  const d = _STATIC_DATA;
  if (!d || !d.slide1) return;
  document.getElementById('timestamp').textContent = d.timestamp || '–';
  document.getElementById('prior_ts').textContent  = d.prior_ts  || '–';
  buildSlide1(d.slide1);
  buildSlide2(d.slide2 || []);
}}
"""

# Replace the server-fetch loadData block
old_load = """// ── Load data from server ──────────────────────────────────────────────────
function loadData() {
  fetch('/data')
    .then(r => r.json())
    .then(d => {
      if (!d || !d.slide1) return;
      document.getElementById('timestamp').textContent = d.timestamp || '–';
      document.getElementById('prior_ts').textContent  = d.prior_ts  || '–';
      buildSlide1(d.slide1);
      buildSlide2(d.slide2 || []);
    })
    .catch(e => showToast('Could not load KPI data: ' + e.message, 'error'));
}"""

html = html.replace(old_load, static_load)

# Remove upload and export JS handlers (not needed in static mode)
# Keep everything from the upload comment block through exportPpt closing brace
upload_start = "// ── Upload handler (refresh dashboard only, no PPT download) ───────────────"
export_end   = "    });\n}"   # last line of exportPpt()

start_idx = html.find(upload_start)
# find the end of exportPpt function
export_fn_start = html.find("// ── Export PPT handler")
# find closing brace after exportPpt
close_idx = html.find("\n}\n", export_fn_start) + 3  # include the closing \n}\n

if start_idx != -1 and close_idx > start_idx:
    html = html[:start_idx] + html[close_idx:]

OUT_PATH.write_text(html, encoding="utf-8")
print(f"\nStatic dashboard written to: {OUT_PATH.name}")
print("\nNext steps:")
print("  1. git add index.html")
print("  2. git commit -m 'Publish dashboard update'")
print("  3. git push  (with your token)")
print("\nThen enable GitHub Pages in your repo settings if not already done.")
