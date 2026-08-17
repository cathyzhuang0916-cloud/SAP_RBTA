"""
publish.py — Generate a static index.html with all data baked in.

Run:  python publish.py
Then: git add index.html && git commit -m "Publish update" && git push

What works in the static version:
  ✅ Full KPI table (Slide 1) with MoM deltas
  ✅ Qualitative Assessment (Slide 2)
  ✅ Drill-down popups — account data pre-baked for every KPI + region
  ✅ Visual charts (bar + trend) — trend data pre-baked for every KPI
  ❌ Upload new data  (needs Python server)
  ❌ Export PPT       (needs Python server)
"""

import json
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

import pandas as pd
import kpi_engine

BASE_DIR   = Path(__file__).parent
EXCEL_PATH = BASE_DIR / "data" / "Impact KPI.xlsx"
TMPL_PATH  = BASE_DIR / "dashboard.html"
OUT_PATH   = BASE_DIR / "index.html"

# ── 1. Compute KPIs ──────────────────────────────────────────────────────────
print("Computing KPIs from Excel...")
raw_data = kpi_engine.compute_kpis(str(EXCEL_PATH))
kpi_json = kpi_engine.kpi_data_to_json(raw_data)
print(f"  Timestamp  : {kpi_json['timestamp']}")
print(f"  Prior month: {kpi_json.get('prior_ts','')}")

# ── 2. Pre-bake drilldown data for every KPI × region ────────────────────────
print("Pre-computing drilldown data...")
df = pd.read_excel(str(EXCEL_PATH), sheet_name="Impact KPI")

ALL_REGIONS = kpi_engine.ALL_REGIONS  # includes APAC
DRILLDOWN_KEYS = list(kpi_engine.KPI_DRILLDOWN_DEFS.keys())

dd_cache = {}   # { "joule|ANZ": {...}, "joule|APAC": {...}, ... }
for key in DRILLDOWN_KEYS:
    for region in ALL_REGIONS:
        cache_key = f"{key}|{region}"
        try:
            result = kpi_engine.drilldown(None, key, region, df_cached=df)
            dd_cache[cache_key] = result
        except Exception as e:
            dd_cache[cache_key] = {"error": str(e)}
    print(f"  {key} — {len(ALL_REGIONS)} regions done")

# ── 3. Pre-bake trend data for every KPI ─────────────────────────────────────
print("Pre-computing trend data...")
trend_cache = {}
all_kpi_keys = list(kpi_engine.Q3_TARGETS.keys())
for key in all_kpi_keys:
    try:
        trend_cache[key] = kpi_engine.trend(None, key, df_cached=df)
    except Exception as e:
        trend_cache[key] = {"error": str(e)}
print(f"  {len(trend_cache)} KPIs done")

# ── 4. Pre-bake snapshot list ─────────────────────────────────────────────────
ts_col    = pd.to_datetime(df["TIME_STAMP"], errors="coerce")
snapshots = sorted([pd.Timestamp(t).strftime("%Y-%m-%d") for t in ts_col.dropna().unique()])

# ── 5. Read dashboard.html template ──────────────────────────────────────────
html = TMPL_PATH.read_text(encoding="utf-8")

# ── 6. Replace Upload + Export buttons with a view-only notice ───────────────
old_buttons = """    <input type="file" id="fileInput" accept=".xlsx" style="display:none">
    <button class="btn-upload" id="uploadBtn" onclick="document.getElementById('fileInput').click()">
      <span>&#128640;</span>
      <span id="uploadLabel">Upload New Data + Refresh</span>
      <span class="spinner" id="spinner"></span>
    </button>
    <button class="btn-export" id="exportBtn" onclick="exportPpt()">
      <span>&#128196;</span>
      <span>Export (PPT)</span>
    </button>"""

new_buttons = """    <span style="font-size:11px;opacity:.75;color:#cde">
      &#128274; View only &nbsp;|&nbsp; Contact data owner to refresh
    </span>"""

html = html.replace(old_buttons, new_buttons)

# ── 7. Inject static data + replace all server fetch() calls ─────────────────
kpi_js      = json.dumps(kpi_json,    ensure_ascii=False)
dd_js       = json.dumps(dd_cache,    ensure_ascii=False)
trend_js    = json.dumps(trend_cache, ensure_ascii=False)
snap_js     = json.dumps(snapshots,   ensure_ascii=False)

static_block = f"""
// ══════════════════════════════════════════════════════════════════════════════
// STATIC DATA — baked in by publish.py. Do not edit manually.
// ══════════════════════════════════════════════════════════════════════════════
const _STATIC_KPI      = {kpi_js};
const _STATIC_DD       = {dd_js};
const _STATIC_TREND    = {trend_js};
const _STATIC_SNAPS    = {snap_js};

// Replace server calls with static lookups
function loadData() {{
  const d = _STATIC_KPI;
  if (!d || !d.slide1) return;
  document.getElementById('timestamp').textContent = d.timestamp || '–';
  document.getElementById('prior_ts').textContent  = d.prior_ts  || '–';
  buildSlide1(d.slide1);
  buildSlide2(d.slide2 || []);
}}

function _staticSnapshots() {{
  const sel = document.getElementById('ddSnapSel');
  sel.innerHTML = '';
  _STATIC_SNAPS.slice().reverse().forEach(s => {{
    const o = document.createElement('option'); o.value = s; o.textContent = s; sel.appendChild(o);
  }});
}}

function fetchDrilldown(kpiKey, region, snap) {{
  // In static mode, snap selector is informational only — always show latest
  const key = kpiKey + '|' + region;
  const d = _STATIC_DD[key];
  document.getElementById('ddBody').innerHTML = '';
  if (!d || d.error) {{
    document.getElementById('ddBody').innerHTML =
      '<tr><td colspan="99" style="text-align:center;padding:18px;color:#888">No drilldown data available.</td></tr>';
    return;
  }}
  _ddData = d;
  document.getElementById('ddSnap').textContent  = d.snapshot || '–';
  document.getElementById('ddTotal').textContent = d.total;
  document.getElementById('ddMet').textContent   = d.met;
  document.getElementById('ddPct').textContent   = d.pct + '%';
  renderDDTable();
}}

function openChart(kpiKey, kpiLabel) {{
  document.getElementById('chartTitle').textContent = '📊 ' + kpiLabel;
  document.getElementById('chartModal').classList.add('open');
  _chartTab = 'bar';
  document.getElementById('tabBar').classList.add('active');
  document.getElementById('tabTrend').classList.remove('active');
  const d = _STATIC_TREND[kpiKey];
  if (!d || d.error) {{ showToast('No chart data for ' + kpiLabel, 'error'); return; }}
  _chartData = d; renderChart();
}}
"""

# Insert static block just before </script> of the first script tag that has loadData
insert_marker = "// ── Load data from server ─────────────────────────────────────────────────"
old_load = """// ── Load data from server ─────────────────────────────────────────────────
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

if old_load in html:
    html = html.replace(old_load, static_block.strip())
else:
    # Fallback: inject before closing </script> of first large script block
    html = html.replace("// ── Init ──────────────────────────────────────────────────────────────────",
                        static_block + "\n// ── Init ──────────────────────────────────────────────────────────────────")

# ── 8. Replace the snapshots fetch with static version ───────────────────────
old_snap_fetch = """fetch('/snapshots').then(r=>r.json()).then(d=>{
  _snapshots = d.snapshots || [];
  const sel = document.getElementById('ddSnapSel');
  _snapshots.slice().reverse().forEach(s => {
    const o = document.createElement('option'); o.value = s; o.textContent = s; sel.appendChild(o);
  });
});"""

new_snap_fetch = """// snapshots pre-loaded from static data
_snapshots = _STATIC_SNAPS;
_staticSnapshots();"""

html = html.replace(old_snap_fetch, new_snap_fetch)

# ── 9. Remove the Upload and Export JS handlers (dead code in static mode) ───
# Remove upload handler
upload_marker = "// ── Upload handler ─────────────────────────────────────────────────────────"
export_marker = "// ── Export PPT ─────────────────────────────────────────────────────────────"
toast_marker  = "// ── Toast ──────────────────────────────────────────────────────────────────"

u_start = html.find(upload_marker)
t_start = html.find(toast_marker)
if u_start != -1 and t_start != -1 and t_start > u_start:
    html = html[:u_start] + html[t_start:]

# Remove fetchDrilldown (replaced above) — but openChart/fetchDrilldown are now static,
# so we only need to remove the original server-calling versions from the second <script> block
old_fetch_dd = """function fetchDrilldown(kpiKey, region, snap) {
  document.getElementById('ddBody').innerHTML =
    '<tr><td colspan="99" style="text-align:center;padding:20px;color:#888">Loading…</td></tr>';
  fetch(`/drilldown?kpi=${kpiKey}&region=${encodeURIComponent(region)}&snapshot=${snap||''}`)
    .then(r=>r.json()).then(d=>{
      if (d.error) { showToast('Error: '+d.error,'error'); return; }
      _ddData = d;
      document.getElementById('ddSnap').textContent  = d.snapshot || '–';
      document.getElementById('ddTotal').textContent = d.total;
      document.getElementById('ddMet').textContent   = d.met;
      document.getElementById('ddPct').textContent   = d.pct + '%';
      renderDDTable();
    }).catch(e=>showToast('Error: '+e.message,'error'));
}"""
html = html.replace(old_fetch_dd, "// fetchDrilldown replaced by static version above")

old_open_chart = """function openChart(kpiKey, kpiLabel) {
  document.getElementById('chartTitle').textContent = '📊 ' + kpiLabel;
  document.getElementById('chartModal').classList.add('open');
  _chartTab = 'bar';
  document.getElementById('tabBar').classList.add('active');
  document.getElementById('tabTrend').classList.remove('active');
  fetch(`/trend?kpi=${kpiKey}`)
    .then(r=>r.json()).then(d=>{
      if (d.error) { showToast('Chart error: '+d.error,'error'); return; }
      _chartData = d; renderChart();
    }).catch(e=>showToast('Chart error: '+e.message,'error'));
}"""
html = html.replace(old_open_chart, "// openChart replaced by static version above")

# ── 10. Write output ──────────────────────────────────────────────────────────
OUT_PATH.write_text(html, encoding="utf-8")
size_kb = OUT_PATH.stat().st_size // 1024
print(f"\nStatic dashboard written to: {OUT_PATH.name} ({size_kb} KB)")
print("\nNext steps:")
print("  1. git add index.html")
print("  2. git commit -m 'Publish dashboard update'")
print("  3. git push")
print("\nGitHub Pages URL: https://cathyzhuang0916-cloud.github.io/SAP_RBTA/")
