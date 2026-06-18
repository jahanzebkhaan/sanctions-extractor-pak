"""
Global Sanctions Lists — Web App
Hosted on Railway. Runs extractor daily, serves Excel download.
"""

import os
import threading
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from flask import Flask, send_file, jsonify, render_template_string

app = Flask(__name__)

BASE_DIR    = Path(__file__).parent
OUTPUT_FILE = BASE_DIR / 'sanctions_output.xlsx'
CACHE_DIR   = BASE_DIR / 'sanctions_cache'
CACHE_DIR.mkdir(exist_ok=True)

# Track status
status = {
    'last_run':     None,
    'last_success': None,
    'running':      False,
    'records':      {},
    'error':        None,
}

# ── HTML Dashboard ────────────────────────────────────────────────────────────

DASHBOARD_HTML = """
<!DOCTYPE html>
<html>
<head>
  <title>Global Sanctions Lists</title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
      background: #0B1120; color: #F1F5F9; min-height: 100vh;
      padding: 40px 20px;
    }
    .container { max-width: 860px; margin: 0 auto; }
    h1 { font-size: 2rem; font-weight: 800; color: #F8FAFC; margin-bottom: 6px; }
    .subtitle { color: #64748B; font-size: 0.95rem; margin-bottom: 36px; }
    .card {
      background: #1E2D45; border-radius: 12px; padding: 28px;
      margin-bottom: 20px; border: 1px solid #2D3F5A;
    }
    .card h2 { font-size: 1rem; font-weight: 700; color: #93C5FD;
               text-transform: uppercase; letter-spacing: 1px; margin-bottom: 16px; }
    .sources { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
               gap: 12px; }
    .source {
      background: #0F1C2E; border-radius: 8px; padding: 16px;
      border-left: 3px solid #3B82F6;
    }
    .source.ofac   { border-color: #BF9000; }
    .source.ukhmt  { border-color: #2E75B6; }
    .source.eu     { border-color: #548235; }
    .source.unsc   { border-color: #C55A11; }
    .source.nacta  { border-color: #7B3DB5; }
    .source-name { font-weight: 700; font-size: 0.9rem; color: #E2E8F0; }
    .source-count { font-size: 1.6rem; font-weight: 800; color: #F8FAFC;
                    margin: 4px 0; }
    .source-sub { font-size: 0.75rem; color: #64748B; }
    .total-row { display: flex; justify-content: space-between; align-items: center;
                 padding: 16px 0; border-top: 1px solid #2D3F5A; margin-top: 16px; }
    .total-label { font-weight: 700; color: #94A3B8; }
    .total-num { font-size: 1.8rem; font-weight: 800; color: #4ADE80; }
    .btn {
      display: inline-block; padding: 14px 32px; border-radius: 8px;
      font-weight: 700; font-size: 1rem; text-decoration: none;
      cursor: pointer; border: none; transition: opacity 0.2s;
    }
    .btn:hover { opacity: 0.85; }
    .btn-download { background: #16A34A; color: #fff; }
    .btn-refresh  { background: #1E40AF; color: #fff; margin-left: 12px; }
    .btn-disabled { background: #374151; color: #6B7280; cursor: not-allowed; }
    .status-row { display: flex; gap: 24px; flex-wrap: wrap; margin-top: 8px; }
    .status-item { font-size: 0.85rem; color: #64748B; }
    .status-item span { color: #94A3B8; font-weight: 600; }
    .spinner { display: inline-block; width: 16px; height: 16px;
               border: 2px solid #3B82F6; border-top-color: transparent;
               border-radius: 50%; animation: spin 0.8s linear infinite;
               vertical-align: middle; margin-right: 8px; }
    @keyframes spin { to { transform: rotate(360deg); } }
    .badge {
      display: inline-block; padding: 3px 10px; border-radius: 20px;
      font-size: 0.75rem; font-weight: 700; margin-left: 8px;
    }
    .badge-green { background: #14532D; color: #4ADE80; }
    .badge-yellow { background: #713F12; color: #FCD34D; }
    .badge-red { background: #7F1D1D; color: #FCA5A5; }
    footer { text-align: center; color: #334155; font-size: 0.8rem; margin-top: 40px; }
  </style>
  {% if running %}
  <meta http-equiv="refresh" content="10">
  {% endif %}
</head>
<body>
<div class="container">

  <h1>🌐 Global Sanctions Lists</h1>
  <p class="subtitle">
    OFAC · UK HMT · EU FSF · UNSC · NACTA Pakistan — consolidated into one Excel file
  </p>

  <!-- Status card -->
  <div class="card">
    <h2>Status</h2>
    {% if running %}
      <p><span class="spinner"></span> <strong>Refreshing all lists...</strong>
         Page auto-reloads every 10 seconds.</p>
    {% elif error %}
      <p>❌ Last run failed: <code>{{ error }}</code></p>
    {% elif last_success %}
      <p>✅ Ready to download
         <span class="badge badge-green">LIVE</span>
      </p>
    {% else %}
      <p>⚠️ No data yet — click <strong>Refresh</strong> to download all lists.</p>
    {% endif %}
    <div class="status-row" style="margin-top:12px;">
      <div class="status-item">Last run: <span>{{ last_run or 'Never' }}</span></div>
      <div class="status-item">Last success: <span>{{ last_success or 'Never' }}</span></div>
    </div>
  </div>

  <!-- Records card -->
  {% if total > 0 %}
  <div class="card">
    <h2>Records</h2>
    <div class="sources">
      <div class="source ofac">
        <div class="source-name">OFAC</div>
        <div class="source-count">{{ records.get('OFAC', 0) | format_num }}</div>
        <div class="source-sub">US Treasury SDN</div>
      </div>
      <div class="source ukhmt">
        <div class="source-name">UK HMT</div>
        <div class="source-count">{{ records.get('UK HMT', 0) | format_num }}</div>
        <div class="source-sub">OFSI Consolidated</div>
      </div>
      <div class="source eu">
        <div class="source-name">EU FSF</div>
        <div class="source-count">{{ records.get('EU', 0) | format_num }}</div>
        <div class="source-sub">European Commission</div>
      </div>
      <div class="source unsc">
        <div class="source-name">UNSC</div>
        <div class="source-count">{{ records.get('UNSC', 0) | format_num }}</div>
        <div class="source-sub">UN Security Council</div>
      </div>
      <div class="source nacta">
        <div class="source-name">NACTA</div>
        <div class="source-count">{{ records.get('NACTA', 0) | format_num }}</div>
        <div class="source-sub">Pakistan Schedule IV</div>
      </div>
    </div>
    <div class="total-row">
      <div class="total-label">TOTAL RECORDS</div>
      <div class="total-num">{{ total | format_num }}</div>
    </div>
  </div>
  {% endif %}

  <!-- Actions card -->
  <div class="card">
    <h2>Actions</h2>
    {% if output_exists and not running %}
      <a href="/download" class="btn btn-download">⬇ Download Excel</a>
    {% else %}
      <span class="btn btn-disabled">⬇ Download Excel</span>
    {% endif %}
    {% if not running %}
      <a href="/refresh" class="btn btn-refresh">🔄 Refresh All Lists</a>
    {% endif %}
    <p style="color:#64748B; font-size:0.8rem; margin-top:12px;">
      Lists auto-refresh daily at midnight UTC.
      Manual refresh triggers an immediate update (~5 mins).
    </p>
  </div>

  <footer>
    Global Sanctions Lists Extractor · For compliance research only ·
    Always verify against official sources
  </footer>
</div>
</body>
</html>
"""

# ── Jinja filter ──────────────────────────────────────────────────────────────
@app.template_filter('format_num')
def format_num(v):
    try: return f'{int(v):,}'
    except: return str(v)

# ── Routes ────────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    total = sum(status['records'].values()) if status['records'] else 0
    return render_template_string(
        DASHBOARD_HTML,
        running      = status['running'],
        error        = status['error'],
        last_run     = status['last_run'],
        last_success = status['last_success'],
        records      = status['records'],
        total        = total,
        output_exists= OUTPUT_FILE.exists(),
    )

@app.route('/download')
def download():
    if not OUTPUT_FILE.exists():
        return 'No file yet — visit / and click Refresh first.', 404
    timestamp = datetime.now().strftime('%Y%m%d')
    return send_file(
        str(OUTPUT_FILE),
        as_attachment=True,
        download_name=f'Global_Sanctions_Lists_{timestamp}.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )

@app.route('/refresh')
def refresh():
    if status['running']:
        return 'Already running', 429
    threading.Thread(target=run_extractor, daemon=True).start()
    # Redirect back to dashboard
    from flask import redirect
    return redirect('/')

@app.route('/status')
def get_status():
    return jsonify({
        **status,
        'output_exists': OUTPUT_FILE.exists(),
        'output_size_mb': round(OUTPUT_FILE.stat().st_size/1024/1024, 1)
                          if OUTPUT_FILE.exists() else 0,
    })

# ── Extractor runner ──────────────────────────────────────────────────────────

def run_extractor():
    status['running'] = True
    status['error']   = None
    status['last_run'] = datetime.utcnow().strftime('%d %b %Y %H:%M UTC')

    try:
        # Import and run parsers directly
        import importlib.util, sys as _sys

        spec = importlib.util.spec_from_file_location(
            'sanctions_extractor',
            str(BASE_DIR / 'sanctions_extractor.py')
        )
        mod = importlib.util.module_from_spec(spec)

        # Patch paths so cached files land in our directory
        import sanctions_extractor as _dummy
        _sys.modules.pop('sanctions_extractor', None)

        spec.loader.exec_module(mod)

        # Override paths
        mod.CACHE_DIR   = CACHE_DIR
        mod.OUTPUT_FILE = OUTPUT_FILE
        CACHE_DIR.mkdir(exist_ok=True)

        # Download + parse
        local_files = {}
        for key, cfg in mod.SOURCES.items():
            try:
                if key == 'NACTA':
                    local_files[key] = mod.download_nacta(cfg)
                else:
                    local_files[key] = mod.download(key, cfg)
            except Exception as e:
                print(f'[{key}] download failed: {e}')
                local_files[key] = None

        all_rows, counts = [], {}
        for key in mod.ORDER:
            path = local_files.get(key)
            if not path or not Path(path).exists():
                counts[key] = 0
                continue
            try:
                rows = mod.PARSERS[mod.SOURCES[key]['format']](path)
            except Exception as e:
                print(f'[{key}] parse failed: {e}')
                rows = []
            counts[key] = len(rows)
            all_rows.extend(rows)

        mod.build_excel(all_rows, counts)

        status['records']      = counts
        status['last_success'] = datetime.utcnow().strftime('%d %b %Y %H:%M UTC')

    except Exception as e:
        status['error'] = str(e)
        print(f'Extractor error: {e}')
    finally:
        status['running'] = False

# ── Scheduler (daily refresh at midnight UTC) ─────────────────────────────────

def schedule_daily():
    import time
    while True:
        now = datetime.utcnow()
        # Sleep until next midnight UTC
        seconds_until_midnight = (
            (23 - now.hour) * 3600 +
            (59 - now.minute) * 60 +
            (60 - now.second)
        )
        time.sleep(seconds_until_midnight)
        if not status['running']:
            print('Daily refresh triggered...')
            run_extractor()

# ── Startup ───────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    # Start daily scheduler in background
    threading.Thread(target=schedule_daily, daemon=True).start()

    # Run initial extraction if no output file exists
    if not OUTPUT_FILE.exists():
        print('No output file found — running initial extraction...')
        threading.Thread(target=run_extractor, daemon=True).start()

    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
