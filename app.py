"""
Global Sanctions Lists — Web App (Render / Railway)
Lightweight version: extraction runs on-demand, not at startup.
"""

import os, threading, json
from datetime import datetime
from pathlib import Path
from flask import Flask, send_file, jsonify, redirect, render_template_string

app = Flask(__name__)

BASE_DIR    = Path(__file__).parent
OUTPUT_FILE = BASE_DIR / 'sanctions_output.xlsx'
CACHE_DIR   = BASE_DIR / 'sanctions_cache'
CACHE_DIR.mkdir(exist_ok=True)

status = {
    'last_run':     None,
    'last_success': None,
    'running':      False,
    'records':      {},
    'error':        None,
    'log':          [],
}

DASHBOARD_HTML = """
<!DOCTYPE html>
<html>
<head>
  <title>Global Sanctions Lists</title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
           background: #0B1120; color: #F1F5F9; min-height: 100vh; padding: 40px 20px; }
    .container { max-width: 860px; margin: 0 auto; }
    h1 { font-size: 2rem; font-weight: 800; color: #F8FAFC; margin-bottom: 6px; }
    .subtitle { color: #64748B; font-size: 0.95rem; margin-bottom: 36px; }
    .card { background: #1E2D45; border-radius: 12px; padding: 28px;
            margin-bottom: 20px; border: 1px solid #2D3F5A; }
    .card h2 { font-size: 1rem; font-weight: 700; color: #93C5FD;
               text-transform: uppercase; letter-spacing: 1px; margin-bottom: 16px; }
    .sources { display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 12px; }
    .source { background: #0F1C2E; border-radius: 8px; padding: 16px; border-left: 3px solid #3B82F6; }
    .source.ofac  { border-color: #BF9000; }
    .source.ukhmt { border-color: #2E75B6; }
    .source.eu    { border-color: #548235; }
    .source.unsc  { border-color: #C55A11; }
    .source.nacta { border-color: #7B3DB5; }
    .source-name  { font-weight: 700; font-size: 0.85rem; color: #E2E8F0; }
    .source-count { font-size: 1.5rem; font-weight: 800; color: #F8FAFC; margin: 4px 0; }
    .source-sub   { font-size: 0.7rem; color: #64748B; }
    .total-row    { display: flex; justify-content: space-between; align-items: center;
                    padding: 16px 0; border-top: 1px solid #2D3F5A; margin-top: 16px; }
    .total-num    { font-size: 1.8rem; font-weight: 800; color: #4ADE80; }
    .btn { display: inline-block; padding: 14px 28px; border-radius: 8px;
           font-weight: 700; font-size: 1rem; text-decoration: none;
           cursor: pointer; border: none; margin-right: 10px; margin-top: 8px; }
    .btn-download { background: #16A34A; color: #fff; }
    .btn-refresh  { background: #1E40AF; color: #fff; }
    .btn-disabled { background: #374151; color: #6B7280; cursor: not-allowed; }
    .log-box { background: #0F1C2E; border-radius: 8px; padding: 16px;
               font-family: monospace; font-size: 0.8rem; color: #94A3B8;
               max-height: 200px; overflow-y: auto; white-space: pre-wrap; }
    .spinner { display: inline-block; width: 14px; height: 14px;
               border: 2px solid #3B82F6; border-top-color: transparent;
               border-radius: 50%; animation: spin 0.8s linear infinite;
               vertical-align: middle; margin-right: 6px; }
    @keyframes spin { to { transform: rotate(360deg); } }
    .info { color: #64748B; font-size: 0.85rem; margin-top: 10px; }
    footer { text-align: center; color: #334155; font-size: 0.8rem; margin-top: 40px; }
  </style>
  {% if running %}<meta http-equiv="refresh" content="8">{% endif %}
</head>
<body>
<div class="container">
  <h1>🌐 Global Sanctions Lists</h1>
  <p class="subtitle">OFAC · UK HMT · EU FSF · UNSC · NACTA Pakistan</p>

  <div style="margin-bottom:32px; padding:24px; background:#1E2D45;
              border-radius:12px; border:1px solid #2D3F5A;
              display:flex; align-items:center; gap:20px;">
    <div>
      <div style="font-size:1.6rem; font-weight:800; color:#F8FAFC;
                  letter-spacing:-0.5px;">Jahanzeb Khan</div>
      <a href="https://www.linkedin.com/in/jahanzeb-khan-537756130/"
         target="_blank"
         style="display:inline-flex; align-items:center; gap:8px; margin-top:8px;
                color:#60A5FA; text-decoration:none; font-weight:600;
                font-size:0.9rem;">
        <svg width="18" height="18" viewBox="0 0 24 24" fill="#60A5FA">
          <path d="M20.447 20.452h-3.554v-5.569c0-1.328-.027-3.037-1.852-3.037
                   -1.853 0-2.136 1.445-2.136 2.939v5.667H9.351V9h3.414v1.561h.046
                   c.477-.9 1.637-1.85 3.37-1.85 3.601 0 4.267 2.37 4.267 5.455v6.286z
                   M5.337 7.433a2.062 2.062 0 0 1-2.063-2.065 2.064 2.064 0 1 1
                   2.063 2.065zm1.782 13.019H3.555V9h3.564v11.452z"/>
        </svg>
        Connect on LinkedIn
      </a>
    </div>
  </div>

  <div class="card">
    <h2>Status</h2>
    {% if running %}
      <p><span class="spinner"></span><strong> Downloading &amp; processing all lists...</strong></p>
      <p class="info">This takes 3–5 minutes. Page refreshes automatically.</p>
    {% elif error %}
      <p>❌ Error: {{ error }}</p>
    {% elif last_success %}
      <p>✅ Ready — last updated {{ last_success }}</p>
    {% else %}
      <p>⚠️ No data yet. Click <strong>Refresh All Lists</strong> to start.</p>
    {% endif %}
    {% if last_run %}
    <p class="info" style="margin-top:8px;">Last run: {{ last_run }}</p>
    {% endif %}
  </div>

  {% if total > 0 %}
  <div class="card">
    <h2>Records</h2>
    <div class="sources">
      <div class="source ofac">
        <div class="source-name">OFAC</div>
        <div class="source-count">{{ "{:,}".format(records.get("OFAC",0)) }}</div>
        <div class="source-sub">US Treasury SDN</div>
      </div>
      <div class="source ukhmt">
        <div class="source-name">UK HMT</div>
        <div class="source-count">{{ "{:,}".format(records.get("UK HMT",0)) }}</div>
        <div class="source-sub">OFSI Consolidated</div>
      </div>
      <div class="source eu">
        <div class="source-name">EU FSF</div>
        <div class="source-count">{{ "{:,}".format(records.get("EU",0)) }}</div>
        <div class="source-sub">European Commission</div>
      </div>
      <div class="source unsc">
        <div class="source-name">UNSC</div>
        <div class="source-count">{{ "{:,}".format(records.get("UNSC",0)) }}</div>
        <div class="source-sub">UN Security Council</div>
      </div>
      <div class="source nacta">
        <div class="source-name">NACTA</div>
        <div class="source-count">{{ "{:,}".format(records.get("NACTA",0)) }}</div>
        <div class="source-sub">Pakistan Schedule IV</div>
      </div>
    </div>
    <div class="total-row">
      <span style="font-weight:700;color:#94A3B8;">TOTAL RECORDS</span>
      <span class="total-num">{{ "{:,}".format(total) }}</span>
    </div>
  </div>
  {% endif %}

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
    <p class="info">Lists auto-refresh daily at midnight UTC.</p>
  </div>

  {% if log %}
  <div class="card">
    <h2>Live Log</h2>
    <div class="log-box">{{ log }}</div>
  </div>
  {% endif %}

  <footer>For compliance research only · Always verify against official sources</footer>
</div>
</body>
</html>
"""

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
        log          = '\n'.join(status['log'][-30:]) if status['log'] else '',
    )

@app.route('/download')
def download():
    if not OUTPUT_FILE.exists():
        return 'No file yet — visit / and click Refresh first.', 404
    ts = datetime.now().strftime('%Y%m%d')
    return send_file(
        str(OUTPUT_FILE),
        as_attachment=True,
        download_name=f'Global_Sanctions_{ts}.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )

@app.route('/refresh')
def refresh():
    if status['running']:
        return redirect('/')
    threading.Thread(target=run_extractor, daemon=True).start()
    return redirect('/')

@app.route('/status')
def get_status():
    return jsonify({k: v for k, v in status.items() if k != 'log'})

def log(msg):
    print(msg)
    status['log'].append(msg)
    if len(status['log']) > 100:
        status['log'] = status['log'][-100:]

def run_extractor():
    import importlib.util, sys
    status['running'] = True
    status['error']   = None
    status['log']     = []
    status['last_run']= datetime.utcnow().strftime('%d %b %Y %H:%M UTC')

    try:
        spec = importlib.util.spec_from_file_location(
            'se', str(BASE_DIR / 'sanctions_extractor.py'))
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)

        mod.CACHE_DIR   = CACHE_DIR
        mod.OUTPUT_FILE = OUTPUT_FILE

        # Download non-NACTA lists first (fast)
        local_files = {}
        for key, cfg in mod.SOURCES.items():
            if key == 'NACTA':
                continue
            try:
                local_files[key] = mod.download(key, cfg)
            except Exception as e:
                log(f'[{key}] download error: {e}')
                local_files[key] = None

        # Parse and build Excel without NACTA first
        all_rows_no_nacta, counts_no_nacta = [], {}
        for key in mod.ORDER:
            if key == 'NACTA':
                continue
            path = local_files.get(key)
            if not path or not Path(path).exists():
                counts_no_nacta[key] = 0; continue
            try:
                rows = mod.PARSERS[mod.SOURCES[key]['format']](path)
            except Exception as e:
                log(f'[{key}] parse error: {e}')
                rows = []
            counts_no_nacta[key] = len(rows)
            all_rows_no_nacta.extend(rows)
            log(f'[{key}] {len(rows):,} records parsed')

        # Save interim Excel (without NACTA)
        counts_no_nacta['NACTA'] = 0
        mod.build_excel(all_rows_no_nacta, counts_no_nacta)
        status['records']      = counts_no_nacta.copy()
        status['last_success'] = datetime.utcnow().strftime('%d %b %Y %H:%M UTC')
        log(f'Interim Excel saved ({sum(counts_no_nacta.values()):,} records — NACTA downloading...)')

        # Now download NACTA (slow — Playwright)
        try:
            nacta_cfg = mod.SOURCES['NACTA']
            local_files['NACTA'] = mod.download_nacta(nacta_cfg)
        except Exception as e:
            log(f'[NACTA] download error: {e}')
            local_files['NACTA'] = None

        # Parse NACTA and rebuild final Excel
        nacta_path = local_files.get('NACTA')
        if nacta_path and Path(nacta_path).exists():
            try:
                nacta_rows = mod.PARSERS['xml_nacta'](nacta_path)
            except Exception as e:
                log(f'[NACTA] parse error: {e}')
                nacta_rows = []
        else:
            nacta_rows = []

        counts_no_nacta['NACTA'] = len(nacta_rows)
        all_rows_final = all_rows_no_nacta + nacta_rows
        mod.build_excel(all_rows_final, counts_no_nacta)
        status['records']      = counts_no_nacta.copy()
        status['last_success'] = datetime.utcnow().strftime('%d %b %Y %H:%M UTC')
        log(f'Done — {sum(counts_no_nacta.values()):,} total records')

    except Exception as e:
        status['error'] = str(e)
        log(f'Fatal error: {e}')
    finally:
        status['running'] = False

def schedule_daily():
    import time
    while True:
        now = datetime.utcnow()
        secs = (23-now.hour)*3600 + (59-now.minute)*60 + (60-now.second)
        time.sleep(secs)
        if not status['running']:
            run_extractor()

if __name__ == '__main__':
    threading.Thread(target=schedule_daily, daemon=True).start()
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port, debug=False)
