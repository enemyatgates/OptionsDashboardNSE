"""
ST+GENERATEWEBSITE~01.py
─────────────────────────
NSE FAO Options Dashboard — Website Generator

Reads  docs/data.json
Writes docs/index.html

Usage:
    python "ST+GENERATEWEBSITE~01.py" --data docs/data.json --out docs/index.html
"""

import argparse
import json
from pathlib import Path
from datetime import datetime, timezone


# ── Helpers ───────────────────────────────────────────────────────────────────
def _meta_from_data(data: dict) -> dict:
    meta    = data.get("meta", {})
    chain   = data.get("chain") or {}
    sent    = data.get("sentiment", {})
    latest  = meta.get("latest_date", "—")
    dates   = meta.get("oi_dates", [])
    excel   = meta.get("excel_filename", "")
    gentime = datetime.now(timezone.utc).strftime("%d %b %Y %H:%M UTC")

    sent_latest = sent.get("data", {}).get(latest, {})
    sent_rows   = []
    sig_col  = {"Bullish": "#10B981", "Cautious": "#F97316",
                 "Bearish": "#EF4444", "Neutral":  "#94A3B8"}
    sig_icon = {"Bullish": "🟢", "Cautious": "🟠",
                 "Bearish": "🔴", "Neutral":  "🟡"}
    for part in ["Client", "DII", "FII", "Pro"]:
        s = sent_latest.get(part, {})
        sig   = s.get("signal", "Neutral")
        color = sig_col.get(sig, "#94A3B8")
        icon  = sig_icon.get(sig, "🟡")
        sent_rows.append({
            "part": part, "sig": sig, "color": color, "icon": icon,
            "dir":  s.get("net_fut_idx_dir", "—"),
            "pcr":  s.get("idx_oi_pcr", "—"),
            "ls":   s.get("ls_ratio", "—"),
        })

    return {
        "latest": latest, "dates": dates, "excel": excel, "gentime": gentime,
        "max_pain":       chain.get("max_pain"),
        "atm":            chain.get("atm"),
        "oi_pcr":         chain.get("oi_pcr"),
        "call_wall":      (chain.get("call_walls") or [{}])[0].get("strike"),
        "put_wall":       (chain.get("put_walls")  or [{}])[0].get("strike"),
        "total_call_oi":  chain.get("total_call_oi"),
        "total_put_oi":   chain.get("total_put_oi"),
        "sent_rows":      sent_rows,
        "n_sessions":     len(dates),
        "n_strikes":      len((data.get("chain") or {}).get("data", [])),
    }


def _fmt(v, fmt="#,"):
    if v is None:
        return "—"
    if fmt == "#,":
        return f"{v:,.0f}"
    if fmt == ".3f":
        return f"{v:.3f}"
    return str(v)


# ── HTML template ──────────────────────────────────────────────────────────────
HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1.0"/>
  <title>FAO Claude — NSE Options Dashboard</title>
  <script src="https://cdn.jsdelivr.net/npm/d3@7/dist/d3.min.js"></script>
  <style>
    /* ── Reset ──────────────────────────────────────── */
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    :root {
      --bg:        #0B0F1A;
      --surface:   #141824;
      --surface2:  #1E2535;
      --border:    #2A3347;
      --border2:   #374151;
      --call:      #3B82F6;
      --call-dim:  rgba(59,130,246,0.12);
      --put:       #F97316;
      --put-dim:   rgba(249,115,22,0.12);
      --green:     #10B981;
      --green-dim: rgba(16,185,129,0.10);
      --amber:     #F59E0B;
      --amber-dim: rgba(245,158,11,0.12);
      --red:       #EF4444;
      --text:      #F1F5F9;
      --muted:     #94A3B8;
      --faint:     #4B5563;
      --mono:      'JetBrains Mono', 'Fira Mono', monospace;
      --sans:      'Inter', system-ui, sans-serif;
      --nav-h:     52px;
      --tab-h:     44px;
    }
    html, body { height: 100%; }
    body {
      background: var(--bg);
      color: var(--text);
      font-family: var(--sans);
      font-size: 14px;
      line-height: 1.5;
      display: flex;
      flex-direction: column;
    }

    /* ── Top nav ────────────────────────────────────── */
    .topnav {
      height: var(--nav-h);
      background: var(--surface);
      border-bottom: 1px solid var(--border);
      display: flex;
      align-items: center;
      padding: 0 24px;
      gap: 20px;
      flex-shrink: 0;
      position: sticky;
      top: 0;
      z-index: 50;
    }
    .nav-logo {
      font-size: 15px;
      font-weight: 700;
      color: var(--text);
      display: flex;
      align-items: center;
      gap: 8px;
      white-space: nowrap;
    }
    .nav-logo-dot {
      width: 8px; height: 8px;
      border-radius: 50%;
      background: var(--call);
      flex-shrink: 0;
    }
    .nav-divider {
      width: 1px;
      height: 24px;
      background: var(--border2);
    }
    .nav-meta {
      font-family: var(--mono);
      font-size: 11px;
      color: var(--muted);
      display: flex;
      align-items: center;
      gap: 16px;
    }
    .nav-meta-item { display: flex; align-items: center; gap: 5px; }
    .nav-meta-label { color: var(--faint); }
    .nav-meta-val   { color: var(--text); }
    .nav-meta-val.green { color: var(--green); }
    .nav-meta-val.amber { color: var(--amber); }
    .nav-meta-val.call  { color: var(--call); }
    .nav-meta-val.put   { color: var(--put); }
    .nav-spacer { flex: 1; }
    .nav-status {
      display: flex;
      align-items: center;
      gap: 6px;
      font-family: var(--mono);
      font-size: 11px;
      color: var(--muted);
    }
    .status-dot {
      width: 6px; height: 6px;
      border-radius: 50%;
      background: var(--faint);
    }
    .status-dot.loading { background: var(--amber); animation: pulse 1s infinite; }
    .status-dot.ok      { background: var(--green); }
    .status-dot.error   { background: var(--red); }
    @keyframes pulse {
      0%,100% { opacity: 1; } 50% { opacity: .3; }
    }
    .nav-dl {
      display: flex;
      align-items: center;
      gap: 6px;
      background: var(--call);
      color: #fff;
      border-radius: 6px;
      padding: 5px 12px;
      font-size: 12px;
      font-weight: 600;
      text-decoration: none;
      white-space: nowrap;
      transition: opacity .2s;
    }
    .nav-dl:hover { opacity: .85; }

    /* ── Tab bar ────────────────────────────────────── */
    .tabbar {
      height: var(--tab-h);
      background: var(--surface);
      border-bottom: 1px solid var(--border);
      display: flex;
      align-items: flex-end;
      padding: 0 24px;
      gap: 2px;
      flex-shrink: 0;
      overflow-x: auto;
      scrollbar-width: none;
    }
    .tabbar::-webkit-scrollbar { display: none; }
    .tab-btn {
      display: flex;
      align-items: center;
      gap: 6px;
      padding: 8px 14px;
      border: none;
      background: transparent;
      color: var(--muted);
      font-family: var(--sans);
      font-size: 13px;
      font-weight: 500;
      cursor: pointer;
      border-bottom: 2px solid transparent;
      white-space: nowrap;
      transition: color .15s, border-color .15s;
      margin-bottom: -1px;
    }
    .tab-btn:hover { color: var(--text); }
    .tab-btn.active {
      color: var(--call);
      border-bottom-color: var(--call);
    }
    .tab-icon { font-size: 14px; }

    /* ── Main content ───────────────────────────────── */
    .main {
      flex: 1;
      overflow-y: auto;
      padding: 28px 24px;
    }
    .tab-panel { display: none; }
    .tab-panel.active { display: block; }

    /* ── Metrics strip ──────────────────────────────── */
    .metrics-strip {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
      gap: 12px;
      margin-bottom: 28px;
    }
    .metric-card {
      background: var(--surface);
      border: 1px solid var(--border);
      border-radius: 10px;
      padding: 16px 18px;
    }
    .metric-val {
      font-size: 26px;
      font-weight: 700;
      font-family: var(--mono);
      line-height: 1.1;
      margin-bottom: 4px;
    }
    .metric-lbl {
      font-size: 11px;
      color: var(--muted);
      text-transform: uppercase;
      letter-spacing: .05em;
    }
    .metric-sub {
      font-size: 11px;
      color: var(--faint);
      margin-top: 3px;
      font-family: var(--mono);
    }
    .c-call  { color: var(--call); }
    .c-put   { color: var(--put); }
    .c-green { color: var(--green); }
    .c-amber { color: var(--amber); }
    .c-muted { color: var(--muted); }

    /* ── Section headers ────────────────────────────── */
    .section-hdr {
      display: flex;
      align-items: center;
      gap: 10px;
      margin: 0 0 16px;
    }
    .section-title {
      font-size: 14px;
      font-weight: 600;
      color: var(--text);
    }
    .section-rule {
      flex: 1;
      height: 1px;
      background: var(--border);
    }
    .section-tag {
      font-family: var(--mono);
      font-size: 10px;
      color: var(--faint);
      letter-spacing: .06em;
      text-transform: uppercase;
    }

    /* ── Sentiment table ────────────────────────────── */
    .sent-table {
      width: 100%;
      border-collapse: collapse;
      background: var(--surface);
      border: 1px solid var(--border);
      border-radius: 10px;
      overflow: hidden;
      margin-bottom: 28px;
    }
    .sent-table th {
      background: var(--surface2);
      padding: 9px 16px;
      text-align: left;
      font-size: 11px;
      text-transform: uppercase;
      letter-spacing: .05em;
      color: var(--muted);
      font-weight: 600;
      border-bottom: 1px solid var(--border);
    }
    .sent-table td {
      padding: 10px 16px;
      border-top: 1px solid var(--border);
      font-size: 13px;
      font-family: var(--mono);
    }
    .sent-table tr:first-child td { border-top: none; }
    .sent-table tr:hover td { background: var(--surface2); }

    /* ── Chain summary cards ────────────────────────── */
    .walls-grid {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 12px;
      margin-bottom: 28px;
    }
    .wall-card {
      background: var(--surface);
      border: 1px solid var(--border);
      border-radius: 10px;
      padding: 16px;
    }
    .wall-card-title {
      font-size: 11px;
      text-transform: uppercase;
      letter-spacing: .06em;
      color: var(--muted);
      margin-bottom: 12px;
      font-weight: 600;
    }
    .wall-row {
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: 5px 0;
      border-bottom: 1px solid var(--border);
      font-family: var(--mono);
      font-size: 12px;
    }
    .wall-row:last-child { border-bottom: none; }
    .wall-rank { color: var(--faint); width: 20px; }
    .wall-strike { color: var(--text); font-weight: 600; }
    .wall-oi { color: var(--muted); }
    .wall-bar-wrap {
      width: 80px;
      height: 4px;
      background: var(--border);
      border-radius: 2px;
      overflow: hidden;
    }
    .wall-bar { height: 100%; border-radius: 2px; }
    /* ── Charts ─────────────────────────────────────── */
    .charts-row {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 12px;
      margin-bottom: 28px;
    }
    .chart-card {
      background: var(--surface);
      border: 1px solid var(--border);
      border-radius: 10px;
      padding: 16px 16px 12px;
    }
    .chart-card.wide {
      grid-column: 1 / -1;
    }
    .chart-title {
      font-size: 12px;
      font-weight: 600;
      color: var(--text);
      margin-bottom: 2px;
    }
    .chart-sub {
      font-size: 11px;
      color: var(--faint);
      margin-bottom: 12px;
      font-family: var(--mono);
    }
    .chart-svg { width: 100%; display: block; height: auto; max-width: 100%; }
    /* D3 axis */
    .axis text {
      fill: var(--muted);
      font-family: var(--mono);
      font-size: 10px;
    }
    .axis line, .axis path { stroke: var(--border2); }
    .axis-zero line { stroke: var(--border2); stroke-dasharray: 3,3; }
    .grid line {
      stroke: var(--border);
      stroke-opacity: .6;
      stroke-dasharray: 2,3;
    }
    .grid path { stroke-width: 0; }
    /* Tooltip */
    .d3-tip {
      position: absolute;
      pointer-events: none;
      background: var(--surface2);
      border: 1px solid var(--border2);
      border-radius: 7px;
      padding: 8px 12px;
      font-family: var(--mono);
      font-size: 11px;
      color: var(--text);
      line-height: 1.7;
      box-shadow: 0 4px 16px rgba(0,0,0,.4);
      max-width: 200px;
      z-index: 99;
    }
    /* Legend */
    .legend { display: flex; gap: 14px; flex-wrap: wrap; margin-bottom: 10px; }
    .legend-item {
      display: flex; align-items: center; gap: 5px;
      font-family: var(--mono); font-size: 11px; color: var(--muted);
    }
    .legend-dot { width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; }
    .legend-line {
      width: 18px; height: 2px; border-radius: 1px; flex-shrink: 0;
    }


    /* ── Placeholder panels ─────────────────────────── */
    .placeholder {
      background: var(--surface);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 60px 40px;
      text-align: center;
    }
    .placeholder-icon { font-size: 40px; margin-bottom: 16px; }
    .placeholder-title {
      font-size: 18px;
      font-weight: 600;
      margin-bottom: 8px;
    }
    .placeholder-sub {
      color: var(--muted);
      font-size: 14px;
      max-width: 400px;
      margin: 0 auto 20px;
      line-height: 1.6;
    }
    .placeholder-tag {
      display: inline-block;
      background: var(--surface2);
      border: 1px solid var(--border);
      border-radius: 20px;
      padding: 4px 14px;
      font-family: var(--mono);
      font-size: 11px;
      color: var(--muted);
    }

    /* ── Data badge (bottom-right) ──────────────────── */
    .data-badge {
      position: fixed;
      bottom: 16px;
      right: 20px;
      background: var(--surface2);
      border: 1px solid var(--border);
      border-radius: 8px;
      padding: 8px 14px;
      font-family: var(--mono);
      font-size: 11px;
      color: var(--muted);
      display: flex;
      align-items: center;
      gap: 10px;
      z-index: 40;
    }
    .data-badge-item { display: flex; gap: 5px; }
    .data-badge-lbl { color: var(--faint); }
    .data-badge-sep { color: var(--faint); }


    /* ── Option Chain table ─────────────────────────── */
    #chain-table th {
      background: var(--surface2);
      padding: 8px 10px;
      text-align: right;
      font-size: 10px;
      text-transform: uppercase;
      letter-spacing: .04em;
      color: var(--muted);
      font-weight: 600;
      border-bottom: 1px solid var(--border);
      white-space: nowrap;
      position: sticky;
      top: 0;
    }
    #chain-table th.strike-col { text-align: center; background: var(--surface); }
    #chain-table th.call-hdr   { color: var(--call); background: rgba(59,130,246,0.06); }
    #chain-table th.put-hdr    { color: var(--put);  background: rgba(249,115,22,0.06); }
    #chain-table td {
      padding: 6px 10px;
      text-align: right;
      border-top: 1px solid rgba(42,51,71,0.5);
      font-size: 11px;
      white-space: nowrap;
    }
    #chain-table td.strike-col {
      text-align: center;
      font-weight: 700;
      font-size: 12px;
      background: var(--surface);
      color: var(--text);
      border-left: 1px solid var(--border);
      border-right: 1px solid var(--border);
    }
    #chain-table tr.atm-row td { background: rgba(245,158,11,0.08) !important; }
    #chain-table tr.atm-row td.strike-col { color: var(--amber); }
    #chain-table tr.maxpain-row td.strike-col { color: var(--green); }
    #chain-table tr:hover td { background: var(--surface2); }
    .delta-pos { color: #10B981; }
    .delta-neg { color: #EF4444; }
    .pcr-high  { color: var(--put);  font-weight: 600; }
    .pcr-low   { color: var(--call); font-weight: 600; }

    /* ── Tab panel fade ────────────────────────────────── */
    .tab-panel { animation: none; }
    .tab-panel.active { animation: fadeIn .18s ease; }
    @keyframes fadeIn { from { opacity: 0; transform: translateY(4px); } to { opacity: 1; transform: translateY(0); } }

    /* ── Chart card hover ───────────────────────────────── */
    .chart-card { transition: border-color .2s; }
    .chart-card:hover { border-color: var(--border2); }

    /* ── Tabs disabled state ────────────────────────────── */
    .tab-btn:disabled { opacity: .35; cursor: not-allowed; pointer-events: none; }

    /* ── Footer ─────────────────────────────────────────── */
    .site-footer {
      border-top: 1px solid var(--border);
      padding: 14px 24px;
      display: flex;
      align-items: center;
      gap: 20px;
      font-family: var(--mono);
      font-size: 11px;
      color: var(--faint);
      flex-shrink: 0;
      background: var(--surface);
    }
    .site-footer a { color: var(--faint); text-decoration: none; }
    .site-footer a:hover { color: var(--muted); }
    .footer-sep { color: var(--border2); }

    /* ── Deep-link badges on section headers ────────────── */
    .tab-link-btn {
      background: var(--surface2);
      border: 1px solid var(--border);
      border-radius: 5px;
      color: var(--call);
      font-family: var(--mono);
      font-size: 10px;
      padding: 2px 8px;
      cursor: pointer;
      transition: background .15s;
      text-decoration: none;
    }
    .tab-link-btn:hover { background: var(--call-dim); }

    /* ── Error banner ───────────────────────────────────── */
    .error-banner {
      background: rgba(239,68,68,0.1);
      border: 1px solid rgba(239,68,68,0.3);
      border-radius: 8px;
      padding: 16px 20px;
      color: #FCA5A5;
      font-family: var(--mono);
      font-size: 13px;
      margin: 0 0 20px;
    }
    /* ── Scrollbar ──────────────────────────────────── */
    ::-webkit-scrollbar { width: 6px; height: 6px; }
    ::-webkit-scrollbar-track { background: transparent; }
    ::-webkit-scrollbar-thumb { background: var(--border2); border-radius: 3px; }
  </style>
</head>
<body>

<!-- ── Top nav ──────────────────────────────────────────────────────── -->
<nav class="topnav">
  <div class="nav-logo">
    <span class="nav-logo-dot"></span>
    FAO Claude
  </div>
  <div class="nav-divider"></div>
  <div class="nav-meta">
    <div class="nav-meta-item">
      <span class="nav-meta-label">Session</span>
      <span class="nav-meta-val">LATEST_DATE</span>
    </div>
    <div class="nav-meta-item">
      <span class="nav-meta-label">Max Pain</span>
      <span class="nav-meta-val amber">MAX_PAIN</span>
    </div>
    <div class="nav-meta-item">
      <span class="nav-meta-label">PCR</span>
      <span class="nav-meta-val green" id="nav-pcr">OI_PCR</span>
    </div>
    <div class="nav-meta-item">
      <span class="nav-meta-label">Call Wall</span>
      <span class="nav-meta-val call">CALL_WALL</span>
    </div>
    <div class="nav-meta-item">
      <span class="nav-meta-label">Put Wall</span>
      <span class="nav-meta-val put">PUT_WALL</span>
    </div>
  </div>
  <div class="nav-spacer"></div>
  <div class="nav-status">
    <span class="status-dot loading" id="status-dot"></span>
    <span id="status-text">Loading data…</span>
  </div>
  <a class="nav-dl" href="../outputs/EXCEL_FILE" download title="Download Excel dashboard">
    ⬇ Excel
  </a>
</nav>

<!-- ── Tab bar ──────────────────────────────────────────────────────── -->
<div class="tabbar" id="tabbar">
  <button class="tab-btn active" data-tab="overview">
    <span class="tab-icon">◎</span> Overview
  </button>
  <button class="tab-btn" data-tab="participant" disabled>
    <span class="tab-icon">⊞</span> Participant OI &amp; TV
  </button>
  <button class="tab-btn" data-tab="pcr" disabled>
    <span class="tab-icon">⇅</span> PCR Analysis
  </button>
  <button class="tab-btn" data-tab="bias" disabled>
    <span class="tab-icon">⊿</span> Net Bias &amp; Efficiency
  </button>
  <button class="tab-btn" data-tab="chain" disabled>
    <span class="tab-icon">≡</span> Option Chain
  </button>
  <button class="tab-btn" data-tab="maxpain" disabled>
    <span class="tab-icon">⊗</span> Max Pain
  </button>
  <button class="tab-btn" data-tab="ivskew" disabled>
    <span class="tab-icon">∿</span> IV Skew
  </button>
  <button class="tab-btn" data-tab="history" disabled>
    <span class="tab-icon">⏱</span> Historical Trends
  </button>
</div>

<!-- ── Main content ─────────────────────────────────────────────────── -->
<main class="main" id="main">

  <!-- Overview -->
  <div class="tab-panel active" id="tab-overview">

    <!-- Key metrics -->
    <div class="metrics-strip">
      <div class="metric-card">
        <div class="metric-val c-amber" id="ov-maxpain">—</div>
        <div class="metric-lbl">Max Pain Strike</div>
        <div class="metric-sub" id="ov-maxpain-sub">loading…</div>
      </div>
      <div class="metric-card">
        <div class="metric-val c-call" id="ov-atm">—</div>
        <div class="metric-lbl">ATM Strike</div>
        <div class="metric-sub" id="ov-atm-sub">loading…</div>
      </div>
      <div class="metric-card">
        <div class="metric-val c-green" id="ov-pcr">—</div>
        <div class="metric-lbl">Aggregate OI PCR</div>
        <div class="metric-sub" id="ov-pcr-sub">Put OI / Call OI</div>
      </div>
      <div class="metric-card">
        <div class="metric-val c-call" id="ov-callwall">—</div>
        <div class="metric-lbl">Call Wall (Resistance)</div>
        <div class="metric-sub" id="ov-callwall-sub">loading…</div>
      </div>
      <div class="metric-card">
        <div class="metric-val c-put" id="ov-putwall">—</div>
        <div class="metric-lbl">Put Wall (Support)</div>
        <div class="metric-sub" id="ov-putwall-sub">loading…</div>
      </div>
      <div class="metric-card">
        <div class="metric-val c-muted" id="ov-sessions">—</div>
        <div class="metric-lbl">Sessions Loaded</div>
        <div class="metric-sub" id="ov-dates-sub">loading…</div>
      </div>
    </div>

    <!-- Sentiment -->
    <div class="section-hdr">
      <div class="section-title">Composite Sentiment</div>
      <div class="section-rule"></div>
      <div class="section-tag" id="sent-date-tag">—</div>
    </div>
    <table class="sent-table" id="sent-table">
      <thead>
        <tr>
          <th>Participant</th>
          <th>Fut Index Direction</th>
          <th>Index OI PCR</th>
          <th>L/S Ratio</th>
          <th>Signal</th>
        </tr>
      </thead>
      <tbody id="sent-tbody">
        <tr><td colspan="5" style="color:var(--muted);text-align:center;padding:20px">
          Loading…
        </td></tr>
      </tbody>
    </table>

    <!-- OI Walls -->
    <div class="section-hdr">
      <div class="section-title">OI Walls</div>
      <div class="section-rule"></div>
      <div class="section-tag">Top 5 strikes</div>
    </div>
    <div class="walls-grid">
      <div class="wall-card">
        <div class="wall-card-title" style="color:var(--call)">⬆ Call Wall — Resistance</div>
        <div id="call-walls-body"></div>
      </div>
      <div class="wall-card">
        <div class="wall-card-title" style="color:var(--put)">⬇ Put Wall — Support</div>
        <div id="put-walls-body"></div>
      </div>
    </div>


    <!-- Chart row 1: Net Futures OI Trend (wide) -->
    <div class="section-hdr">
      <div class="section-title">Net Futures OI — All Participants</div>
      <div class="section-rule"></div>
      <div class="section-tag">across loaded sessions</div>
    </div>
    <div class="chart-card wide" style="margin-bottom:12px">
      <div class="chart-sub">Net Index Futures OI (Long − Short) per participant</div>
      <div class="legend" id="trend-legend"></div>
      <svg class="chart-svg" id="chart-trend"></svg>
    </div>

    <!-- Chart row 2: OI Breakdown + DoD Change -->
    <div class="section-hdr">
      <div class="section-title">Latest Session Analysis</div>
      <div class="section-rule"></div>
      <div class="section-tag" id="chart-date-tag">—</div>
    </div>
    <div class="charts-row">
      <div class="chart-card">
        <div class="chart-title">OI Breakdown — Long vs Short</div>
        <div class="chart-sub">Total Long &amp; Short OI per participant</div>
        <svg class="chart-svg" id="chart-breakdown"></svg>
      </div>
      <div class="chart-card">
        <div class="chart-title">Day-on-Day Net OI Change</div>
        <div class="chart-sub" id="dod-sub">vs previous session</div>
        <svg class="chart-svg" id="chart-dod"></svg>
      </div>
    </div>

  </div><!-- /overview -->

  <!-- Participant OI & TV -->
  <div class="tab-panel" id="tab-participant">
    <div class="placeholder">

    <!-- P1: Net Total OI Trend (wide) -->
    <div class="section-hdr">
      <div class="section-title">Net Total OI — Participant Trend</div>
      <div class="section-rule"></div>
      <div class="section-tag">all sessions</div>
    </div>
    <div class="chart-card wide" style="margin-bottom:12px">
      <div class="chart-sub">Net Total OI (Long − Short) per participant across sessions</div>
      <div class="legend" id="part-trend-legend"></div>
      <svg class="chart-svg" id="chart-part-trend"></svg>
    </div>

    <!-- P2: Long vs Short + L/S Ratio -->
    <div class="section-hdr">
      <div class="section-title">Latest Session Detail</div>
      <div class="section-rule"></div>
      <div class="section-tag" id="part-latest-tag">—</div>
    </div>
    <div class="charts-row">
      <div class="chart-card">
        <div class="chart-title">Long vs Short OI</div>
        <div class="chart-sub">Total Long (solid) &amp; Short (dimmed) per participant</div>
        <svg class="chart-svg" id="chart-part-ls"></svg>
      </div>
      <div class="chart-card">
        <div class="chart-title">L/S Ratio Trend</div>
        <div class="chart-sub">Total Long / Total Short across sessions</div>
        <div class="legend" id="ls-ratio-legend"></div>
        <svg class="chart-svg" id="chart-ls-ratio"></svg>
      </div>
    </div>

    <!-- P3: Instrument breakdown (wide) -->
    <div class="section-hdr">
      <div class="section-title">Net OI by Instrument — Latest Session</div>
      <div class="section-rule"></div>
      <div class="section-tag">Fut Idx · Fut Stk · Idx Calls · Idx Puts · Stk Calls · Stk Puts</div>
    </div>
    <div class="chart-card wide" style="margin-bottom:12px">
      <div class="chart-sub">Net OI per instrument type per participant (green = net long, red = net short)</div>
      <svg class="chart-svg" id="chart-instrument"></svg>
    </div>

    <!-- P4: TV trend (conditional) -->
    <div id="tv-section" style="display:none">
      <div class="section-hdr">
        <div class="section-title">Net Trading Volume — Participant Trend</div>
        <div class="section-rule"></div>
        <div class="section-tag">all sessions</div>
      </div>
      <div class="chart-card wide" style="margin-bottom:12px">
        <div class="chart-sub">Net Volume (Buy − Sell) per participant across sessions</div>
        <div class="legend" id="tv-trend-legend"></div>
        <svg class="chart-svg" id="chart-tv-trend"></svg>
      </div>
    </div>

  </div>

  <!-- PCR Analysis -->
  <div class="tab-panel" id="tab-pcr">
    <div class="placeholder">

    <!-- PCR1: Index OI PCR Trend (wide) -->
    <div class="section-hdr">
      <div class="section-title">Index OI PCR — Participant Trend</div>
      <div class="section-rule"></div>
      <div class="section-tag">Put Long / Call Long</div>
    </div>
    <div class="chart-card wide" style="margin-bottom:12px">
      <div class="chart-sub">Index OI PCR per participant · reference bands: &lt;0.8 bullish · &gt;1.5 bearish</div>
      <div class="legend" id="pcr-trend-legend"></div>
      <svg class="chart-svg" id="chart-pcr-trend"></svg>
    </div>

    <!-- PCR2: Latest Session comparison + Short vs Long -->
    <div class="section-hdr">
      <div class="section-title">Latest Session PCR Analysis</div>
      <div class="section-rule"></div>
      <div class="section-tag" id="pcr-latest-tag">—</div>
    </div>
    <div class="charts-row">
      <div class="chart-card">
        <div class="chart-title">PCR by Type — All Participants</div>
        <div class="chart-sub">Index · Stock · Combined OI PCR</div>
        <svg class="chart-svg" id="chart-pcr-compare"></svg>
      </div>
      <div class="chart-card">
        <div class="chart-title">Long-side vs Short-side PCR</div>
        <div class="chart-sub">Writer's PCR vs Buyer's PCR · divergence = signal</div>
        <div class="legend" id="pcr-sides-legend"></div>
        <svg class="chart-svg" id="chart-pcr-sides"></svg>
      </div>
    </div>

    <!-- PCR3: Chain-level aggregate -->
    <div class="section-hdr">
      <div class="section-title">Chain-Level Aggregate PCR</div>
      <div class="section-rule"></div>
      <div class="section-tag">market-wide · all strikes</div>
    </div>
    <div class="metrics-strip" id="chain-pcr-strip" style="margin-bottom:28px"></div>

    <!-- PCR4: Volume PCR trend (conditional) -->
    <div id="vol-pcr-section" style="display:none">
      <div class="section-hdr">
        <div class="section-title">Index Volume PCR — Participant Trend</div>
        <div class="section-rule"></div>
        <div class="section-tag">Put Buy / Call Buy</div>
      </div>
      <div class="chart-card wide" style="margin-bottom:12px">
        <div class="chart-sub">Volume-based PCR per participant (dashed = volume, solid = OI for reference)</div>
        <div class="legend" id="vol-pcr-legend"></div>
        <svg class="chart-svg" id="chart-vol-pcr"></svg>
      </div>
    </div>

  </div>

  <!-- Net Bias & Efficiency -->
  <div class="tab-panel" id="tab-bias">
    <div class="placeholder">

    <!-- B1: OI-TV Efficiency (wide) -->
    <div class="section-hdr">
      <div class="section-title">OI / TV Efficiency Ratio</div>
      <div class="section-rule"></div>
      <div class="section-tag">OI held per unit volume · high = conviction · low = churn</div>
    </div>
    <div class="chart-card wide" style="margin-bottom:12px">
      <div class="chart-sub">Ratio of Open Interest to Trading Volume per instrument per participant</div>
      <div class="legend" id="eff-legend"></div>
      <svg class="chart-svg" id="chart-efficiency"></svg>
    </div>

    <!-- B2: Net Conviction + Net OI Composition -->
    <div class="section-hdr">
      <div class="section-title">Latest Session Analysis</div>
      <div class="section-rule"></div>
      <div class="section-tag" id="bias-latest-tag">—</div>
    </div>
    <div class="charts-row">
      <div class="chart-card">
        <div class="chart-title">Net Conviction — OI vs Volume Direction</div>
        <div class="chart-sub">Both sides same sign = strong signal · divergence = noise</div>
        <svg class="chart-svg" id="chart-conviction"></svg>
      </div>
      <div class="chart-card">
        <div class="chart-title">Net OI Composition</div>
        <div class="chart-sub">Futures vs Options split per participant</div>
        <svg class="chart-svg" id="chart-composition"></svg>
      </div>
    </div>

    <!-- B3: DoD Net OI Change (wide) -->
    <div class="section-hdr">
      <div class="section-title">Day-on-Day Net OI Change — By Instrument</div>
      <div class="section-rule"></div>
      <div class="section-tag" id="dod-instr-tag">—</div>
    </div>
    <div class="chart-card wide" style="margin-bottom:12px">
      <div class="chart-sub">Delta in Net OI per instrument type vs previous session</div>
      <svg class="chart-svg" id="chart-dod-instr"></svg>
    </div>

  </div>

  <!-- Option Chain -->
  <div class="tab-panel" id="tab-chain">
    <div class="placeholder">

    <!-- Chain metrics strip -->
    <div class="metrics-strip" id="chain-strip" style="margin-bottom:20px"></div>

    <!-- OI Profile Chart -->
    <div class="section-hdr">
      <div class="section-title">Open Interest Profile</div>
      <div class="section-rule"></div>
      <div class="section-tag">Call OI ← · → Put OI · ATM ±20 strikes shown</div>
    </div>
    <div class="chart-card wide" style="margin-bottom:20px">
      <div class="chart-sub" id="oi-profile-sub">Call OI (left) vs Put OI (right) · amber = ATM · dashed = Max Pain</div>
      <svg class="chart-svg" id="chart-oi-profile"></svg>
    </div>

    <!-- Strike Table -->
    <div class="section-hdr">
      <div class="section-title">Strike-wise Option Chain</div>
      <div class="section-rule"></div>
      <button id="chain-expand-btn" style="
        background:var(--surface2);border:1px solid var(--border2);
        color:var(--muted);border-radius:6px;padding:4px 12px;
        font-family:var(--mono);font-size:11px;cursor:pointer;">
        Show all strikes
      </button>
    </div>
    <div style="overflow-x:auto;border:1px solid var(--border);border-radius:10px;margin-bottom:28px">
      <table id="chain-table" style="
        width:100%;border-collapse:collapse;font-family:var(--mono);font-size:12px;">
        <thead id="chain-thead"></thead>
        <tbody id="chain-tbody"></tbody>
      </table>
    </div>

  </div>

  <!-- Max Pain -->
  <div class="tab-panel" id="tab-maxpain">
    <div class="placeholder">

    <!-- Max Pain metrics -->
    <div class="metrics-strip" id="mp-strip" style="margin-bottom:20px"></div>

    <!-- Pain Curve -->
    <div class="section-hdr">
      <div class="section-title">Pain Curve — All Strikes</div>
      <div class="section-rule"></div>
      <div class="section-tag">minimum = max pain strike · lower = more OI pressure at expiry</div>
    </div>
    <div class="chart-card wide" style="margin-bottom:20px">
      <div class="chart-sub">Total aggregate OI pain at each strike · amber vertical = ATM · green = Max Pain</div>
      <svg class="chart-svg" id="chart-pain-curve"></svg>
    </div>

    <!-- Breakdown + OI Region -->
    <div class="section-hdr">
      <div class="section-title">Max Pain Region Detail</div>
      <div class="section-rule"></div>
      <div class="section-tag" id="mp-region-tag">±15 / ±5 strikes around max pain</div>
    </div>
    <div class="charts-row">
      <div class="chart-card">
        <div class="chart-title">Call vs Put Pain Contribution</div>
        <div class="chart-sub">±15 strikes around max pain · stacked call + put pain</div>
        <svg class="chart-svg" id="chart-pain-breakdown"></svg>
      </div>
      <div class="chart-card">
        <div class="chart-title">OI at Max Pain Region</div>
        <div class="chart-sub">Call OI vs Put OI · ±5 strikes around max pain</div>
        <svg class="chart-svg" id="chart-mp-oi"></svg>
      </div>
    </div>

  </div>

  <!-- IV Skew -->
  <div class="tab-panel" id="tab-ivskew">
    <div class="placeholder">

    <!-- IV metrics strip -->
    <div class="metrics-strip" id="iv-strip" style="margin-bottom:20px"></div>

    <!-- IV Smile (wide) -->
    <div class="section-hdr">
      <div class="section-title">Implied Volatility Smile</div>
      <div class="section-rule"></div>
      <div class="section-tag">Call IV · Put IV · all strikes</div>
    </div>
    <div class="chart-card wide" style="margin-bottom:20px">
      <div class="chart-sub">
        <span style="color:var(--call)">▬ Call IV</span> &nbsp;
        <span style="color:var(--put)">▬ Put IV</span> &nbsp;·&nbsp;
        amber dashed = ATM · steeper left wing = put skew
      </div>
      <svg class="chart-svg" id="chart-iv-smile"></svg>
    </div>

    <!-- Differential + ATM region -->
    <div class="section-hdr">
      <div class="section-title">Skew Analysis</div>
      <div class="section-rule"></div>
      <div class="section-tag" id="iv-region-tag">Put-Call differential · ATM ±10 zoom</div>
    </div>
    <div class="charts-row">
      <div class="chart-card">
        <div class="chart-title">Put-Call IV Differential (Skew)</div>
        <div class="chart-sub">Put IV − Call IV · positive = puts expensive · red &gt;5 · blue &lt;0</div>
        <svg class="chart-svg" id="chart-iv-diff"></svg>
      </div>
      <div class="chart-card">
        <div class="chart-title">ATM Region IV — Zoomed</div>
        <div class="chart-sub" id="iv-zoom-sub">ATM ±10 strikes · tighter scale</div>
        <svg class="chart-svg" id="chart-iv-zoom"></svg>
      </div>
    </div>

  </div>

  <!-- Historical Trends -->
  <div class="tab-panel" id="tab-history">
    <div class="placeholder">

    <!-- H0: sessions summary strip -->
    <div class="metrics-strip" id="hist-strip" style="margin-bottom:20px"></div>

    <!-- H1: FII Net Index Futures (wide, prominent) -->
    <div class="section-hdr">
      <div class="section-title">FII Net Index Futures OI</div>
      <div class="section-rule"></div>
      <div class="section-tag">primary directional signal · all sessions</div>
    </div>
    <div class="chart-card wide" style="margin-bottom:20px">
      <div class="chart-sub">FII Long − Short on Index Futures · positive = net long · negative = net short</div>
      <svg class="chart-svg" id="chart-fii-trend"></svg>
    </div>

    <!-- H2: Market Share + PCR -->
    <div class="section-hdr">
      <div class="section-title">Market Dynamics</div>
      <div class="section-rule"></div>
      <div class="section-tag">share of total OI · Index PCR trend</div>
    </div>
    <div class="charts-row" style="margin-bottom:20px">
      <div class="chart-card">
        <div class="chart-title">Market OI Share — Long %</div>
        <div class="chart-sub">Each participant's % of total long OI across sessions</div>
        <svg class="chart-svg" id="chart-mkt-share"></svg>
      </div>
      <div class="chart-card">
        <div class="chart-title">Index OI PCR — All Participants</div>
        <div class="chart-sub">Put Long / Call Long · 0.8 bullish · 1.5 bearish</div>
        <div class="legend" id="hist-pcr-legend"></div>
        <svg class="chart-svg" id="chart-hist-pcr"></svg>
      </div>
    </div>

    <!-- H3: DoD Changes Summary (wide) -->
    <div class="section-hdr">
      <div class="section-title">Day-on-Day Net OI Changes</div>
      <div class="section-rule"></div>
      <div class="section-tag">ΔNet Total OI per participant per session pair</div>
    </div>
    <div class="chart-card wide" style="margin-bottom:20px">
      <div class="chart-sub">Change in Net Total OI between consecutive sessions · green = net long added · red = net short added</div>
      <div class="legend" id="hist-dod-legend"></div>
      <svg class="chart-svg" id="chart-hist-dod"></svg>
    </div>

    <!-- H4: Key Signals Table -->
    <div class="section-hdr">
      <div class="section-title">Key Signals — All Sessions</div>
      <div class="section-rule"></div>
      <div class="section-tag">net fut idx · pcr · l/s ratio · signal</div>
    </div>
    <div style="overflow-x:auto;border:1px solid var(--border);border-radius:10px;margin-bottom:28px">
      <table id="hist-table" style="width:100%;border-collapse:collapse;font-family:var(--mono);font-size:11px;">
        <thead id="hist-thead"></thead>
        <tbody id="hist-tbody"></tbody>
      </table>
    </div>

  </div>

</main>

<!-- ── Data badge ────────────────────────────────────────────────────── -->
<div class="d3-tip" id="d3-tip" style="display:none"></div>
<footer class="site-footer" id="site-footer" style="display:none">
  <span id="footer-gen">Generated —</span>
  <span class="footer-sep">·</span>
  <span id="footer-sessions">— sessions</span>
  <span class="footer-sep">·</span>
  <span id="footer-strikes">— strikes</span>
  <span style="flex:1"></span>
  <a href="https://github.com/enemyatgates/OptionsDashboardNSE" target="_blank">enemyatgates/OptionsDashboardNSE</a>
</footer>
<div class="data-badge" id="data-badge">
  <div class="data-badge-item">
    <span class="data-badge-lbl">generated</span>
    <span id="badge-gen">—</span>
  </div>
  <span class="data-badge-sep">·</span>
  <div class="data-badge-item">
    <span class="data-badge-lbl">sessions</span>
    <span id="badge-sessions">—</span>
  </div>
  <span class="data-badge-sep">·</span>
  <div class="data-badge-item">
    <span class="data-badge-lbl">strikes</span>
    <span id="badge-strikes">—</span>
  </div>
</div>

<script>
/* ── Tab switching ─────────────────────────────────────────────────── */
document.querySelectorAll('.tab-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    const id = btn.dataset.tab;
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById('tab-' + id).classList.add('active');
    // Re-render the activated tab so charts use correct clientWidth
    if (window._dashData) {
      const renderMap = {
        overview:    () => { drawTrend(window._dashData.oi.data, window._dashData.meta.oi_dates);
                             drawBreakdown(window._dashData.oi.data, window._dashData.meta.latest_date);
                             drawDoD(window._dashData.dod, (window._dashData.dod||{}).pairs||[]); },
        participant: () => renderParticipantTab(window._dashData),
        pcr:         () => renderPCRTab(window._dashData),
        bias:        () => renderBiasTab(window._dashData),
        chain:       () => renderChainTab(window._dashData),
        maxpain:     () => renderMaxPainTab(window._dashData),
        ivskew:      () => renderIVSkewTab(window._dashData),
        history:     () => renderHistoryTab(window._dashData),
      };
      if (renderMap[id]) renderMap[id]();
    }
  });
});

/* ── Formatters ────────────────────────────────────────────────────── */
const fmtN  = v => v == null ? '—' : d3.format(',')(Math.round(v));
const fmtF  = v => v == null ? '—' : d3.format('.3f')(v);
const fmtF2 = v => v == null ? '—' : d3.format('.2f')(v);

/* ── Sentiment colours ─────────────────────────────────────────────── */
const SIG_COL  = { Bullish:'#10B981', Cautious:'#F97316', Bearish:'#EF4444', Neutral:'#94A3B8' };
const SIG_ICON = { Bullish:'🟢', Cautious:'🟠', Bearish:'🔴', Neutral:'🟡' };

/* ── Load data.json ────────────────────────────────────────────────── */
const setStatus = (state, msg) => {
  const dot  = document.getElementById('status-dot');
  const text = document.getElementById('status-text');
  dot.className  = 'status-dot ' + state;
  text.textContent = msg;
};

window._dashData = null;
/* ── Fetch meta.json for Excel download link ─────────────────────── */
async function loadMeta() {
  try {
    const res  = await fetch('./meta.json');
    if (!res.ok) return;
    const meta = await res.json();
    if (meta.excel_filename) {
      const link = document.getElementById('nav-dl');
      link.href  = `./outputs/${meta.excel_filename}`;
      link.title = `Download ${meta.excel_filename}`;
    }
  } catch(_) {}
}
loadMeta();

async function loadData() {
  setStatus('loading', 'Loading data…');
  let data;
  try {
    const res = await fetch('./data.json');
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    data = await res.json();
    setStatus('ok', 'Data loaded');
    console.log('[FAO Claude] data.json loaded:', data);
  } catch (err) {
    setStatus('error', 'Failed to load data.json');
    console.error('[FAO Claude] fetch error:', err);
    document.getElementById('tab-overview').innerHTML =
      `<div class="error-banner">
         ⚠ Could not load <code>data.json</code>. Run the dashboard generator and push to <code>data/</code> to trigger a rebuild.<br/>
         <small style="opacity:.6">${err.message}</small>
       </div>`;
    return;
  }
  render(data);
}

/* ── Main render ───────────────────────────────────────────────────── */
function render(data) {
  const meta  = data.meta  || {};
  const chain = data.chain || {};
  const sent  = data.sentiment || {};
  const latest = meta.latest_date || '—';
  const dates  = meta.oi_dates   || [];

  /* Nav meta already server-rendered; update badge */
  document.getElementById('badge-gen').textContent      = (meta.generated_at || '').slice(0,10);
  document.getElementById('badge-sessions').textContent = dates.length;
  document.getElementById('badge-strikes').textContent  = (chain.data || []).length;

  /* ── Overview metrics ── */
  document.getElementById('ov-maxpain').textContent     = fmtN(chain.max_pain);
  document.getElementById('ov-atm').textContent         = fmtN(chain.atm);
  document.getElementById('ov-pcr').textContent         = fmtF(chain.oi_pcr);
  document.getElementById('ov-sessions').textContent    = dates.length;
  document.getElementById('ov-maxpain-sub').textContent = latest;
  document.getElementById('ov-atm-sub').textContent     = 'Highest combined OI';
  document.getElementById('ov-pcr-sub').textContent     =
    `Put ${fmtN(chain.total_put_oi)} / Call ${fmtN(chain.total_call_oi)}`;
  document.getElementById('ov-dates-sub').textContent   =
    dates.length ? dates[0] + ' → ' + dates[dates.length-1] : '—';

  const callWalls = chain.call_walls || [];
  const putWalls  = chain.put_walls  || [];
  document.getElementById('ov-callwall').textContent     = fmtN(callWalls[0]?.strike);
  document.getElementById('ov-putwall').textContent      = fmtN(putWalls[0]?.strike);
  document.getElementById('ov-callwall-sub').textContent =
    callWalls[0] ? `OI: ${fmtN(callWalls[0].oi)}` : '—';
  document.getElementById('ov-putwall-sub').textContent  =
    putWalls[0] ? `OI: ${fmtN(putWalls[0].oi)}` : '—';

  /* ── Sentiment table ── */
  document.getElementById('sent-date-tag').textContent = latest;
  const sentLatest = (sent.data || {})[latest] || {};
  const tbody = document.getElementById('sent-tbody');
  tbody.innerHTML = '';
  ['Client','DII','FII','Pro'].forEach(part => {
    const s    = sentLatest[part] || {};
    const sig  = s.signal || 'Neutral';
    const col  = SIG_COL[sig]  || '#94A3B8';
    const icon = SIG_ICON[sig] || '🟡';
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="font-weight:600">${part}</td>
      <td>${s.net_fut_idx_dir || '—'}</td>
      <td>${fmtF2(s.idx_oi_pcr)}</td>
      <td>${fmtF2(s.ls_ratio)}</td>
      <td style="color:${col};font-weight:600">${icon} ${sig}</td>`;
    tbody.appendChild(tr);
  });

  /* ── OI Walls ── */
  const maxCallOI = Math.max(...callWalls.map(w => w.oi || 0), 1);
  const maxPutOI  = Math.max(...putWalls.map(w => w.oi  || 0), 1);

  function wallRows(walls, maxOI, colour) {
    return walls.slice(0, 5).map((w, i) => {
      const pct = ((w.oi || 0) / maxOI * 100).toFixed(0);
      return `
        <div class="wall-row">
          <span class="wall-rank">${i+1}</span>
          <span class="wall-strike">${fmtN(w.strike)}</span>
          <span class="wall-oi">${fmtN(w.oi)}</span>
          <div class="wall-bar-wrap">
            <div class="wall-bar" style="width:${pct}%;background:${colour}"></div>
          </div>
        </div>`;
    }).join('');
  }
  document.getElementById('call-walls-body').innerHTML =
    wallRows(callWalls, maxCallOI, 'var(--call)');
  document.getElementById('put-walls-body').innerHTML  =
    wallRows(putWalls,  maxPutOI,  'var(--put)');


  /* ── Chart constants ── */
  const PART_COLORS = {
    Client: '#94A3B8',
    DII:    '#10B981',
    FII:    '#3B82F6',
    Pro:    '#F97316',
  };
  const PARTS = ['Client','DII','FII','Pro'];
  const tip   = document.getElementById('d3-tip');

  function showTip(html, event) {
    tip.innerHTML = html;
    tip.style.display = 'block';
    moveTip(event);
  }
  function moveTip(event) {
    const pad  = 16;
    const tw   = tip.offsetWidth  || 180;
    const th   = tip.offsetHeight || 80;
    const vw   = window.innerWidth;
    const vh   = window.innerHeight + window.scrollY;
    let x = event.pageX + 14;
    let y = event.pageY - 28;
    if (x + tw + pad > vw)  x = event.pageX - tw - 10;
    if (y + th + pad > vh)  y = event.pageY - th - 10;
    if (x < pad) x = pad;
    if (y < pad) y = pad;
    tip.style.left = x + 'px';
    tip.style.top  = y + 'px';
  }
  function hideTip() { tip.style.display = 'none'; }

  /* ── Chart 1: Net Futures OI Trend ── */
  function drawTrend(oiData, dates) {
    if (!dates || dates.length < 2) return;
    document.getElementById('chart-date-tag').textContent = dates[dates.length-1];

    const el   = document.getElementById('chart-trend');
    const W    = el.parentElement.clientWidth - 32;
    const H    = 220;
    const mg   = { top: 12, right: 20, bottom: 32, left: 80 };
    const iw   = W - mg.left - mg.right;
    const ih   = H - mg.top  - mg.bottom;

    el.setAttribute('viewBox', `0 0 ${W} ${H}`);
    el.setAttribute('height', H);

    const svg = d3.select(el);
    svg.selectAll('*').remove();
    const g = svg.append('g').attr('transform', `translate(${mg.left},${mg.top})`);

    // Build series: {part, values:[{date, val}]}
    const series = PARTS.map(part => ({
      part,
      values: dates.map(d => ({
        date: d,
        val:  (((oiData[d] || {})[part] || {}).net_fut_idx) ?? 0,
      })),
    }));

    const allVals = series.flatMap(s => s.values.map(v => v.val));
    const yMax    = Math.max(Math.abs(d3.max(allVals)), Math.abs(d3.min(allVals)), 1);

    const xScale = d3.scalePoint().domain(dates).range([0, iw]).padding(0.1);
    const yScale = d3.scaleLinear().domain([-yMax * 1.1, yMax * 1.1]).range([ih, 0]);

    // Grid
    g.append('g').attr('class','grid')
      .call(d3.axisLeft(yScale).tickSize(-iw).tickFormat('').ticks(5));

    // Zero line
    g.append('line')
      .attr('x1', 0).attr('x2', iw)
      .attr('y1', yScale(0)).attr('y2', yScale(0))
      .attr('stroke','var(--border2)').attr('stroke-dasharray','4,3');

    // Axes
    g.append('g').attr('class','axis')
      .attr('transform', `translate(0,${ih})`)
      .call(d3.axisBottom(xScale).tickSize(3));

    g.append('g').attr('class','axis')
      .call(d3.axisLeft(yScale).ticks(5)
        .tickFormat(v => {
          const a = Math.abs(v);
          return (a >= 1e6 ? d3.format('.1f')(v/1e6)+'M'
                : a >= 1e3 ? d3.format('.0f')(v/1e3)+'K' : v);
        }));

    // Lines
    const line = d3.line()
      .x(d => xScale(d.date))
      .y(d => yScale(d.val))
      .curve(d3.curveMonotoneX);

    series.forEach(s => {
      g.append('path')
        .datum(s.values)
        .attr('fill','none')
        .attr('stroke', PART_COLORS[s.part])
        .attr('stroke-width', 2)
        .attr('d', line);

      // Dots
      g.selectAll(`.dot-${s.part}`)
        .data(s.values)
        .join('circle')
        .attr('cx', d => xScale(d.date))
        .attr('cy', d => yScale(d.val))
        .attr('r', 4)
        .attr('fill', PART_COLORS[s.part])
        .attr('stroke','var(--surface)').attr('stroke-width',2)
        .style('cursor','pointer')
        .on('mouseover', (ev, d) => {
          const dir = d.val > 0 ? '▲ Net Long' : '▼ Net Short';
          showTip(`<b style="color:${PART_COLORS[s.part]}">${s.part}</b><br/>
            ${d.date}<br/>Net Fut Idx: <b>${fmtN(d.val)}</b><br/>${dir}`, ev);
        })
        .on('mousemove', moveTip)
        .on('mouseout',  hideTip);
    });

    // Legend
    const lgnd = document.getElementById('trend-legend');
    lgnd.innerHTML = '';
    PARTS.forEach(p => {
      lgnd.innerHTML += `<div class="legend-item">
        <div class="legend-line" style="background:${PART_COLORS[p]}"></div>${p}</div>`;
    });
  }

  /* ── Chart 2: OI Breakdown — Long vs Short ── */
  function drawBreakdown(oiData, latest) {
    const el = document.getElementById('chart-breakdown');
    const W  = Math.max(el.parentElement.clientWidth - 32, 860);
    const H  = 180;
    const mg = { top: 8, right: 16, bottom: 20, left: 52 };
    const iw = W - mg.left - mg.right;
    const ih = H - mg.top  - mg.bottom;

    el.setAttribute('viewBox', `0 0 ${W} ${H}`);
    el.setAttribute('height', H);
    const svg = d3.select(el);
    svg.selectAll('*').remove();
    const g = svg.append('g').attr('transform', `translate(${mg.left},${mg.top})`);

    const rows = PARTS.map(p => {
      const rec = ((oiData[latest] || {})[p]) || {};
      return { part: p, long: rec.total_long || 0, short: rec.total_short || 0 };
    });

    const xMax = d3.max(rows, r => Math.max(r.long, r.short)) * 1.05;
    const yBand = d3.scaleBand().domain(PARTS).range([0, ih]).padding(0.28);
    const xScale = d3.scaleLinear().domain([0, xMax]).range([0, iw]);

    g.append('g').attr('class','axis')
      .attr('transform', `translate(0,${ih})`)
      .call(d3.axisBottom(xScale).ticks(4)
        .tickFormat(v => v >= 1e6 ? d3.format('.1f')(v/1e6)+'M'
                       : v >= 1e3 ? d3.format('.0f')(v/1e3)+'K' : v));

    g.append('g').attr('class','axis').call(d3.axisLeft(yBand).tickSize(0))
      .select('.domain').remove();

    const barH = yBand.bandwidth() / 2 - 1;

    rows.forEach(r => {
      const y0 = yBand(r.part);
      // Long bar
      g.append('rect')
        .attr('x', 0).attr('y', y0)
        .attr('width', xScale(r.long)).attr('height', barH)
        .attr('fill', PART_COLORS[r.part]).attr('opacity', 0.85)
        .attr('rx', 2)
        .on('mouseover', ev => showTip(
          `<b style="color:${PART_COLORS[r.part]}">${r.part}</b><br/>
           Long OI: <b>${fmtN(r.long)}</b>`, ev))
        .on('mousemove', moveTip).on('mouseout', hideTip);
      // Short bar
      g.append('rect')
        .attr('x', 0).attr('y', y0 + barH + 2)
        .attr('width', xScale(r.short)).attr('height', barH)
        .attr('fill', PART_COLORS[r.part]).attr('opacity', 0.35)
        .attr('rx', 2)
        .on('mouseover', ev => showTip(
          `<b style="color:${PART_COLORS[r.part]}">${r.part}</b><br/>
           Short OI: <b>${fmtN(r.short)}</b>`, ev))
        .on('mousemove', moveTip).on('mouseout', hideTip);
    });

    // Tiny legend
    const lgEl = el.parentElement.querySelector('.chart-sub');
    if (lgEl) lgEl.innerHTML =
      `Total Long &amp; Short OI &nbsp;
       <span style="opacity:.85">▬</span> Long &nbsp;
       <span style="opacity:.35">▬</span> Short`;
  }

  /* ── Chart 3: Day-on-Day Net OI Change ── */
  function drawDoD(dodData, pairs) {
    if (!pairs || pairs.length === 0) return;
    const [d0, d1] = pairs[pairs.length - 1];
    const pairKey  = `${d0}|${d1}`;
    const pairData = (dodData.data || {})[pairKey] || {};
    document.getElementById('dod-sub').textContent = `${d0} → ${d1}`;

    const el = document.getElementById('chart-dod');
    const W  = Math.max(el.parentElement.clientWidth - 32, 860);
    const H  = 180;
    const mg = { top: 8, right: 16, bottom: 20, left: 52 };
    const iw = W - mg.left - mg.right;
    const ih = H - mg.top  - mg.bottom;

    el.setAttribute('viewBox', `0 0 ${W} ${H}`);
    el.setAttribute('height', H);
    const svg = d3.select(el);
    svg.selectAll('*').remove();
    const g = svg.append('g').attr('transform', `translate(${mg.left},${mg.top})`);

    const rows = PARTS.map(p => {
      const rec = pairData[p] || {};
      return { part: p, val: rec.d_net_total ?? 0 };
    });

    const absMax = Math.max(d3.max(rows, r => Math.abs(r.val)), 1);
    const yBand  = d3.scaleBand().domain(PARTS).range([0, ih]).padding(0.3);
    const xScale = d3.scaleLinear().domain([-absMax * 1.1, absMax * 1.1]).range([0, iw]);
    const xMid   = xScale(0);

    // Grid + zero
    g.append('line')
      .attr('x1', xMid).attr('x2', xMid)
      .attr('y1', 0).attr('y2', ih)
      .attr('stroke','var(--border2)').attr('stroke-dasharray','4,3');

    g.append('g').attr('class','axis')
      .attr('transform', `translate(0,${ih})`)
      .call(d3.axisBottom(xScale).ticks(5)
        .tickFormat(v => v >= 1e6 ? d3.format('.1f')(v/1e6)+'M'
                       : v >= 1e3 ? d3.format('.0f')(v/1e3)+'K'
                       : v <= -1e3 ? d3.format('.0f')(v/1e3)+'K' : v));

    g.append('g').attr('class','axis').call(d3.axisLeft(yBand).tickSize(0))
      .select('.domain').remove();

    rows.forEach(r => {
      const positive = r.val >= 0;
      const barW  = Math.abs(xScale(r.val) - xMid);
      const barX  = positive ? xMid : xMid - barW;
      const color = positive ? '#10B981' : '#EF4444';

      g.append('rect')
        .attr('x', barX).attr('y', yBand(r.part))
        .attr('width', Math.max(barW, 1)).attr('height', yBand.bandwidth())
        .attr('fill', color).attr('opacity', 0.8).attr('rx', 2)
        .on('mouseover', ev => showTip(
          `<b style="color:${PART_COLORS[r.part]}">${r.part}</b><br/>
           Δ Net Total OI: <b style="color:${color}">${fmtN(r.val)}</b>`, ev))
        .on('mousemove', moveTip).on('mouseout', hideTip);
    });
  }


  /* ════════════════════════════════════════════════════════
     PARTICIPANT OI & TV CHARTS
  ════════════════════════════════════════════════════════ */

  /* Shared axis tick formatter */
  const fmtAxis = v => {
    const a = Math.abs(v);
    return a >= 1e6 ? d3.format('.1f')(v/1e6)+'M'
         : a >= 1e3 ? d3.format('.0f')(v/1e3)+'K'
         : String(v);
  };

  /* ── P1: Net Total OI Trend ── */
  function drawPartTrend(oiData, dates) {
    if (!dates || dates.length < 1) return;
    const el = document.getElementById('chart-part-trend');
    const W  = Math.max(el.parentElement.clientWidth - 32, 860);
    const H  = 230;
    const mg = { top: 12, right: 20, bottom: 32, left: 80 };
    const iw = W - mg.left - mg.right;
    const ih = H - mg.top  - mg.bottom;

    el.setAttribute('viewBox', `0 0 ${W} ${H}`);
    el.setAttribute('height', H);
    const svg = d3.select(el); svg.selectAll('*').remove();
    const g   = svg.append('g').attr('transform', `translate(${mg.left},${mg.top})`);

    const series = PARTS.map(p => ({
      part: p,
      values: dates.map(d => ({
        date: d,
        val: ((oiData[d] || {})[p] || {}).net_total ?? 0,
      })),
    }));
    const allVals = series.flatMap(s => s.values.map(v => v.val));
    const yMax = Math.max(Math.abs(d3.max(allVals)), Math.abs(d3.min(allVals)), 1);

    const xSc = d3.scalePoint().domain(dates).range([0, iw]).padding(0.1);
    const ySc = d3.scaleLinear().domain([-yMax*1.12, yMax*1.12]).range([ih, 0]);

    g.append('g').attr('class','grid')
     .call(d3.axisLeft(ySc).tickSize(-iw).tickFormat('').ticks(5));
    g.append('line')
     .attr('x1',0).attr('x2',iw).attr('y1',ySc(0)).attr('y2',ySc(0))
     .attr('stroke','var(--border2)').attr('stroke-dasharray','4,3');
    g.append('g').attr('class','axis').attr('transform',`translate(0,${ih})`)
     .call(d3.axisBottom(xSc).tickSize(3));
    g.append('g').attr('class','axis')
     .call(d3.axisLeft(ySc).ticks(5).tickFormat(fmtAxis));

    const line = d3.line()
      .x(d => xSc(d.date)).y(d => ySc(d.val))
      .curve(d3.curveMonotoneX);

    series.forEach(s => {
      g.append('path').datum(s.values)
        .attr('fill','none').attr('stroke', PART_COLORS[s.part])
        .attr('stroke-width', 2.5).attr('d', line);
      g.selectAll(null).data(s.values).join('circle')
        .attr('cx', d => xSc(d.date)).attr('cy', d => ySc(d.val))
        .attr('r', 4.5).attr('fill', PART_COLORS[s.part])
        .attr('stroke','var(--surface)').attr('stroke-width', 2)
        .style('cursor','pointer')
        .on('mouseover', (ev, d) => showTip(
          `<b style="color:${PART_COLORS[s.part]}">${s.part}</b><br/>
           ${d.date}<br/>Net Total OI: <b>${fmtN(d.val)}</b>`, ev))
        .on('mousemove', moveTip).on('mouseout', hideTip);
    });

    const lg = document.getElementById('part-trend-legend');
    lg.innerHTML = PARTS.map(p =>
      `<div class="legend-item">
         <div class="legend-line" style="background:${PART_COLORS[p]}"></div>${p}
       </div>`).join('');
  }

  /* ── P2a: Long vs Short OI — latest session ── */
  function drawPartLS(oiData, latest) {
    const el = document.getElementById('chart-part-ls');
    const W  = Math.max(el.parentElement.clientWidth - 32, 860);
    const H  = 190;
    const mg = { top: 8, right: 16, bottom: 28, left: 56 };
    const iw = W - mg.left - mg.right;
    const ih = H - mg.top  - mg.bottom;

    el.setAttribute('viewBox', `0 0 ${W} ${H}`);
    el.setAttribute('height', H);
    const svg = d3.select(el); svg.selectAll('*').remove();
    const g   = svg.append('g').attr('transform', `translate(${mg.left},${mg.top})`);

    const rows = PARTS.map(p => {
      const r = ((oiData[latest] || {})[p]) || {};
      return { part: p, long: r.total_long || 0, short: r.total_short || 0 };
    });
    const xMax = d3.max(rows, r => Math.max(r.long, r.short)) * 1.05;
    const yBand = d3.scaleBand().domain(PARTS).range([0, ih]).padding(0.26);
    const xSc   = d3.scaleLinear().domain([0, xMax]).range([0, iw]);
    const barH  = yBand.bandwidth() / 2 - 1;

    g.append('g').attr('class','grid')
     .attr('transform',`translate(0,${ih})`)
     .call(d3.axisBottom(xSc).tickSize(-ih).tickFormat('').ticks(4));
    g.append('g').attr('class','axis').attr('transform',`translate(0,${ih})`)
     .call(d3.axisBottom(xSc).ticks(4).tickFormat(fmtAxis));
    g.append('g').attr('class','axis').call(d3.axisLeft(yBand).tickSize(0))
     .select('.domain').remove();

    rows.forEach(r => {
      const y0 = yBand(r.part);
      [[r.long, 0.9, 'Long'], [r.short, 0.35, 'Short']].forEach(([val, op, lbl], i) => {
        g.append('rect')
          .attr('x', 0).attr('y', y0 + i * (barH + 2))
          .attr('width', xSc(val)).attr('height', barH)
          .attr('fill', PART_COLORS[r.part]).attr('opacity', op).attr('rx', 2)
          .on('mouseover', ev => showTip(
            `<b style="color:${PART_COLORS[r.part]}">${r.part}</b><br/>
             ${lbl} OI: <b>${fmtN(val)}</b>`, ev))
          .on('mousemove', moveTip).on('mouseout', hideTip);
      });
    });
  }

  /* ── P2b: L/S Ratio Trend ── */
  function drawLSRatio(oiData, dates) {
    if (!dates || dates.length < 1) return;
    const el = document.getElementById('chart-ls-ratio');
    const W  = Math.max(el.parentElement.clientWidth - 32, 860);
    const H  = 190;
    const mg = { top: 8, right: 20, bottom: 28, left: 52 };
    const iw = W - mg.left - mg.right;
    const ih = H - mg.top  - mg.bottom;

    el.setAttribute('viewBox', `0 0 ${W} ${H}`);
    el.setAttribute('height', H);
    const svg = d3.select(el); svg.selectAll('*').remove();
    const g   = svg.append('g').attr('transform', `translate(${mg.left},${mg.top})`);

    const series = PARTS.map(p => ({
      part: p,
      values: dates.map(d => ({
        date: d,
        val: ((oiData[d] || {})[p] || {}).ls_ratio ?? 1,
      })),
    }));
    const allVals = series.flatMap(s => s.values.map(v => v.val)).filter(v => v != null);
    const [vMin, vMax] = [Math.min(...allVals) * 0.95, Math.max(...allVals) * 1.05];

    const xSc = d3.scalePoint().domain(dates).range([0, iw]).padding(0.1);
    const ySc = d3.scaleLinear().domain([Math.min(vMin, 0.8), Math.max(vMax, 1.2)]).range([ih, 0]);

    g.append('g').attr('class','grid')
     .call(d3.axisLeft(ySc).tickSize(-iw).tickFormat('').ticks(5));
    // parity line at 1.0
    g.append('line')
     .attr('x1',0).attr('x2',iw).attr('y1',ySc(1)).attr('y2',ySc(1))
     .attr('stroke','var(--border2)').attr('stroke-dasharray','4,3');
    g.append('g').attr('class','axis').attr('transform',`translate(0,${ih})`)
     .call(d3.axisBottom(xSc).tickSize(3));
    g.append('g').attr('class','axis')
     .call(d3.axisLeft(ySc).ticks(5).tickFormat(v => d3.format('.2f')(v)));

    const line = d3.line()
      .x(d => xSc(d.date)).y(d => ySc(d.val))
      .curve(d3.curveMonotoneX);

    series.forEach(s => {
      g.append('path').datum(s.values)
        .attr('fill','none').attr('stroke', PART_COLORS[s.part])
        .attr('stroke-width', 2).attr('d', line);
      g.selectAll(null).data(s.values).join('circle')
        .attr('cx', d => xSc(d.date)).attr('cy', d => ySc(d.val))
        .attr('r', 3.5).attr('fill', PART_COLORS[s.part])
        .attr('stroke','var(--surface)').attr('stroke-width', 1.5)
        .on('mouseover', (ev, d) => showTip(
          `<b style="color:${PART_COLORS[s.part]}">${s.part}</b><br/>
           ${d.date}<br/>L/S Ratio: <b>${d3.format('.3f')(d.val)}</b>`, ev))
        .on('mousemove', moveTip).on('mouseout', hideTip);
    });

    const lg = document.getElementById('ls-ratio-legend');
    lg.innerHTML = PARTS.map(p =>
      `<div class="legend-item">
         <div class="legend-line" style="background:${PART_COLORS[p]}"></div>${p}
       </div>`).join('');
  }

  /* ── P3: Instrument Breakdown — latest session ── */
  function drawInstrument(oiData, latest) {
    const INSTRUMENTS = [
      { key: 'net_fut_idx',   label: 'Fut Index' },
      { key: 'net_fut_stk',   label: 'Fut Stock'  },
      { key: 'net_idx_call',  label: 'Idx Call'   },
      { key: 'net_idx_put',   label: 'Idx Put'    },
      { key: 'net_stk_call',  label: 'Stk Call'   },
      { key: 'net_stk_put',   label: 'Stk Put'    },
    ];
    const rows = [];
    PARTS.forEach(p => {
      const rec = ((oiData[latest] || {})[p]) || {};
      INSTRUMENTS.forEach(ins => {
        rows.push({ part: p, label: `${p} · ${ins.label}`, val: rec[ins.key] ?? 0 });
      });
    });

    const el = document.getElementById('chart-instrument');
    const W  = Math.max(el.parentElement.clientWidth - 32, 860);
    const H  = Math.max(rows.length * 22 + 40, 220);
    const mg = { top: 8, right: 20, bottom: 28, left: 110 };
    const iw = W - mg.left - mg.right;
    const ih = H - mg.top  - mg.bottom;

    el.setAttribute('viewBox', `0 0 ${W} ${H}`);
    el.setAttribute('height', H);
    const svg = d3.select(el); svg.selectAll('*').remove();
    const g   = svg.append('g').attr('transform', `translate(${mg.left},${mg.top})`);

    const absMax = Math.max(d3.max(rows, r => Math.abs(r.val)), 1);
    const yBand  = d3.scaleBand().domain(rows.map(r => r.label)).range([0, ih]).padding(0.18);
    const xSc    = d3.scaleLinear().domain([-absMax*1.1, absMax*1.1]).range([0, iw]);
    const xMid   = xSc(0);

    g.append('g').attr('class','grid')
     .attr('transform',`translate(0,${ih})`)
     .call(d3.axisBottom(xSc).tickSize(-ih).tickFormat('').ticks(6));
    g.append('line')
     .attr('x1',xMid).attr('x2',xMid).attr('y1',0).attr('y2',ih)
     .attr('stroke','var(--border2)').attr('stroke-dasharray','4,3');
    g.append('g').attr('class','axis').attr('transform',`translate(0,${ih})`)
     .call(d3.axisBottom(xSc).ticks(6).tickFormat(fmtAxis));
    g.append('g').attr('class','axis').call(d3.axisLeft(yBand).tickSize(0))
     .selectAll('text').style('font-size','9px');
    d3.select(el).select('.axis:last-of-type .domain').remove();

    rows.forEach(r => {
      const pos  = r.val >= 0;
      const barW = Math.max(Math.abs(xSc(r.val) - xMid), 1);
      const barX = pos ? xMid : xMid - barW;
      const col  = PART_COLORS[r.part];

      g.append('rect')
        .attr('x', barX).attr('y', yBand(r.label))
        .attr('width', barW).attr('height', yBand.bandwidth())
        .attr('fill', col).attr('opacity', pos ? 0.8 : 0.4).attr('rx', 2)
        .on('mouseover', ev => showTip(
          `<b style="color:${col}">${r.label}</b><br/>
           Net OI: <b>${fmtN(r.val)}</b><br/>
           ${pos ? '▲ Net Long' : '▼ Net Short'}`, ev))
        .on('mousemove', moveTip).on('mouseout', hideTip);
    });
  }

  /* ── P4: Net TV Trend (conditional) ── */
  function drawTVTrend(tvData, dates) {
    const tvDates = dates.filter(d => tvData[d]);
    if (!tvDates.length) return;
    document.getElementById('tv-section').style.display = 'block';

    const el = document.getElementById('chart-tv-trend');
    const W  = Math.max(el.parentElement.clientWidth - 32, 860);
    const H  = 200;
    const mg = { top: 12, right: 20, bottom: 32, left: 80 };
    const iw = W - mg.left - mg.right;
    const ih = H - mg.top  - mg.bottom;

    el.setAttribute('viewBox', `0 0 ${W} ${H}`);
    el.setAttribute('height', H);
    const svg = d3.select(el); svg.selectAll('*').remove();
    const g   = svg.append('g').attr('transform', `translate(${mg.left},${mg.top})`);

    const series = PARTS.map(p => ({
      part: p,
      values: tvDates.map(d => ({
        date: d,
        val: ((tvData[d] || {})[p] || {}).net_total ?? 0,
      })),
    }));
    const allVals = series.flatMap(s => s.values.map(v => v.val));
    const yMax = Math.max(Math.abs(d3.max(allVals)), Math.abs(d3.min(allVals)), 1);

    const xSc = d3.scalePoint().domain(tvDates).range([0, iw]).padding(0.1);
    const ySc = d3.scaleLinear().domain([-yMax*1.12, yMax*1.12]).range([ih, 0]);

    g.append('g').attr('class','grid')
     .call(d3.axisLeft(ySc).tickSize(-iw).tickFormat('').ticks(5));
    g.append('line')
     .attr('x1',0).attr('x2',iw).attr('y1',ySc(0)).attr('y2',ySc(0))
     .attr('stroke','var(--border2)').attr('stroke-dasharray','4,3');
    g.append('g').attr('class','axis').attr('transform',`translate(0,${ih})`)
     .call(d3.axisBottom(xSc).tickSize(3));
    g.append('g').attr('class','axis')
     .call(d3.axisLeft(ySc).ticks(5).tickFormat(fmtAxis));

    const line = d3.line()
      .x(d => xSc(d.date)).y(d => ySc(d.val))
      .curve(d3.curveMonotoneX);

    series.forEach(s => {
      g.append('path').datum(s.values)
        .attr('fill','none').attr('stroke', PART_COLORS[s.part])
        .attr('stroke-width', 2).attr('stroke-dasharray','5,3')
        .attr('d', line);
      g.selectAll(null).data(s.values).join('circle')
        .attr('cx', d => xSc(d.date)).attr('cy', d => ySc(d.val))
        .attr('r', 3.5).attr('fill', PART_COLORS[s.part])
        .attr('stroke','var(--surface)').attr('stroke-width', 1.5)
        .on('mouseover', (ev, d) => showTip(
          `<b style="color:${PART_COLORS[s.part]}">${s.part}</b><br/>
           ${d.date}<br/>Net Volume: <b>${fmtN(d.val)}</b>`, ev))
        .on('mousemove', moveTip).on('mouseout', hideTip);
    });

    const lg = document.getElementById('tv-trend-legend');
    lg.innerHTML = PARTS.map(p =>
      `<div class="legend-item">
         <div class="legend-line" style="background:${PART_COLORS[p]};
           border-bottom:2px dashed ${PART_COLORS[p]};height:0"></div>${p}
       </div>`).join('');
  }

  /* ── Wire participant tab ── */
  function renderParticipantTab(data) {
    const oiData = (data.oi || {}).data || {};
    const tvData = (data.tv || {}).data || {};
    const dates  = (data.oi || {}).dates || [];
    const latest = dates[dates.length - 1];
    if (!latest) return;

    document.getElementById('part-latest-tag').textContent = latest;
    drawPartTrend(oiData, dates);
    drawPartLS(oiData, latest);
    drawLSRatio(oiData, dates);
    drawInstrument(oiData, latest);
    drawTVTrend(tvData, dates);
  }


  /* ════════════════════════════════════════════════════════
     PCR ANALYSIS CHARTS
  ════════════════════════════════════════════════════════ */

  /* ── PCR1: Index OI PCR Trend ── */
  function drawPCRTrend(oiData, dates) {
    if (!dates || dates.length < 1) return;
    const el = document.getElementById('chart-pcr-trend');
    const W  = Math.max(el.parentElement.clientWidth - 32, 860);
    const H  = 230;
    const mg = { top: 16, right: 20, bottom: 32, left: 56 };
    const iw = W - mg.left - mg.right;
    const ih = H - mg.top  - mg.bottom;

    el.setAttribute('viewBox', `0 0 ${W} ${H}`);
    el.setAttribute('height', H);
    const svg = d3.select(el); svg.selectAll('*').remove();
    const g   = svg.append('g').attr('transform', `translate(${mg.left},${mg.top})`);

    const series = PARTS.map(p => ({
      part: p,
      values: dates.map(d => ({
        date: d,
        val: ((oiData[d] || {})[p] || {}).idx_oi_pcr ?? null,
      })).filter(v => v.val != null),
    }));

    const allVals = series.flatMap(s => s.values.map(v => v.val));
    const yMax = Math.max(d3.max(allVals) * 1.15, 2.0);
    const yMin = Math.min(d3.min(allVals) * 0.85, 0.5);

    const xSc = d3.scalePoint().domain(dates).range([0, iw]).padding(0.1);
    const ySc = d3.scaleLinear().domain([yMin, yMax]).range([ih, 0]);

    // Reference bands
    const bands = [
      { y1: 0, y2: 0.8,  col: 'rgba(16,185,129,0.06)',  label: 'Bullish zone' },
      { y1: 1.5, y2: yMax, col: 'rgba(239,68,68,0.06)', label: 'Bearish zone' },
    ];
    bands.forEach(b => {
      if (b.y2 > yMin && b.y1 < yMax) {
        g.append('rect')
          .attr('x', 0).attr('width', iw)
          .attr('y', ySc(Math.min(b.y2, yMax)))
          .attr('height', Math.abs(ySc(Math.max(b.y1, yMin)) - ySc(Math.min(b.y2, yMax))))
          .attr('fill', b.col);
      }
    });

    // Reference lines
    [{ v: 0.8, col:'#10B981', lbl:'0.8' },
     { v: 1.0, col:'#94A3B8', lbl:'1.0' },
     { v: 1.5, col:'#EF4444', lbl:'1.5' }].forEach(ref => {
      if (ref.v >= yMin && ref.v <= yMax) {
        g.append('line')
          .attr('x1',0).attr('x2',iw).attr('y1',ySc(ref.v)).attr('y2',ySc(ref.v))
          .attr('stroke', ref.col).attr('stroke-dasharray','4,3').attr('opacity',0.6);
        g.append('text').attr('x', iw + 4).attr('y', ySc(ref.v) + 3)
          .attr('fill', ref.col).attr('font-size','9px').attr('font-family','var(--mono)')
          .text(ref.lbl);
      }
    });

    g.append('g').attr('class','grid')
     .call(d3.axisLeft(ySc).tickSize(-iw).tickFormat('').ticks(6));
    g.append('g').attr('class','axis').attr('transform',`translate(0,${ih})`)
     .call(d3.axisBottom(xSc).tickSize(3));
    g.append('g').attr('class','axis')
     .call(d3.axisLeft(ySc).ticks(6).tickFormat(v => d3.format('.2f')(v)));

    const line = d3.line()
      .x(d => xSc(d.date)).y(d => ySc(d.val))
      .curve(d3.curveMonotoneX);

    series.forEach(s => {
      if (!s.values.length) return;
      g.append('path').datum(s.values)
        .attr('fill','none').attr('stroke', PART_COLORS[s.part])
        .attr('stroke-width', 2.5).attr('d', line);
      g.selectAll(null).data(s.values).join('circle')
        .attr('cx', d => xSc(d.date)).attr('cy', d => ySc(d.val))
        .attr('r', 4.5).attr('fill', PART_COLORS[s.part])
        .attr('stroke','var(--surface)').attr('stroke-width', 2)
        .on('mouseover', (ev, d) => {
          const sig = d.val < 0.8 ? '🟢 Bullish' : d.val > 1.5 ? '🔴 Bearish' : '🟡 Neutral';
          showTip(`<b style="color:${PART_COLORS[s.part]}">${s.part}</b><br/>
            ${d.date}<br/>Index OI PCR: <b>${d3.format('.3f')(d.val)}</b><br/>${sig}`, ev);
        })
        .on('mousemove', moveTip).on('mouseout', hideTip);
    });

    const lg = document.getElementById('pcr-trend-legend');
    lg.innerHTML = PARTS.map(p =>
      `<div class="legend-item">
         <div class="legend-line" style="background:${PART_COLORS[p]}"></div>${p}
       </div>`).join('');
  }

  /* ── PCR2: PCR Comparison — latest session ── */
  function drawPCRCompare(oiData, latest) {
    const PCR_TYPES = [
      { key: 'idx_oi_pcr',  label: 'Index',    col: 'var(--call)' },
      { key: 'stk_oi_pcr',  label: 'Stock',    col: 'var(--put)'  },
      { key: 'comb_oi_pcr', label: 'Combined', col: 'var(--green)' },
    ];
    const rows = [];
    PARTS.forEach(p => {
      const rec = ((oiData[latest] || {})[p]) || {};
      PCR_TYPES.forEach(t => {
        rows.push({ group: p, type: t.label, val: rec[t.key] ?? 0, col: t.col });
      });
    });

    const el = document.getElementById('chart-pcr-compare');
    const W  = Math.max(el.parentElement.clientWidth - 32, 860);
    const H  = 210;
    const mg = { top: 8, right: 60, bottom: 28, left: 56 };
    const iw = W - mg.left - mg.right;
    const ih = H - mg.top  - mg.bottom;

    el.setAttribute('viewBox', `0 0 ${W} ${H}`);
    el.setAttribute('height', H);
    const svg = d3.select(el); svg.selectAll('*').remove();
    const g   = svg.append('g').attr('transform', `translate(${mg.left},${mg.top})`);

    const yBand   = d3.scaleBand().domain(PARTS).range([0, ih]).padding(0.28);
    const subBand = d3.scaleBand().domain(PCR_TYPES.map(t => t.label))
                      .range([0, yBand.bandwidth()]).padding(0.1);
    const xMax = Math.max(d3.max(rows, r => r.val) * 1.1, 2.0);
    const xSc  = d3.scaleLinear().domain([0, xMax]).range([0, iw]);

    // Reference lines
    [{ v: 0.8, col:'#10B981' }, { v: 1.0, col:'#94A3B8' }, { v: 1.5, col:'#EF4444' }]
      .forEach(ref => {
        g.append('line')
          .attr('x1',xSc(ref.v)).attr('x2',xSc(ref.v)).attr('y1',0).attr('y2',ih)
          .attr('stroke', ref.col).attr('stroke-dasharray','3,3').attr('opacity',0.5);
        g.append('text').attr('x',xSc(ref.v)+2).attr('y',-2)
          .attr('fill',ref.col).attr('font-size','9px').attr('font-family','var(--mono)')
          .text(ref.v.toFixed(1));
      });

    g.append('g').attr('class','axis').attr('transform',`translate(0,${ih})`)
     .call(d3.axisBottom(xSc).ticks(5).tickFormat(v => d3.format('.1f')(v)));
    g.append('g').attr('class','axis').call(d3.axisLeft(yBand).tickSize(0))
     .select('.domain').remove();

    PARTS.forEach(p => {
      const partRows = rows.filter(r => r.group === p);
      partRows.forEach(r => {
        g.append('rect')
          .attr('x', 0)
          .attr('y', yBand(p) + subBand(r.type))
          .attr('width', Math.max(xSc(r.val), 2))
          .attr('height', subBand.bandwidth())
          .attr('fill', r.col).attr('opacity', 0.75).attr('rx', 2)
          .on('mouseover', ev => {
            const sig = r.val < 0.8 ? '🟢 Bullish' : r.val > 1.5 ? '🔴 Bearish' : '🟡 Neutral';
            showTip(`<b>${p}</b> · ${r.type} PCR<br/>
              <b>${d3.format('.3f')(r.val)}</b> ${sig}`, ev);
          })
          .on('mousemove', moveTip).on('mouseout', hideTip);
      });
    });

    // Legend (right side)
    PCR_TYPES.forEach((t, i) => {
      svg.append('rect')
        .attr('x', W - mg.right + 6).attr('y', mg.top + i * 18)
        .attr('width', 10).attr('height', 10).attr('rx', 2)
        .attr('fill', t.col).attr('opacity', 0.75);
      svg.append('text')
        .attr('x', W - mg.right + 20).attr('y', mg.top + i * 18 + 9)
        .attr('fill','var(--muted)').attr('font-size','10px')
        .attr('font-family','var(--mono)').text(t.label);
    });
  }

  /* ── PCR3: Long-side vs Short-side PCR ── */
  function drawPCRSides(oiData, dates) {
    if (!dates || dates.length < 1) return;
    const el = document.getElementById('chart-pcr-sides');
    const W  = Math.max(el.parentElement.clientWidth - 32, 860);
    const H  = 210;
    const mg = { top: 16, right: 20, bottom: 32, left: 56 };
    const iw = W - mg.left - mg.right;
    const ih = H - mg.top  - mg.bottom;

    el.setAttribute('viewBox', `0 0 ${W} ${H}`);
    el.setAttribute('height', H);
    const svg = d3.select(el); svg.selectAll('*').remove();
    const g   = svg.append('g').attr('transform', `translate(${mg.left},${mg.top})`);

    // Two series per participant: idx_oi_pcr (long-side) + short_pcr (short-side)
    const seriesAll = [];
    PARTS.forEach(p => {
      seriesAll.push({
        part: p, side: 'long',
        dash: null,
        values: dates.map(d => ({
          date: d, val: ((oiData[d] || {})[p] || {}).idx_oi_pcr ?? null
        })).filter(v => v.val != null),
      });
      seriesAll.push({
        part: p, side: 'short',
        dash: '5,3',
        values: dates.map(d => ({
          date: d, val: ((oiData[d] || {})[p] || {}).short_pcr ?? null
        })).filter(v => v.val != null),
      });
    });

    const allVals = seriesAll.flatMap(s => s.values.map(v => v.val)).filter(v => v != null);
    const yMax = Math.max(d3.max(allVals) * 1.1, 2.0);
    const yMin = Math.min(d3.min(allVals) * 0.9, 0.5);

    const xSc = d3.scalePoint().domain(dates).range([0, iw]).padding(0.1);
    const ySc = d3.scaleLinear().domain([yMin, yMax]).range([ih, 0]);

    g.append('g').attr('class','grid')
     .call(d3.axisLeft(ySc).tickSize(-iw).tickFormat('').ticks(5));
    g.append('line')
     .attr('x1',0).attr('x2',iw).attr('y1',ySc(1)).attr('y2',ySc(1))
     .attr('stroke','var(--border2)').attr('stroke-dasharray','4,3');
    g.append('g').attr('class','axis').attr('transform',`translate(0,${ih})`)
     .call(d3.axisBottom(xSc).tickSize(3));
    g.append('g').attr('class','axis')
     .call(d3.axisLeft(ySc).ticks(5).tickFormat(v => d3.format('.2f')(v)));

    const line = d3.line()
      .x(d => xSc(d.date)).y(d => ySc(d.val)).curve(d3.curveMonotoneX);

    seriesAll.forEach(s => {
      if (!s.values.length) return;
      g.append('path').datum(s.values)
        .attr('fill','none').attr('stroke', PART_COLORS[s.part])
        .attr('stroke-width', s.side === 'long' ? 2 : 1.5)
        .attr('stroke-dasharray', s.dash || null)
        .attr('opacity', s.side === 'long' ? 1 : 0.55)
        .attr('d', line);
      g.selectAll(null).data(s.values).join('circle')
        .attr('cx', d => xSc(d.date)).attr('cy', d => ySc(d.val))
        .attr('r', s.side === 'long' ? 3.5 : 2.5)
        .attr('fill', PART_COLORS[s.part])
        .attr('opacity', s.side === 'long' ? 1 : 0.55)
        .attr('stroke','var(--surface)').attr('stroke-width', 1.5)
        .on('mouseover', (ev, d) => showTip(
          `<b style="color:${PART_COLORS[s.part]}">${s.part}</b> · ${s.side === 'long' ? 'Long-side' : 'Short-side'}<br/>
           ${d.date}<br/>PCR: <b>${d3.format('.3f')(d.val)}</b>`, ev))
        .on('mousemove', moveTip).on('mouseout', hideTip);
    });

    const lg = document.getElementById('pcr-sides-legend');
    lg.innerHTML = PARTS.map(p =>
      `<div class="legend-item">
         <div class="legend-line" style="background:${PART_COLORS[p]}"></div>${p}
       </div>`).join('') +
      `<div class="legend-item" style="margin-left:8px;padding-left:8px;border-left:1px solid var(--border)">
         <div class="legend-line" style="background:var(--muted)"></div>solid = long-side
       </div>
       <div class="legend-item">
         <div class="legend-line" style="background:var(--muted);opacity:.5;
           border-bottom:2px dashed var(--muted);height:0"></div>dashed = short-side
       </div>`;
  }

  /* ── PCR4: Chain-level aggregate PCR cards ── */
  function drawChainPCR(chain) {
    const strip = document.getElementById('chain-pcr-strip');
    if (!chain || !strip) return;
    const cards = [
      { val: chain.oi_pcr,         lbl: 'OI PCR (Chain)',
        sub: 'Put OI / Call OI · all strikes', cls: 'c-green' },
      { val: chain.vol_pcr,        lbl: 'Volume PCR (Chain)',
        sub: 'Put Vol / Call Vol · all strikes', cls: 'c-amber' },
      { val: chain.total_call_oi,  lbl: 'Total Call OI',
        sub: chain.date || '—', cls: 'c-call', fmt: 'N' },
      { val: chain.total_put_oi,   lbl: 'Total Put OI',
        sub: chain.date || '—', cls: 'c-put', fmt: 'N' },
    ];
    strip.innerHTML = cards.map(c => `
      <div class="metric-card">
        <div class="metric-val ${c.cls}">${c.fmt === 'N' ? fmtN(c.val) : (c.val != null ? d3.format('.3f')(c.val) : '—')}</div>
        <div class="metric-lbl">${c.lbl}</div>
        <div class="metric-sub">${c.sub}</div>
      </div>`).join('');
  }

  /* ── PCR5: Volume PCR Trend (conditional) ── */
  function drawVolPCRTrend(oiData, tvData, dates) {
    const tvDates = dates.filter(d => tvData[d]);
    if (!tvDates.length) return;
    document.getElementById('vol-pcr-section').style.display = 'block';

    const el = document.getElementById('chart-vol-pcr');
    const W  = Math.max(el.parentElement.clientWidth - 32, 860);
    const H  = 200;
    const mg = { top: 16, right: 20, bottom: 32, left: 56 };
    const iw = W - mg.left - mg.right;
    const ih = H - mg.top  - mg.bottom;

    el.setAttribute('viewBox', `0 0 ${W} ${H}`);
    el.setAttribute('height', H);
    const svg = d3.select(el); svg.selectAll('*').remove();
    const g   = svg.append('g').attr('transform', `translate(${mg.left},${mg.top})`);

    // OI PCR (solid reference) + Volume PCR (dashed)
    const seriesAll = [];
    PARTS.forEach(p => {
      seriesAll.push({
        part: p, type: 'OI', dash: null,
        values: dates.map(d => ({
          date: d, val: ((oiData[d] || {})[p] || {}).idx_oi_pcr ?? null
        })).filter(v => v.val != null),
      });
      seriesAll.push({
        part: p, type: 'Vol', dash: '5,3',
        values: tvDates.map(d => ({
          date: d, val: ((tvData[d] || {})[p] || {}).idx_oi_pcr ?? null
        })).filter(v => v.val != null),
      });
    });

    const allVals = seriesAll.flatMap(s => s.values.map(v => v.val)).filter(v => v != null);
    const yMax = Math.max(d3.max(allVals) * 1.1, 2.0);
    const yMin = Math.min(d3.min(allVals) * 0.9, 0.5);

    const xSc = d3.scalePoint().domain(dates).range([0, iw]).padding(0.1);
    const ySc = d3.scaleLinear().domain([yMin, yMax]).range([ih, 0]);

    g.append('g').attr('class','grid')
     .call(d3.axisLeft(ySc).tickSize(-iw).tickFormat('').ticks(5));
    [0.8, 1.0, 1.5].forEach(ref => {
      if (ref >= yMin && ref <= yMax) {
        g.append('line')
          .attr('x1',0).attr('x2',iw).attr('y1',ySc(ref)).attr('y2',ySc(ref))
          .attr('stroke', ref===1.0?'#94A3B8':ref<1?'#10B981':'#EF4444')
          .attr('stroke-dasharray','3,4').attr('opacity',0.45);
      }
    });
    g.append('g').attr('class','axis').attr('transform',`translate(0,${ih})`)
     .call(d3.axisBottom(xSc).tickSize(3));
    g.append('g').attr('class','axis')
     .call(d3.axisLeft(ySc).ticks(5).tickFormat(v => d3.format('.2f')(v)));

    const line = d3.line()
      .x(d => xSc(d.date)).y(d => ySc(d.val)).curve(d3.curveMonotoneX);

    seriesAll.forEach(s => {
      if (!s.values.length) return;
      g.append('path').datum(s.values)
        .attr('fill','none').attr('stroke', PART_COLORS[s.part])
        .attr('stroke-width', s.type === 'OI' ? 1.5 : 2)
        .attr('stroke-dasharray', s.dash || null)
        .attr('opacity', s.type === 'OI' ? 0.4 : 1)
        .attr('d', line);
    });

    const lg = document.getElementById('vol-pcr-legend');
    lg.innerHTML = PARTS.map(p =>
      `<div class="legend-item">
         <div class="legend-line" style="background:${PART_COLORS[p]}"></div>${p}
       </div>`).join('') +
      `<div class="legend-item" style="margin-left:8px;padding-left:8px;border-left:1px solid var(--border)">
         <div class="legend-line" style="background:var(--muted)"></div>solid=OI (ref)
       </div>
       <div class="legend-item">
         <div class="legend-line" style="background:var(--muted)"></div>dashed=Volume
       </div>`;
  }

  /* ── Wire PCR tab ── */
  function renderPCRTab(data) {
    const oiData = (data.oi || {}).data || {};
    const tvData = (data.tv || {}).data || {};
    const dates  = (data.oi || {}).dates || [];
    const latest = dates[dates.length - 1];
    if (!latest) return;

    document.getElementById('pcr-latest-tag').textContent = latest;
    drawPCRTrend(oiData, dates);
    drawPCRCompare(oiData, latest);
    drawPCRSides(oiData, dates);
    drawChainPCR(data.chain || null);
    drawVolPCRTrend(oiData, tvData, dates);
  }


  /* ════════════════════════════════════════════════════════
     NET BIAS & EFFICIENCY CHARTS
  ════════════════════════════════════════════════════════ */

  /* ── B1: OI-TV Efficiency Ratio ── */
  function drawEfficiency(effData, dates) {
    const effDates = dates.filter(d => effData[d]);
    if (!effDates.length) {
      document.getElementById('chart-efficiency').parentElement.querySelector('.chart-sub')
        .textContent = 'TV data not available — efficiency ratios require trading volume input.';
      return;
    }
    const latest = effDates[effDates.length - 1];

    const INSTRS = [
      { key: 'oi_tv_fut_idx',  label: 'Fut Index' },
      { key: 'oi_tv_fut_stk',  label: 'Fut Stock'  },
      { key: 'oi_tv_idx_call', label: 'Idx Call'   },
      { key: 'oi_tv_idx_put',  label: 'Idx Put'    },
      { key: 'oi_tv_total',    label: 'Total'       },
    ];

    const rows = [];
    PARTS.forEach(p => {
      const rec = (effData[latest] || {})[p] || {};
      INSTRS.forEach(ins => {
        rows.push({ part: p, instr: ins.label, val: rec[ins.key] ?? 0 });
      });
    });

    const el = document.getElementById('chart-efficiency');
    const W  = Math.max(el.parentElement.clientWidth - 32, 860);
    const H  = 210;
    const mg = { top: 8, right: 16, bottom: 32, left: 60 };
    const iw = W - mg.left - mg.right;
    const ih = H - mg.top  - mg.bottom;

    el.setAttribute('viewBox', `0 0 ${W} ${H}`);
    el.setAttribute('height', H);
    const svg = d3.select(el); svg.selectAll('*').remove();
    const g   = svg.append('g').attr('transform', `translate(${mg.left},${mg.top})`);

    const instrLabels = INSTRS.map(i => i.label);
    const xBand  = d3.scaleBand().domain(instrLabels).range([0, iw]).padding(0.22);
    const subBand= d3.scaleBand().domain(PARTS).range([0, xBand.bandwidth()]).padding(0.08);
    const yMax   = Math.max(d3.max(rows, r => r.val) * 1.1, 1);
    const ySc    = d3.scaleLinear().domain([0, yMax]).range([ih, 0]);

    g.append('g').attr('class','grid')
     .call(d3.axisLeft(ySc).tickSize(-iw).tickFormat('').ticks(5));
    // Reference line at 1.0
    g.append('line')
     .attr('x1',0).attr('x2',iw).attr('y1',ySc(1)).attr('y2',ySc(1))
     .attr('stroke','var(--amber)').attr('stroke-dasharray','4,3').attr('opacity',0.6);
    g.append('text').attr('x',iw+3).attr('y',ySc(1)+3)
     .attr('fill','var(--amber)').attr('font-size','9px').attr('font-family','var(--mono)').text('1.0');

    g.append('g').attr('class','axis').attr('transform',`translate(0,${ih})`)
     .call(d3.axisBottom(xBand).tickSize(3));
    g.append('g').attr('class','axis')
     .call(d3.axisLeft(ySc).ticks(5).tickFormat(v => d3.format('.1f')(v)));

    instrLabels.forEach(instr => {
      PARTS.forEach(p => {
        const row = rows.find(r => r.part === p && r.instr === instr);
        const val = row ? row.val : 0;
        g.append('rect')
          .attr('x', xBand(instr) + subBand(p))
          .attr('y', ySc(val))
          .attr('width', subBand.bandwidth())
          .attr('height', ih - ySc(val))
          .attr('fill', PART_COLORS[p]).attr('opacity', 0.8).attr('rx', 2)
          .on('mouseover', ev => showTip(
            `<b style="color:${PART_COLORS[p]}">${p}</b> · ${instr}<br/>
             OI/TV Ratio: <b>${d3.format('.2f')(val)}</b><br/>
             ${val > 1 ? '▲ Holding positions' : '▼ High turnover'}`, ev))
          .on('mousemove', moveTip).on('mouseout', hideTip);
      });
    });

    const lg = document.getElementById('eff-legend');
    lg.innerHTML = PARTS.map(p =>
      `<div class="legend-item">
         <div class="legend-dot" style="background:${PART_COLORS[p]}"></div>${p}
       </div>`).join('');
  }

  /* ── B2a: Net Conviction — OI vs Volume Direction ── */
  function drawConviction(oiData, effData, dates, latest) {
    const effDates = dates.filter(d => effData[d]);
    const el = document.getElementById('chart-conviction');
    const W  = Math.max(el.parentElement.clientWidth - 32, 860);
    const H  = 200;
    const mg = { top: 16, right: 16, bottom: 28, left: 56 };
    const iw = W - mg.left - mg.right;
    const ih = H - mg.top  - mg.bottom;

    el.setAttribute('viewBox', `0 0 ${W} ${H}`);
    el.setAttribute('height', H);
    const svg = d3.select(el); svg.selectAll('*').remove();
    const g   = svg.append('g').attr('transform', `translate(${mg.left},${mg.top})`);

    // For each participant: net_oi_fut_idx (OI signal) vs net_tv_fut_idx (Volume signal)
    const hasEff = effDates.includes(latest);
    const effLatest = hasEff ? (effData[latest] || {}) : {};

    const rows = PARTS.map(p => {
      const oiRec  = ((oiData[latest] || {})[p]) || {};
      const effRec = (effLatest[p]) || {};
      return {
        part: p,
        oi_net:  oiRec.net_fut_idx  ?? 0,
        tv_net:  effRec.net_tv_fut_idx ?? null,
      };
    });

    // Grouped bar: OI net (solid) + TV net (dimmed) per participant
    const absMax = Math.max(
      d3.max(rows, r => Math.abs(r.oi_net)),
      d3.max(rows, r => Math.abs(r.tv_net || 0)),
      1
    );
    const yBand  = d3.scaleBand().domain(PARTS).range([0, ih]).padding(0.28);
    const xSc    = d3.scaleLinear().domain([-absMax*1.1, absMax*1.1]).range([0, iw]);
    const xMid   = xSc(0);

    g.append('line')
     .attr('x1',xMid).attr('x2',xMid).attr('y1',0).attr('y2',ih)
     .attr('stroke','var(--border2)').attr('stroke-dasharray','4,3');
    g.append('g').attr('class','axis').attr('transform',`translate(0,${ih})`)
     .call(d3.axisBottom(xSc).ticks(5).tickFormat(fmtAxis));
    g.append('g').attr('class','axis').call(d3.axisLeft(yBand).tickSize(0))
     .select('.domain').remove();

    const barH = yBand.bandwidth() / 2 - 1;
    rows.forEach(r => {
      const y0 = yBand(r.part);
      // OI bar
      const oiW = Math.abs(xSc(r.oi_net) - xMid);
      g.append('rect')
        .attr('x', r.oi_net >= 0 ? xMid : xMid - oiW)
        .attr('y', y0).attr('width', Math.max(oiW,1)).attr('height', barH)
        .attr('fill', r.oi_net >= 0 ? '#10B981' : '#EF4444').attr('opacity', 0.85).attr('rx',2)
        .on('mouseover', ev => showTip(
          `<b style="color:${PART_COLORS[r.part]}">${r.part}</b><br/>
           Net Fut Idx OI: <b>${fmtN(r.oi_net)}</b>`, ev))
        .on('mousemove', moveTip).on('mouseout', hideTip);
      // TV bar (if available)
      if (r.tv_net != null) {
        const tvW = Math.abs(xSc(r.tv_net) - xMid);
        g.append('rect')
          .attr('x', r.tv_net >= 0 ? xMid : xMid - tvW)
          .attr('y', y0 + barH + 2).attr('width', Math.max(tvW,1)).attr('height', barH)
          .attr('fill', r.tv_net >= 0 ? '#10B981' : '#EF4444').attr('opacity', 0.4).attr('rx',2)
          .on('mouseover', ev => showTip(
            `<b style="color:${PART_COLORS[r.part]}">${r.part}</b><br/>
             Net Fut Idx Vol: <b>${fmtN(r.tv_net)}</b>`, ev))
          .on('mousemove', moveTip).on('mouseout', hideTip);
      }
    });

    // Labels
    g.append('text').attr('x', xMid/2).attr('y', -4)
     .attr('fill','var(--faint)').attr('font-size','9px')
     .attr('text-anchor','middle').attr('font-family','var(--mono)').text('← Net Short');
    g.append('text').attr('x', xMid + (iw-xMid)/2).attr('y', -4)
     .attr('fill','var(--faint)').attr('font-size','9px')
     .attr('text-anchor','middle').attr('font-family','var(--mono)').text('Net Long →');
  }

  /* ── B2b: Net OI Composition ── */
  function drawComposition(oiData, latest) {
    const el = document.getElementById('chart-composition');
    const W  = Math.max(el.parentElement.clientWidth - 32, 860);
    const H  = 200;
    const mg = { top: 8, right: 80, bottom: 28, left: 56 };
    const iw = W - mg.left - mg.right;
    const ih = H - mg.top  - mg.bottom;

    el.setAttribute('viewBox', `0 0 ${W} ${H}`);
    el.setAttribute('height', H);
    const svg = d3.select(el); svg.selectAll('*').remove();
    const g   = svg.append('g').attr('transform', `translate(${mg.left},${mg.top})`);

    const SEGS = [
      { key: 'net_futures', label: 'Futures', col: 'var(--call)' },
      { key: 'net_options', label: 'Options', col: 'var(--put)'  },
    ];

    const rows = PARTS.map(p => {
      const rec = ((oiData[latest] || {})[p]) || {};
      return { part: p, net_futures: rec.net_futures||0, net_options: rec.net_options||0 };
    });

    const absMax = Math.max(
      d3.max(rows, r => Math.abs(r.net_futures) + Math.abs(r.net_options)), 1
    );
    const yBand = d3.scaleBand().domain(PARTS).range([0, ih]).padding(0.28);
    const xSc   = d3.scaleLinear().domain([-absMax*1.1, absMax*1.1]).range([0, iw]);
    const xMid  = xSc(0);

    g.append('line')
     .attr('x1',xMid).attr('x2',xMid).attr('y1',0).attr('y2',ih)
     .attr('stroke','var(--border2)').attr('stroke-dasharray','4,3');
    g.append('g').attr('class','axis').attr('transform',`translate(0,${ih})`)
     .call(d3.axisBottom(xSc).ticks(4).tickFormat(fmtAxis));
    g.append('g').attr('class','axis').call(d3.axisLeft(yBand).tickSize(0))
     .select('.domain').remove();

    rows.forEach(r => {
      let xCursor = xMid;
      SEGS.forEach(seg => {
        const val = r[seg.key] || 0;
        const w   = Math.abs(xSc(val) - xSc(0));
        const x   = val >= 0 ? xCursor : xCursor - w;
        g.append('rect')
          .attr('x', x).attr('y', yBand(r.part))
          .attr('width', Math.max(w,1)).attr('height', yBand.bandwidth())
          .attr('fill', seg.col).attr('opacity', 0.8).attr('rx', 2)
          .on('mouseover', ev => showTip(
            `<b style="color:${PART_COLORS[r.part]}">${r.part}</b> · ${seg.label}<br/>
             Net OI: <b>${fmtN(val)}</b>`, ev))
          .on('mousemove', moveTip).on('mouseout', hideTip);
        xCursor = val >= 0 ? xCursor + w : xCursor - w;
      });
    });

    // Legend
    SEGS.forEach((seg, i) => {
      svg.append('rect')
        .attr('x', W - mg.right + 6).attr('y', mg.top + i*18)
        .attr('width',10).attr('height',10).attr('rx',2)
        .attr('fill', seg.col).attr('opacity', 0.8);
      svg.append('text')
        .attr('x', W - mg.right + 20).attr('y', mg.top + i*18 + 9)
        .attr('fill','var(--muted)').attr('font-size','10px')
        .attr('font-family','var(--mono)').text(seg.label);
    });
  }

  /* ── B3: DoD Net OI Change — By Instrument ── */
  function drawDoDInstrument(dodData) {
    const pairs = (dodData.pairs || []);
    if (!pairs.length) return;
    const [d0, d1] = pairs[pairs.length - 1];
    const pairKey  = `${d0}|${d1}`;
    const pairData = (dodData.data || {})[pairKey] || {};

    document.getElementById('dod-instr-tag').textContent = `${d0} → ${d1}`;

    const INSTRS = [
      { key: 'd_net_fut_idx',  label: 'Fut Index'  },
      { key: 'd_net_fut_stk',  label: 'Fut Stock'  },
      { key: 'd_net_idx_call', label: 'Idx Call'   },
      { key: 'd_net_idx_put',  label: 'Idx Put'    },
      { key: 'd_net_stk_call', label: 'Stk Call'   },
      { key: 'd_net_stk_put',  label: 'Stk Put'    },
    ];

    const rows = [];
    PARTS.forEach(p => {
      const rec = pairData[p] || {};
      INSTRS.forEach(ins => {
        rows.push({ part: p, instr: ins.label, val: rec[ins.key] ?? 0 });
      });
    });

    const el = document.getElementById('chart-dod-instr');
    const W  = Math.max(el.parentElement.clientWidth - 32, 860);
    const H  = 230;
    const mg = { top: 8, right: 16, bottom: 40, left: 60 };
    const iw = W - mg.left - mg.right;
    const ih = H - mg.top  - mg.bottom;

    el.setAttribute('viewBox', `0 0 ${W} ${H}`);
    el.setAttribute('height', H);
    const svg = d3.select(el); svg.selectAll('*').remove();
    const g   = svg.append('g').attr('transform', `translate(${mg.left},${mg.top})`);

    const instrLabels = INSTRS.map(i => i.label);
    const xBand  = d3.scaleBand().domain(instrLabels).range([0, iw]).padding(0.22);
    const subBand= d3.scaleBand().domain(PARTS).range([0, xBand.bandwidth()]).padding(0.08);
    const absMax = Math.max(d3.max(rows, r => Math.abs(r.val)), 1);
    const ySc    = d3.scaleLinear().domain([-absMax*1.15, absMax*1.15]).range([ih, 0]);

    g.append('g').attr('class','grid')
     .call(d3.axisLeft(ySc).tickSize(-iw).tickFormat('').ticks(5));
    g.append('line')
     .attr('x1',0).attr('x2',iw).attr('y1',ySc(0)).attr('y2',ySc(0))
     .attr('stroke','var(--border2)').attr('stroke-dasharray','4,3');
    g.append('g').attr('class','axis').attr('transform',`translate(0,${ih})`)
     .call(d3.axisBottom(xBand).tickSize(3));
    g.append('g').attr('class','axis')
     .call(d3.axisLeft(ySc).ticks(5).tickFormat(fmtAxis));

    instrLabels.forEach(instr => {
      PARTS.forEach(p => {
        const row = rows.find(r => r.part === p && r.instr === instr);
        const val = row ? row.val : 0;
        const pos = val >= 0;
        const barH = Math.abs(ySc(val) - ySc(0));
        g.append('rect')
          .attr('x', xBand(instr) + subBand(p))
          .attr('y', pos ? ySc(val) : ySc(0))
          .attr('width', subBand.bandwidth())
          .attr('height', Math.max(barH, 1))
          .attr('fill', PART_COLORS[p])
          .attr('opacity', pos ? 0.85 : 0.45)
          .attr('rx', 2)
          .on('mouseover', ev => showTip(
            `<b style="color:${PART_COLORS[p]}">${p}</b> · ${instr}<br/>
             Δ Net OI: <b>${fmtN(val)}</b><br/>
             ${pos ? '▲ OI added net long' : '▼ OI added net short'}`, ev))
          .on('mousemove', moveTip).on('mouseout', hideTip);
      });
    });

    // Participant legend below x axis
    const lgY = H - mg.bottom + 18;
    PARTS.forEach((p, i) => {
      const lx = mg.left + i * 100;
      svg.append('rect').attr('x',lx).attr('y',lgY)
        .attr('width',10).attr('height',10).attr('rx',2)
        .attr('fill',PART_COLORS[p]).attr('opacity',0.85);
      svg.append('text').attr('x',lx+14).attr('y',lgY+9)
        .attr('fill','var(--muted)').attr('font-size','11px')
        .attr('font-family','var(--mono)').text(p);
    });
  }

  /* ── Wire Bias tab ── */
  function renderBiasTab(data) {
    const oiData  = (data.oi  || {}).data || {};
    const effData = (data.efficiency || {}).data || {};
    const dates   = (data.oi  || {}).dates || [];
    const latest  = dates[dates.length - 1];
    if (!latest) return;

    document.getElementById('bias-latest-tag').textContent = latest;
    drawEfficiency(effData, dates);
    drawConviction(oiData, effData, dates, latest);
    drawComposition(oiData, latest);
    drawDoDInstrument(data.dod || {});
  }


  /* ════════════════════════════════════════════════════════
     OPTION CHAIN TAB
  ════════════════════════════════════════════════════════ */

  let _chainShowAll = false;

  function renderChainTab(data) {
    const chain = data.chain;
    if (!chain || !chain.data || !chain.data.length) {
      document.getElementById('chain-strip').innerHTML =
        '<div style="color:var(--muted);padding:16px">No option chain data available.</div>';
      return;
    }

    /* ── Metrics strip ── */
    const callWall = (chain.call_walls || [])[0] || {};
    const putWall  = (chain.put_walls  || [])[0] || {};
    const cards = [
      { val: fmtN(chain.atm),            lbl: 'ATM Strike',      cls: 'c-call' },
      { val: fmtN(chain.max_pain),        lbl: 'Max Pain',        cls: 'c-amber' },
      { val: fmtF(chain.oi_pcr),          lbl: 'OI PCR',          cls: 'c-green' },
      { val: fmtN(callWall.strike),       lbl: 'Call Wall',       cls: 'c-call' },
      { val: fmtN(putWall.strike),        lbl: 'Put Wall',        cls: 'c-put' },
      { val: fmtN(chain.total_call_oi),   lbl: 'Total Call OI',   cls: 'c-muted' },
      { val: fmtN(chain.total_put_oi),    lbl: 'Total Put OI',    cls: 'c-muted' },
    ];
    document.getElementById('chain-strip').innerHTML = cards.map(c =>
      `<div class="metric-card">
         <div class="metric-val ${c.cls}">${c.val}</div>
         <div class="metric-lbl">${c.lbl}</div>
       </div>`).join('');

    /* ── Sorted strikes ── */
    const rows   = [...chain.data].sort((a, b) => b.strike - a.strike); // desc
    const atm    = chain.atm;
    const maxPain= chain.max_pain;
    const allStrikes = rows.map(r => r.strike);
    const atmIdx = allStrikes.indexOf(atm);

    function getVisibleRows() {
      if (_chainShowAll) return rows;
      const lo = Math.max(0, atmIdx - 20);
      const hi = Math.min(rows.length - 1, atmIdx + 20);
      return rows.slice(lo, hi + 1);
    }

    /* ── OI Profile Chart ── */
    function drawOIProfile(visRows) {
      const el = document.getElementById('chart-oi-profile');
      const W  = Math.max(el.parentElement.clientWidth - 32, 860);
      const H  = Math.min(visRows.length * 18 + 50, 640);
      const mg = { top: 20, right: 20, bottom: 20, left: 20 };
      const iw = W - mg.left - mg.right;
      const ih = H - mg.top  - mg.bottom;

      el.setAttribute('viewBox', `0 0 ${W} ${H}`);
      el.setAttribute('height', H);
      const svg = d3.select(el); svg.selectAll('*').remove();

      const strikeLabels = visRows.map(r => String(r.strike));
      const STRIKE_W = 80;
      const halfW    = (iw - STRIKE_W) / 2;
      const callX0   = mg.left;
      const strikeX0 = mg.left + halfW;
      const putX0    = mg.left + halfW + STRIKE_W;

      const maxOI = d3.max(visRows, r => Math.max(r.c_oi || 0, r.p_oi || 0));
      const callSc = d3.scaleLinear().domain([0, maxOI]).range([0, halfW]);
      const putSc  = d3.scaleLinear().domain([0, maxOI]).range([0, halfW]);
      const ySc    = d3.scaleBand().domain(strikeLabels).range([mg.top, mg.top + ih]).padding(0.15);

      const g = svg.append('g');

      // Column headers
      svg.append('text').attr('x', callX0 + halfW/2).attr('y', 12)
        .attr('text-anchor','middle').attr('fill','var(--call)')
        .attr('font-size','10px').attr('font-family','var(--mono)')
        .attr('font-weight','600').attr('letter-spacing','.05em').text('CALL OI ←');
      svg.append('text').attr('x', putX0 + halfW/2).attr('y', 12)
        .attr('text-anchor','middle').attr('fill','var(--put)')
        .attr('font-size','10px').attr('font-family','var(--mono)')
        .attr('font-weight','600').attr('letter-spacing','.05em').text('→ PUT OI');

      // Rows
      visRows.forEach(r => {
        const ys    = String(r.strike);
        const y     = ySc(ys);
        const bh    = ySc.bandwidth();
        const isATM = r.strike === atm;
        const isMP  = r.strike === maxPain;
        const rowBg = isATM ? 'rgba(245,158,11,0.10)' : 'transparent';

        // Row bg
        g.append('rect').attr('x', callX0).attr('y', y)
          .attr('width', iw).attr('height', bh).attr('fill', rowBg);

        // Call OI bar (left, extends from strike leftward)
        const cw = callSc(r.c_oi || 0);
        g.append('rect')
          .attr('x', strikeX0 - cw).attr('y', y + 1)
          .attr('width', cw).attr('height', bh - 2)
          .attr('fill','var(--call)').attr('opacity', isATM ? 1 : 0.7).attr('rx', 2)
          .on('mouseover', ev => showTip(
            `<b style="color:var(--call)">${r.strike} Call</b><br/>
             OI: <b>${fmtN(r.c_oi)}</b><br/>
             ΔOI: ${r.c_chng_oi >= 0 ? '+' : ''}${fmtN(r.c_chng_oi)}<br/>
             Vol: ${fmtN(r.c_volume)}<br/>IV: ${r.c_iv ?? '—'}`, ev))
          .on('mousemove', moveTip).on('mouseout', hideTip);

        // Put OI bar (right)
        const pw = putSc(r.p_oi || 0);
        g.append('rect')
          .attr('x', putX0).attr('y', y + 1)
          .attr('width', pw).attr('height', bh - 2)
          .attr('fill','var(--put)').attr('opacity', isATM ? 1 : 0.7).attr('rx', 2)
          .on('mouseover', ev => showTip(
            `<b style="color:var(--put)">${r.strike} Put</b><br/>
             OI: <b>${fmtN(r.p_oi)}</b><br/>
             ΔOI: ${r.p_chng_oi >= 0 ? '+' : ''}${fmtN(r.p_chng_oi)}<br/>
             Vol: ${fmtN(r.p_volume)}<br/>IV: ${r.p_iv ?? '—'}`, ev))
          .on('mousemove', moveTip).on('mouseout', hideTip);

        // Strike label
        g.append('text')
          .attr('x', strikeX0 + STRIKE_W/2).attr('y', y + bh/2 + 4)
          .attr('text-anchor','middle')
          .attr('fill', isATM ? 'var(--amber)' : isMP ? 'var(--green)' : 'var(--muted)')
          .attr('font-size', isATM ? '11px' : '10px')
          .attr('font-weight', isATM || isMP ? '700' : '400')
          .attr('font-family','var(--mono)')
          .text(isATM ? `★ ${r.strike}` : isMP ? `◆ ${r.strike}` : r.strike);

        // PCR label
        if (r.pcr != null) {
          g.append('text')
            .attr('x', putX0 + Math.min(pw + 6, halfW - 30)).attr('y', y + bh/2 + 4)
            .attr('fill', r.pcr > 1.5 ? 'var(--put)' : r.pcr < 0.8 ? 'var(--call)' : 'var(--faint)')
            .attr('font-size','9px').attr('font-family','var(--mono)')
            .text(d3.format('.2f')(r.pcr));
        }
      });

      // Max Pain line
      if (allStrikes.includes(maxPain)) {
        const mpY = ySc(String(maxPain)) + ySc.bandwidth()/2;
        g.append('line')
          .attr('x1', callX0).attr('x2', callX0 + iw)
          .attr('y1', mpY).attr('y2', mpY)
          .attr('stroke','var(--green)').attr('stroke-dasharray','6,3')
          .attr('stroke-width', 1).attr('opacity', 0.6);
        g.append('text').attr('x', callX0 + 4).attr('y', mpY - 3)
          .attr('fill','var(--green)').attr('font-size','9px')
          .attr('font-family','var(--mono)').text('Max Pain');
      }
    }

    /* ── Strike Table ── */
    function buildTable(visRows) {
      const thead = document.getElementById('chain-thead');
      const tbody = document.getElementById('chain-tbody');

      // Header — two super-rows
      thead.innerHTML = `
        <tr>
          <th colspan="5" class="call-hdr" style="text-align:center;border-bottom:none">← CALLS</th>
          <th class="strike-col" rowspan="2" style="vertical-align:middle">Strike</th>
          <th colspan="5" class="put-hdr" style="text-align:center;border-bottom:none">PUTS →</th>
          <th class="put-hdr" rowspan="2" style="vertical-align:middle">PCR</th>
        </tr>
        <tr>
          <th class="call-hdr">OI</th>
          <th class="call-hdr">ΔOI</th>
          <th class="call-hdr">Vol</th>
          <th class="call-hdr">IV</th>
          <th class="call-hdr">LTP</th>
          <th class="put-hdr">LTP</th>
          <th class="put-hdr">IV</th>
          <th class="put-hdr">Vol</th>
          <th class="put-hdr">ΔOI</th>
          <th class="put-hdr">OI</th>
        </tr>`;

      const fmtDelta = (v) => {
        if (v == null || isNaN(v)) return '<span style="color:var(--faint)">—</span>';
        const cls = v > 0 ? 'delta-pos' : v < 0 ? 'delta-neg' : '';
        const str = (v > 0 ? '+' : '') + fmtN(v);
        return cls ? `<span class="${cls}">${str}</span>` : str;
      };
      const fmtPCR = (v) => {
        if (v == null) return '—';
        const cls = v > 1.5 ? 'pcr-high' : v < 0.8 ? 'pcr-low' : '';
        return cls ? `<span class="${cls}">${d3.format('.2f')(v)}</span>` : d3.format('.2f')(v);
      };
      const fmtIV = v => v != null ? d3.format('.1f')(v) : '—';
      const fmtLTP = v => v != null ? d3.format('.2f')(v) : '—';

      tbody.innerHTML = visRows.map(r => {
        const isATM = r.strike === atm;
        const isMP  = r.strike === maxPain;
        const cls   = isATM ? ' class="atm-row"' : isMP ? ' class="maxpain-row"' : '';
        return `<tr${cls}>
          <td style="color:var(--call)">${fmtN(r.c_oi)}</td>
          <td>${fmtDelta(r.c_chng_oi)}</td>
          <td style="color:var(--muted)">${fmtN(r.c_volume)}</td>
          <td style="color:rgba(59,130,246,0.7)">${fmtIV(r.c_iv)}</td>
          <td style="color:var(--text)">${fmtLTP(r.c_ltp)}</td>
          <td class="strike-col">${isATM ? '★ ' : isMP ? '◆ ' : ''}${fmtN(r.strike)}</td>
          <td style="color:var(--text)">${fmtLTP(r.p_ltp)}</td>
          <td style="color:rgba(249,115,22,0.7)">${fmtIV(r.p_iv)}</td>
          <td style="color:var(--muted)">${fmtN(r.p_volume)}</td>
          <td>${fmtDelta(r.p_chng_oi)}</td>
          <td style="color:var(--put)">${fmtN(r.p_oi)}</td>
          <td>${fmtPCR(r.pcr)}</td>
        </tr>`;
      }).join('');
    }

    /* ── Initial render ── */
    const visRows = getVisibleRows();
    drawOIProfile(visRows);
    buildTable(visRows);

    /* ── Expand/collapse toggle ── */
    const btn = document.getElementById('chain-expand-btn');
    btn.onclick = () => {
      _chainShowAll = !_chainShowAll;
      btn.textContent = _chainShowAll ? 'Show ATM ±20' : 'Show all strikes';
      const vr = getVisibleRows();
      drawOIProfile(vr);
      buildTable(vr);
    };
  }


  /* ════════════════════════════════════════════════════════
     MAX PAIN TAB
  ════════════════════════════════════════════════════════ */

  function renderMaxPainTab(data) {
    const chain = data.chain;
    if (!chain || !chain.data || !chain.data.length) return;

    const rows    = [...chain.data].sort((a, b) => a.strike - b.strike);
    const atm     = chain.atm;
    const maxPain = chain.max_pain;
    const strikes = rows.map(r => r.strike);

    /* ── Recompute call/put pain breakdown in JS ── */
    const painRows = rows.map(target => {
      let callPain = 0, putPain = 0;
      rows.forEach(k => {
        callPain += Math.max(0, target.strike - k.strike) * (k.c_oi || 0);
        putPain  += Math.max(0, k.strike - target.strike) * (k.p_oi || 0);
      });
      return { strike: target.strike, callPain, putPain, total: callPain + putPain };
    });

    const mpRow   = painRows.find(r => r.strike === maxPain) || painRows[0];
    const atmRow  = painRows.find(r => r.strike === atm);
    const mpIdx   = strikes.indexOf(maxPain);
    const atmDist = maxPain - atm;

    /* ── Metrics strip ── */
    document.getElementById('mp-strip').innerHTML = [
      { val: fmtN(maxPain),            lbl: 'Max Pain Strike',   cls: 'c-amber' },
      { val: fmtN(atm),                lbl: 'ATM Strike',        cls: 'c-call'  },
      { val: (atmDist >= 0 ? '+' : '') + fmtN(atmDist), lbl: 'Distance ATM→MaxPain', cls: atmDist >= 0 ? 'c-green' : 'c-put' },
      { val: fmtN(mpRow.total),        lbl: 'Pain at Max Pain',  cls: 'c-muted' },
      { val: atmRow ? fmtN(atmRow.total) : '—', lbl: 'Pain at ATM', cls: 'c-muted' },
      { val: fmtF(chain.oi_pcr),       lbl: 'OI PCR',            cls: 'c-green' },
    ].map(c => `<div class="metric-card">
      <div class="metric-val ${c.cls}">${c.val}</div>
      <div class="metric-lbl">${c.lbl}</div>
    </div>`).join('');

    /* ── Pain Curve ── */
    (function() {
      const el = document.getElementById('chart-pain-curve');
      const W  = Math.max(el.parentElement.clientWidth - 32, 860);
      const H  = 240;
      const mg = { top: 16, right: 20, bottom: 40, left: 72 };
      const iw = W - mg.left - mg.right;
      const ih = H - mg.top  - mg.bottom;

      el.setAttribute('viewBox', `0 0 ${W} ${H}`);
      el.setAttribute('height', H);
      const svg = d3.select(el); svg.selectAll('*').remove();
      const g   = svg.append('g').attr('transform', `translate(${mg.left},${mg.top})`);

      const xSc = d3.scaleLinear()
        .domain([d3.min(strikes), d3.max(strikes)])
        .range([0, iw]);
      const yMax = d3.max(painRows, r => r.total);
      const ySc  = d3.scaleLinear().domain([0, yMax * 1.05]).range([ih, 0]);

      // Grid
      g.append('g').attr('class','grid')
       .call(d3.axisLeft(ySc).tickSize(-iw).tickFormat('').ticks(5));

      // Area fill
      const area = d3.area()
        .x(d => xSc(d.strike))
        .y0(ih).y1(d => ySc(d.total))
        .curve(d3.curveCatmullRom);
      g.append('path').datum(painRows)
        .attr('fill','rgba(99,102,241,0.12)')
        .attr('d', area);

      // Line
      const line = d3.line()
        .x(d => xSc(d.strike))
        .y(d => ySc(d.total))
        .curve(d3.curveCatmullRom);
      g.append('path').datum(painRows)
        .attr('fill','none').attr('stroke','#6366F1')
        .attr('stroke-width', 2).attr('d', line);

      // ATM vertical
      g.append('line')
        .attr('x1', xSc(atm)).attr('x2', xSc(atm))
        .attr('y1', 0).attr('y2', ih)
        .attr('stroke','var(--amber)').attr('stroke-dasharray','5,3')
        .attr('stroke-width', 1.5).attr('opacity', 0.8);
      g.append('text').attr('x', xSc(atm) + 4).attr('y', 10)
        .attr('fill','var(--amber)').attr('font-size','9px')
        .attr('font-family','var(--mono)').text('ATM');

      // Max Pain marker
      g.append('line')
        .attr('x1', xSc(maxPain)).attr('x2', xSc(maxPain))
        .attr('y1', 0).attr('y2', ih)
        .attr('stroke','var(--green)').attr('stroke-width', 2).attr('opacity', 0.9);
      g.append('text').attr('x', xSc(maxPain) + 4).attr('y', 22)
        .attr('fill','var(--green)').attr('font-size','9px')
        .attr('font-family','var(--mono)').attr('font-weight','600')
        .text(`Max Pain: ${fmtN(maxPain)}`);

      // Dot at max pain minimum
      g.append('circle')
        .attr('cx', xSc(mpRow.strike)).attr('cy', ySc(mpRow.total))
        .attr('r', 5).attr('fill','var(--green)')
        .attr('stroke','var(--surface)').attr('stroke-width', 2);

      // Axes
      g.append('g').attr('class','axis')
       .attr('transform',`translate(0,${ih})`)
       .call(d3.axisBottom(xSc).ticks(10).tickFormat(v => fmtN(v)));
      g.append('g').attr('class','axis')
       .call(d3.axisLeft(ySc).ticks(5)
         .tickFormat(v => v >= 1e9 ? d3.format('.1f')(v/1e9)+'B'
                        : v >= 1e6 ? d3.format('.0f')(v/1e6)+'M'
                        : v >= 1e3 ? d3.format('.0f')(v/1e3)+'K' : v));

      // Hover interaction
      const bisect = d3.bisector(d => d.strike).left;
      const overlay = g.append('rect')
        .attr('width', iw).attr('height', ih)
        .attr('fill','transparent').style('cursor','crosshair');
      const vline = g.append('line').attr('y1',0).attr('y2',ih)
        .attr('stroke','var(--border2)').attr('stroke-dasharray','3,3')
        .attr('opacity', 0);
      overlay.on('mousemove', function(ev) {
        const [mx] = d3.pointer(ev, this);
        const s    = xSc.invert(mx);
        const idx  = Math.min(bisect(painRows, s), painRows.length - 1);
        const d    = painRows[idx];
        vline.attr('x1', xSc(d.strike)).attr('x2', xSc(d.strike)).attr('opacity',1);
        showTip(`Strike: <b>${fmtN(d.strike)}</b><br/>
          Total Pain: <b>${fmtN(d.total)}</b><br/>
          Call Pain: ${fmtN(d.callPain)}<br/>
          Put Pain: ${fmtN(d.putPain)}${d.strike === maxPain ? '<br/><b style="color:var(--green)">★ Max Pain</b>' : ''}`, ev);
      }).on('mouseout', () => { vline.attr('opacity',0); hideTip(); });
    })();

    /* ── Call vs Put Pain Breakdown ── */
    (function() {
      const n    = 15;
      const lo   = Math.max(0, mpIdx - n);
      const hi   = Math.min(painRows.length - 1, mpIdx + n);
      const sub  = painRows.slice(lo, hi + 1);

      const el = document.getElementById('chart-pain-breakdown');
      const W  = Math.max(el.parentElement.clientWidth - 32, 400);
      const H  = 230;
      const mg = { top: 8, right: 8, bottom: 60, left: 60 };
      const iw = W - mg.left - mg.right;
      const ih = H - mg.top  - mg.bottom;

      el.setAttribute('viewBox', `0 0 ${W} ${H}`);
      el.setAttribute('height', H);
      const svg = d3.select(el); svg.selectAll('*').remove();
      const g   = svg.append('g').attr('transform', `translate(${mg.left},${mg.top})`);

      const xBand = d3.scaleBand().domain(sub.map(r => String(r.strike)))
                      .range([0, iw]).padding(0.1);
      const yMax  = d3.max(sub, r => r.total);
      const ySc   = d3.scaleLinear().domain([0, yMax * 1.05]).range([ih, 0]);

      g.append('g').attr('class','grid')
       .call(d3.axisLeft(ySc).tickSize(-iw).tickFormat('').ticks(4));

      // Stacked bars: callPain (bottom, blue), putPain (top, orange)
      sub.forEach(r => {
        const x    = xBand(String(r.strike));
        const bw   = xBand.bandwidth();
        const isMP = r.strike === maxPain;
        const isATM= r.strike === atm;

        // Call pain (bottom)
        const cy = ySc(r.callPain);
        g.append('rect').attr('x',x).attr('y',cy)
          .attr('width',bw).attr('height', ih - cy)
          .attr('fill','var(--call)').attr('opacity', isMP ? 1 : 0.6).attr('rx',2)
          .on('mouseover', ev => showTip(
            `Strike: <b>${fmtN(r.strike)}</b><br/>
             Call Pain: ${fmtN(r.callPain)}<br/>
             Put Pain: ${fmtN(r.putPain)}<br/>
             Total: ${fmtN(r.total)}`, ev))
          .on('mousemove', moveTip).on('mouseout', hideTip);

        // Put pain (stacked on top)
        const py = ySc(r.total);
        g.append('rect').attr('x',x).attr('y',py)
          .attr('width',bw).attr('height', cy - py)
          .attr('fill','var(--put)').attr('opacity', isMP ? 1 : 0.6).attr('rx',2)
          .on('mouseover', ev => showTip(
            `Strike: <b>${fmtN(r.strike)}</b><br/>
             Call Pain: ${fmtN(r.callPain)}<br/>
             Put Pain: ${fmtN(r.putPain)}<br/>
             Total: ${fmtN(r.total)}`, ev))
          .on('mousemove', moveTip).on('mouseout', hideTip);

        // Marker for max pain and ATM
        if (isMP) {
          g.append('text').attr('x', x + bw/2).attr('y', py - 4)
            .attr('text-anchor','middle').attr('fill','var(--green)')
            .attr('font-size','9px').attr('font-family','var(--mono)').text('★');
        }
        if (isATM) {
          g.append('text').attr('x', x + bw/2).attr('y', ih + 26)
            .attr('text-anchor','middle').attr('fill','var(--amber)')
            .attr('font-size','8px').attr('font-family','var(--mono)').text('ATM');
        }
      });

      g.append('g').attr('class','axis').attr('transform',`translate(0,${ih})`)
       .call(d3.axisBottom(xBand).tickSize(3))
       .selectAll('text')
       .attr('transform','rotate(-45)').attr('text-anchor','end')
       .attr('dx','-4px').attr('dy','4px');
      g.append('g').attr('class','axis')
       .call(d3.axisLeft(ySc).ticks(4)
         .tickFormat(v => v >= 1e9 ? d3.format('.1f')(v/1e9)+'B'
                        : v >= 1e6 ? d3.format('.0f')(v/1e6)+'M' : v));

      // Legend
      ['Call Pain','Put Pain'].forEach((lbl, i) => {
        const col = i === 0 ? 'var(--call)' : 'var(--put)';
        const lx  = iw - 110 + i * 58;
        g.append('rect').attr('x',lx).attr('y',-10)
         .attr('width',10).attr('height',8).attr('rx',2).attr('fill',col).attr('opacity',.7);
        g.append('text').attr('x',lx+13).attr('y',-3)
         .attr('fill','var(--muted)').attr('font-size','9px')
         .attr('font-family','var(--mono)').text(lbl);
      });
    })();

    /* ── OI at Max Pain Region ── */
    (function() {
      const n   = 5;
      const lo  = Math.max(0, mpIdx - n);
      const hi  = Math.min(rows.length - 1, mpIdx + n);
      const sub = rows.slice(lo, hi + 1);

      const el = document.getElementById('chart-mp-oi');
      const W  = Math.max(el.parentElement.clientWidth - 32, 400);
      const H  = 230;
      const mg = { top: 8, right: 8, bottom: 60, left: 64 };
      const iw = W - mg.left - mg.right;
      const ih = H - mg.top  - mg.bottom;

      el.setAttribute('viewBox', `0 0 ${W} ${H}`);
      el.setAttribute('height', H);
      const svg = d3.select(el); svg.selectAll('*').remove();
      const g   = svg.append('g').attr('transform', `translate(${mg.left},${mg.top})`);

      const labels  = sub.map(r => String(r.strike));
      const xBand   = d3.scaleBand().domain(labels).range([0, iw]).padding(0.2);
      const subBand = d3.scaleBand().domain(['Call','Put'])
                        .range([0, xBand.bandwidth()]).padding(0.08);
      const yMax = d3.max(sub, r => Math.max(r.c_oi||0, r.p_oi||0)) * 1.1;
      const ySc  = d3.scaleLinear().domain([0, yMax]).range([ih, 0]);

      g.append('g').attr('class','grid')
       .call(d3.axisLeft(ySc).tickSize(-iw).tickFormat('').ticks(4));

      sub.forEach(r => {
        const x    = xBand(String(r.strike));
        const isMP = r.strike === maxPain;
        const isATM= r.strike === atm;

        [['Call', r.c_oi||0, 'var(--call)'], ['Put', r.p_oi||0, 'var(--put)']].forEach(([side, val, col]) => {
          g.append('rect')
            .attr('x', x + subBand(side)).attr('y', ySc(val))
            .attr('width', subBand.bandwidth()).attr('height', ih - ySc(val))
            .attr('fill', col).attr('opacity', isMP ? 1 : 0.65).attr('rx', 2)
            .on('mouseover', ev => showTip(
              `Strike: <b>${fmtN(r.strike)}</b><br/>
               ${side} OI: <b>${fmtN(val)}</b>${isMP ? '<br/><b style="color:var(--green)">★ Max Pain</b>' : ''}`, ev))
            .on('mousemove', moveTip).on('mouseout', hideTip);
        });

        if (isATM) {
          g.append('text').attr('x', x + xBand.bandwidth()/2).attr('y', ih + 28)
            .attr('text-anchor','middle').attr('fill','var(--amber)')
            .attr('font-size','8px').attr('font-family','var(--mono)').text('ATM');
        }
        if (isMP) {
          g.append('text').attr('x', x + xBand.bandwidth()/2).attr('y', ih + 38)
            .attr('text-anchor','middle').attr('fill','var(--green)')
            .attr('font-size','8px').attr('font-family','var(--mono)').text('MaxPain');
        }
      });

      g.append('g').attr('class','axis').attr('transform',`translate(0,${ih})`)
       .call(d3.axisBottom(xBand).tickSize(3))
       .selectAll('text')
       .attr('transform','rotate(-35)').attr('text-anchor','end')
       .attr('dx','-4px').attr('dy','4px');
      g.append('g').attr('class','axis')
       .call(d3.axisLeft(ySc).ticks(4).tickFormat(fmtAxis));
    })();
  }


  /* ════════════════════════════════════════════════════════
     IV SKEW TAB
  ════════════════════════════════════════════════════════ */

  function renderIVSkewTab(data) {
    const chain = data.chain;
    if (!chain || !chain.data || !chain.data.length) return;

    const rows   = [...chain.data]
      .filter(r => r.c_iv != null && r.p_iv != null)
      .sort((a, b) => a.strike - b.strike);
    if (!rows.length) return;

    const atm    = chain.atm;
    const atmIdx = rows.findIndex(r => r.strike === atm);

    /* ── IV metrics strip ── */
    const atmRow  = rows.find(r => r.strike === atm) || rows[Math.floor(rows.length/2)];
    const avgSkew = d3.mean(rows, r => r.iv_skew || 0);
    const maxSkew = d3.max(rows, r => r.iv_skew || 0);
    const minSkew = d3.min(rows, r => r.iv_skew || 0);
    const atmSkew = atmRow ? (atmRow.iv_skew ?? 0) : 0;

    document.getElementById('iv-strip').innerHTML = [
      { val: atmRow ? d3.format('.2f')(atmRow.c_iv) + '%' : '—', lbl: 'ATM Call IV',     cls: 'c-call'  },
      { val: atmRow ? d3.format('.2f')(atmRow.p_iv) + '%' : '—', lbl: 'ATM Put IV',      cls: 'c-put'   },
      { val: d3.format('.2f')(atmSkew) + '%',                     lbl: 'ATM Skew (P−C)',  cls: atmSkew > 2 ? 'c-put' : atmSkew < 0 ? 'c-call' : 'c-green' },
      { val: d3.format('.2f')(avgSkew) + '%',                     lbl: 'Avg Skew',        cls: 'c-muted' },
      { val: d3.format('.2f')(maxSkew) + '%',                     lbl: 'Max Skew',        cls: 'c-put'   },
      { val: d3.format('.2f')(minSkew) + '%',                     lbl: 'Min Skew',        cls: 'c-call'  },
    ].map(c => `<div class="metric-card">
      <div class="metric-val ${c.cls}">${c.val}</div>
      <div class="metric-lbl">${c.lbl}</div>
    </div>`).join('');

    /* ── IV Smile ── */
    (function() {
      const el = document.getElementById('chart-iv-smile');
      const W  = Math.max(el.parentElement.clientWidth - 32, 860);
      const H  = 240;
      const mg = { top: 16, right: 20, bottom: 40, left: 52 };
      const iw = W - mg.left - mg.right;
      const ih = H - mg.top  - mg.bottom;

      el.setAttribute('viewBox', `0 0 ${W} ${H}`);
      el.setAttribute('height', H);
      const svg = d3.select(el); svg.selectAll('*').remove();
      const g   = svg.append('g').attr('transform', `translate(${mg.left},${mg.top})`);

      const allIV = rows.flatMap(r => [r.c_iv, r.p_iv]).filter(v => v != null);
      const xSc   = d3.scaleLinear()
        .domain([d3.min(rows, r => r.strike), d3.max(rows, r => r.strike)])
        .range([0, iw]);
      const ySc   = d3.scaleLinear()
        .domain([Math.max(0, d3.min(allIV) * 0.9), d3.max(allIV) * 1.08])
        .range([ih, 0]);

      // Grid
      g.append('g').attr('class','grid')
       .call(d3.axisLeft(ySc).tickSize(-iw).tickFormat('').ticks(5));

      // ATM line
      g.append('line')
        .attr('x1', xSc(atm)).attr('x2', xSc(atm))
        .attr('y1', 0).attr('y2', ih)
        .attr('stroke','var(--amber)').attr('stroke-dasharray','5,3')
        .attr('stroke-width', 1.5).attr('opacity', 0.7);
      g.append('text').attr('x', xSc(atm) + 4).attr('y', 10)
        .attr('fill','var(--amber)').attr('font-size','9px')
        .attr('font-family','var(--mono)').text('ATM');

      // Area fills (subtle)
      const callArea = d3.area()
        .x(d => xSc(d.strike)).y0(ih).y1(d => ySc(d.c_iv))
        .curve(d3.curveCatmullRom).defined(d => d.c_iv != null);
      const putArea = d3.area()
        .x(d => xSc(d.strike)).y0(ih).y1(d => ySc(d.p_iv))
        .curve(d3.curveCatmullRom).defined(d => d.p_iv != null);

      g.append('path').datum(rows).attr('fill','rgba(59,130,246,0.06)').attr('d', callArea);
      g.append('path').datum(rows).attr('fill','rgba(249,115,22,0.06)').attr('d', putArea);

      // Lines
      const mkLine = key => d3.line()
        .x(d => xSc(d.strike)).y(d => ySc(d[key]))
        .curve(d3.curveCatmullRom).defined(d => d[key] != null);

      g.append('path').datum(rows)
        .attr('fill','none').attr('stroke','var(--call)')
        .attr('stroke-width', 2).attr('d', mkLine('c_iv'));
      g.append('path').datum(rows)
        .attr('fill','none').attr('stroke','var(--put)')
        .attr('stroke-width', 2).attr('d', mkLine('p_iv'));

      // Axes
      g.append('g').attr('class','axis').attr('transform',`translate(0,${ih})`)
       .call(d3.axisBottom(xSc).ticks(10).tickFormat(v => fmtN(v)));
      g.append('g').attr('class','axis')
       .call(d3.axisLeft(ySc).ticks(5).tickFormat(v => d3.format('.1f')(v) + '%'));

      // Hover overlay
      const bisect = d3.bisector(d => d.strike).left;
      const vline  = g.append('line').attr('y1',0).attr('y2',ih)
        .attr('stroke','var(--border2)').attr('stroke-dasharray','3,3').attr('opacity',0);
      const dotC = g.append('circle').attr('r',4)
        .attr('fill','var(--call)').attr('stroke','var(--surface)').attr('stroke-width',2)
        .attr('opacity',0);
      const dotP = g.append('circle').attr('r',4)
        .attr('fill','var(--put)').attr('stroke','var(--surface)').attr('stroke-width',2)
        .attr('opacity',0);

      g.append('rect').attr('width',iw).attr('height',ih)
        .attr('fill','transparent').style('cursor','crosshair')
        .on('mousemove', function(ev) {
          const [mx] = d3.pointer(ev, this);
          const s    = xSc.invert(mx);
          const idx  = Math.min(bisect(rows, s), rows.length - 1);
          const d    = rows[idx];
          vline.attr('x1',xSc(d.strike)).attr('x2',xSc(d.strike)).attr('opacity',1);
          if (d.c_iv != null) { dotC.attr('cx',xSc(d.strike)).attr('cy',ySc(d.c_iv)).attr('opacity',1); }
          if (d.p_iv != null) { dotP.attr('cx',xSc(d.strike)).attr('cy',ySc(d.p_iv)).attr('opacity',1); }
          showTip(`Strike: <b>${fmtN(d.strike)}</b><br/>
            Call IV: <b style="color:var(--call)">${d.c_iv != null ? d3.format('.2f')(d.c_iv)+'%' : '—'}</b><br/>
            Put IV:  <b style="color:var(--put)">${d.p_iv != null ? d3.format('.2f')(d.p_iv)+'%' : '—'}</b><br/>
            Skew (P-C): <b>${d.iv_skew != null ? d3.format('.2f')(d.iv_skew)+'%' : '—'}</b>`, ev);
        })
        .on('mouseout', () => {
          vline.attr('opacity',0); dotC.attr('opacity',0); dotP.attr('opacity',0); hideTip();
        });
    })();

    /* ── Put-Call IV Differential ── */
    (function() {
      const el = document.getElementById('chart-iv-diff');
      const W  = Math.max(el.parentElement.clientWidth - 32, 400);
      const H  = 230;
      const mg = { top: 16, right: 8, bottom: 40, left: 52 };
      const iw = W - mg.left - mg.right;
      const ih = H - mg.top  - mg.bottom;

      el.setAttribute('viewBox', `0 0 ${W} ${H}`);
      el.setAttribute('height', H);
      const svg = d3.select(el); svg.selectAll('*').remove();
      const g   = svg.append('g').attr('transform', `translate(${mg.left},${mg.top})`);

      const validRows = rows.filter(r => r.iv_skew != null);
      const absMax    = Math.max(d3.max(validRows, r => Math.abs(r.iv_skew)), 1);

      const xSc = d3.scaleLinear()
        .domain([d3.min(validRows, r => r.strike), d3.max(validRows, r => r.strike)])
        .range([0, iw]);
      const ySc = d3.scaleLinear()
        .domain([-absMax * 1.1, absMax * 1.1]).range([ih, 0]);

      g.append('g').attr('class','grid')
       .call(d3.axisLeft(ySc).tickSize(-iw).tickFormat('').ticks(5));

      // Zero line
      g.append('line')
        .attr('x1',0).attr('x2',iw)
        .attr('y1',ySc(0)).attr('y2',ySc(0))
        .attr('stroke','var(--border2)').attr('stroke-dasharray','4,3');

      // ATM line
      g.append('line')
        .attr('x1', xSc(atm)).attr('x2', xSc(atm))
        .attr('y1', 0).attr('y2', ih)
        .attr('stroke','var(--amber)').attr('stroke-dasharray','4,3').attr('opacity',0.7);

      // Bars
      const barW = Math.max((iw / validRows.length) * 0.7, 1);
      validRows.forEach(r => {
        const v   = r.iv_skew ?? 0;
        const pos = v >= 0;
        const col = v > 5 ? '#EF4444' : v < 0 ? '#3B82F6' : 'rgba(249,115,22,0.7)';
        const bH  = Math.abs(ySc(v) - ySc(0));
        g.append('rect')
          .attr('x', xSc(r.strike) - barW/2)
          .attr('y', pos ? ySc(v) : ySc(0))
          .attr('width', barW).attr('height', Math.max(bH, 1))
          .attr('fill', col).attr('opacity', 0.8).attr('rx', 1)
          .on('mouseover', ev => showTip(
            `Strike: <b>${fmtN(r.strike)}</b><br/>
             Skew (P−C): <b>${d3.format('.2f')(v)}%</b><br/>
             ${v > 5 ? '🔴 Extreme put skew' : v < 0 ? '🔵 Call expensive' : '🟡 Normal'}`, ev))
          .on('mousemove', moveTip).on('mouseout', hideTip);
      });

      g.append('g').attr('class','axis').attr('transform',`translate(0,${ih})`)
       .call(d3.axisBottom(xSc).ticks(8).tickFormat(v => fmtN(v)));
      g.append('g').attr('class','axis')
       .call(d3.axisLeft(ySc).ticks(5).tickFormat(v => d3.format('.1f')(v)+'%'));

      // Threshold annotation
      if (absMax > 5) {
        g.append('line')
          .attr('x1',0).attr('x2',iw)
          .attr('y1',ySc(5)).attr('y2',ySc(5))
          .attr('stroke','#EF4444').attr('stroke-dasharray','3,3').attr('opacity',0.5);
        g.append('text').attr('x',4).attr('y',ySc(5)-3)
          .attr('fill','#EF4444').attr('font-size','9px')
          .attr('font-family','var(--mono)').text('+5% threshold');
      }
    })();

    /* ── ATM Region IV — Zoomed ── */
    (function() {
      const n    = 10;
      const lo   = Math.max(0, atmIdx - n);
      const hi   = Math.min(rows.length - 1, atmIdx + n);
      const sub  = rows.slice(lo, hi + 1);
      if (!sub.length) return;

      document.getElementById('iv-zoom-sub').textContent =
        `${sub[0].strike} – ${sub[sub.length-1].strike} · ${sub.length} strikes`;

      const el = document.getElementById('chart-iv-zoom');
      const W  = Math.max(el.parentElement.clientWidth - 32, 400);
      const H  = 230;
      const mg = { top: 16, right: 20, bottom: 40, left: 52 };
      const iw = W - mg.left - mg.right;
      const ih = H - mg.top  - mg.bottom;

      el.setAttribute('viewBox', `0 0 ${W} ${H}`);
      el.setAttribute('height', H);
      const svg = d3.select(el); svg.selectAll('*').remove();
      const g   = svg.append('g').attr('transform', `translate(${mg.left},${mg.top})`);

      const allIV = sub.flatMap(r => [r.c_iv, r.p_iv]).filter(v => v != null);
      const xSc   = d3.scaleLinear()
        .domain([sub[0].strike, sub[sub.length-1].strike]).range([0, iw]);
      const ySc   = d3.scaleLinear()
        .domain([Math.max(0, d3.min(allIV)*0.92), d3.max(allIV)*1.08]).range([ih, 0]);

      g.append('g').attr('class','grid')
       .call(d3.axisLeft(ySc).tickSize(-iw).tickFormat('').ticks(5));

      // ATM band
      g.append('rect')
        .attr('x', xSc(atm) - 12).attr('y', 0)
        .attr('width', 24).attr('height', ih)
        .attr('fill','rgba(245,158,11,0.08)');
      g.append('line')
        .attr('x1',xSc(atm)).attr('x2',xSc(atm))
        .attr('y1',0).attr('y2',ih)
        .attr('stroke','var(--amber)').attr('stroke-dasharray','5,3').attr('opacity',0.8);

      // Lines + dots
      [['c_iv','var(--call)'], ['p_iv','var(--put)']].forEach(([key, col]) => {
        const validSub = sub.filter(d => d[key] != null);
        g.append('path').datum(validSub)
          .attr('fill','none').attr('stroke', col).attr('stroke-width', 2)
          .attr('d', d3.line()
            .x(d => xSc(d.strike)).y(d => ySc(d[key]))
            .curve(d3.curveCatmullRom));
        g.selectAll(null).data(validSub).join('circle')
          .attr('cx', d => xSc(d.strike)).attr('cy', d => ySc(d[key]))
          .attr('r', 3.5).attr('fill', col)
          .attr('stroke','var(--surface)').attr('stroke-width', 1.5)
          .on('mouseover', (ev, d) => showTip(
            `Strike: <b>${fmtN(d.strike)}</b><br/>
             ${key === 'c_iv' ? 'Call' : 'Put'} IV: <b style="color:${col}">${d3.format('.2f')(d[key])}%</b><br/>
             Skew: ${d.iv_skew != null ? d3.format('.2f')(d.iv_skew)+'%' : '—'}`, ev))
          .on('mousemove', moveTip).on('mouseout', hideTip);
      });

      g.append('g').attr('class','axis').attr('transform',`translate(0,${ih})`)
       .call(d3.axisBottom(xSc)
         .tickValues(sub.map(r => r.strike))
         .tickFormat(v => fmtN(v)))
       .selectAll('text')
       .attr('transform','rotate(-35)').attr('text-anchor','end')
       .attr('dx','-4px').attr('dy','4px');
      g.append('g').attr('class','axis')
       .call(d3.axisLeft(ySc).ticks(5).tickFormat(v => d3.format('.1f')(v)+'%'));

      g.append('text').attr('x',xSc(atm)).attr('y',-4)
        .attr('text-anchor','middle').attr('fill','var(--amber)')
        .attr('font-size','9px').attr('font-family','var(--mono)').text('ATM');
    })();
  }


  /* ════════════════════════════════════════════════════════
     HISTORICAL TRENDS TAB
  ════════════════════════════════════════════════════════ */

  function renderHistoryTab(data) {
    const oiData  = (data.oi  || {}).data  || {};
    const dodData = data.dod  || {};
    const sentData= (data.sentiment || {}).data || {};
    const dates   = (data.oi  || {}).dates || [];
    if (!dates.length) return;

    const SIG_COL  = { Bullish:'#10B981', Cautious:'#F97316', Bearish:'#EF4444', Neutral:'#94A3B8' };
    const SIG_ICON = { Bullish:'🟢', Cautious:'🟠', Bearish:'🔴', Neutral:'🟡' };

    /* ── Summary strip ── */
    const latest = dates[dates.length - 1];
    const fiiLatest = ((oiData[latest] || {}).FII) || {};
    document.getElementById('hist-strip').innerHTML = [
      { val: dates.length,                   lbl: 'Sessions Loaded',     cls: 'c-muted' },
      { val: dates[0],                       lbl: 'Earliest Session',    cls: 'c-muted' },
      { val: latest,                         lbl: 'Latest Session',      cls: 'c-call'  },
      { val: fmtN(fiiLatest.net_fut_idx),    lbl: 'FII Net Fut Idx (Latest)', cls: (fiiLatest.net_fut_idx||0) >= 0 ? 'c-green' : 'c-put' },
      { val: fmtF(fiiLatest.idx_oi_pcr),    lbl: 'FII Index PCR (Latest)',   cls: 'c-muted' },
      { val: (dodData.pairs||[]).length,     lbl: 'Session Pairs (DoD)',  cls: 'c-muted' },
    ].map(c => `<div class="metric-card">
      <div class="metric-val ${c.cls}">${c.val}</div>
      <div class="metric-lbl">${c.lbl}</div>
    </div>`).join('');

    /* ── H1: FII Net Index Futures (wide, prominent) ── */
    (function() {
      const el = document.getElementById('chart-fii-trend');
      const W  = Math.max(el.parentElement.clientWidth - 32, 860);
      const H  = 260;
      const mg = { top: 20, right: 20, bottom: 36, left: 80 };
      const iw = W - mg.left - mg.right;
      const ih = H - mg.top  - mg.bottom;

      el.setAttribute('viewBox', `0 0 ${W} ${H}`);
      el.setAttribute('height', H);
      const svg = d3.select(el); svg.selectAll('*').remove();
      const g   = svg.append('g').attr('transform', `translate(${mg.left},${mg.top})`);

      const vals = dates.map(d => ({
        date: d,
        val:  ((oiData[d] || {}).FII || {}).net_fut_idx ?? 0,
      }));
      const absMax = Math.max(d3.max(vals, v => Math.abs(v.val)), 1);
      const xSc = d3.scalePoint().domain(dates).range([0, iw]).padding(0.15);
      const ySc = d3.scaleLinear().domain([-absMax*1.15, absMax*1.15]).range([ih, 0]);

      // Positive/negative fill areas
      const areaPos = d3.area()
        .x(d => xSc(d.date)).y0(ySc(0)).y1(d => ySc(Math.max(d.val, 0)))
        .curve(d3.curveMonotoneX);
      const areaNeg = d3.area()
        .x(d => xSc(d.date)).y0(ySc(0)).y1(d => ySc(Math.min(d.val, 0)))
        .curve(d3.curveMonotoneX);

      g.append('path').datum(vals).attr('fill','rgba(16,185,129,0.15)').attr('d', areaPos);
      g.append('path').datum(vals).attr('fill','rgba(239,68,68,0.15)').attr('d', areaNeg);

      // Grid + zero line
      g.append('g').attr('class','grid')
       .call(d3.axisLeft(ySc).tickSize(-iw).tickFormat('').ticks(6));
      g.append('line')
        .attr('x1',0).attr('x2',iw).attr('y1',ySc(0)).attr('y2',ySc(0))
        .attr('stroke','var(--border2)').attr('stroke-width',1.5);

      // FII line
      g.append('path').datum(vals)
        .attr('fill','none').attr('stroke','var(--call)')
        .attr('stroke-width', 2.5)
        .attr('d', d3.line().x(d => xSc(d.date)).y(d => ySc(d.val)).curve(d3.curveMonotoneX));

      // Dots with coloured fill
      g.selectAll(null).data(vals).join('circle')
        .attr('cx', d => xSc(d.date)).attr('cy', d => ySc(d.val))
        .attr('r', 5)
        .attr('fill', d => d.val >= 0 ? '#10B981' : '#EF4444')
        .attr('stroke','var(--surface)').attr('stroke-width', 2)
        .on('mouseover', (ev, d) => showTip(
          `<b style="color:var(--call)">FII</b> · ${d.date}<br/>
           Net Fut Idx OI: <b>${fmtN(d.val)}</b><br/>
           ${d.val >= 0 ? '▲ Net Long' : '▼ Net Short'}`, ev))
        .on('mousemove', moveTip).on('mouseout', hideTip);

      // Value labels
      vals.forEach(d => {
        g.append('text')
          .attr('x', xSc(d.date)).attr('y', ySc(d.val) + (d.val >= 0 ? -10 : 16))
          .attr('text-anchor','middle')
          .attr('fill', d.val >= 0 ? '#10B981' : '#EF4444')
          .attr('font-size','10px').attr('font-family','var(--mono)').attr('font-weight','600')
          .text(d.val >= 1e6 ? d3.format('.1f')(d.val/1e6)+'M'
              : d.val <= -1e6 ? d3.format('.1f')(d.val/1e6)+'M'
              : d.val >= 1e3 ? d3.format('.0f')(d.val/1e3)+'K'
              : fmtN(d.val));
      });

      // Axes
      g.append('g').attr('class','axis').attr('transform',`translate(0,${ih})`)
       .call(d3.axisBottom(xSc).tickSize(4));
      g.append('g').attr('class','axis')
       .call(d3.axisLeft(ySc).ticks(6).tickFormat(fmtAxis));

      // Net Short / Net Long labels
      g.append('text').attr('x',4).attr('y',ySc(absMax*.85))
        .attr('fill','rgba(16,185,129,0.6)').attr('font-size','9px')
        .attr('font-family','var(--mono)').text('Net Long ↑');
      g.append('text').attr('x',4).attr('y',ySc(-absMax*.85))
        .attr('fill','rgba(239,68,68,0.6)').attr('font-size','9px')
        .attr('font-family','var(--mono)').text('Net Short ↓');
    })();

    /* ── H2a: Market OI Share ── */
    (function() {
      const el = document.getElementById('chart-mkt-share');
      const W  = Math.max(el.parentElement.clientWidth - 32, 400);
      const H  = 220;
      const mg = { top: 8, right: 80, bottom: 36, left: 40 };
      const iw = W - mg.left - mg.right;
      const ih = H - mg.top  - mg.bottom;

      el.setAttribute('viewBox', `0 0 ${W} ${H}`);
      el.setAttribute('height', H);
      const svg = d3.select(el); svg.selectAll('*').remove();
      const g   = svg.append('g').attr('transform', `translate(${mg.left},${mg.top})`);

      // Stack data: each date has 4 participant % values summing to 1
      const stackData = dates.map(d => {
        const entry = { date: d };
        PARTS.forEach(p => {
          entry[p] = ((oiData[d] || {})[p] || {}).mkt_share_long ?? 0;
        });
        return entry;
      });

      const stack   = d3.stack().keys(PARTS);
      const stacked = stack(stackData);

      const xSc  = d3.scalePoint().domain(dates).range([0, iw]).padding(0.3);
      const ySc  = d3.scaleLinear().domain([0, 1]).range([ih, 0]);
      const bw   = Math.min(xSc.step() * 0.55, 60);

      g.append('g').attr('class','grid')
       .call(d3.axisLeft(ySc).tickSize(-iw).tickFormat('').ticks(5));

      stacked.forEach(layer => {
        const part = layer.key;
        layer.forEach(d => {
          const [y0, y1] = d;
          const x = xSc(d.data.date);
          g.append('rect')
            .attr('x', x - bw/2).attr('y', ySc(y1))
            .attr('width', bw).attr('height', Math.max(ySc(y0) - ySc(y1), 0))
            .attr('fill', PART_COLORS[part]).attr('opacity', 0.8).attr('rx', 1)
            .on('mouseover', ev => showTip(
              `<b style="color:${PART_COLORS[part]}">${part}</b><br/>
               ${d.data.date}<br/>Market Share: <b>${d3.format('.1%')(y1 - y0)}</b>`, ev))
            .on('mousemove', moveTip).on('mouseout', hideTip);
        });
      });

      g.append('g').attr('class','axis').attr('transform',`translate(0,${ih})`)
       .call(d3.axisBottom(xSc).tickSize(3));
      g.append('g').attr('class','axis')
       .call(d3.axisLeft(ySc).ticks(5).tickFormat(d3.format('.0%')));

      // Legend
      PARTS.forEach((p, i) => {
        svg.append('rect').attr('x', W-mg.right+6).attr('y', mg.top + i*18)
          .attr('width',10).attr('height',10).attr('rx',2)
          .attr('fill', PART_COLORS[p]).attr('opacity',0.8);
        svg.append('text').attr('x', W-mg.right+20).attr('y', mg.top + i*18 + 9)
          .attr('fill','var(--muted)').attr('font-size','10px')
          .attr('font-family','var(--mono)').text(p);
      });
    })();

    /* ── H2b: Index OI PCR trend ── */
    (function() {
      const el = document.getElementById('chart-hist-pcr');
      const W  = Math.max(el.parentElement.clientWidth - 32, 400);
      const H  = 220;
      const mg = { top: 16, right: 20, bottom: 36, left: 48 };
      const iw = W - mg.left - mg.right;
      const ih = H - mg.top  - mg.bottom;

      el.setAttribute('viewBox', `0 0 ${W} ${H}`);
      el.setAttribute('height', H);
      const svg = d3.select(el); svg.selectAll('*').remove();
      const g   = svg.append('g').attr('transform', `translate(${mg.left},${mg.top})`);

      const series = PARTS.map(p => ({
        part: p,
        values: dates.map(d => ({
          date: d,
          val: ((oiData[d] || {})[p] || {}).idx_oi_pcr ?? null,
        })).filter(v => v.val != null),
      }));

      const allVals = series.flatMap(s => s.values.map(v => v.val));
      const yMax = Math.max(d3.max(allVals) * 1.1, 2.0);
      const yMin = Math.min(d3.min(allVals) * 0.9, 0.6);

      const xSc = d3.scalePoint().domain(dates).range([0, iw]).padding(0.1);
      const ySc = d3.scaleLinear().domain([yMin, yMax]).range([ih, 0]);

      // Bands
      if (0.8 >= yMin && 0.8 <= yMax)
        g.append('rect').attr('x',0).attr('width',iw)
          .attr('y',0).attr('height', ySc(yMin) - ySc(0.8))
          .attr('fill','rgba(16,185,129,0.05)');
      if (1.5 >= yMin && 1.5 <= yMax)
        g.append('rect').attr('x',0).attr('width',iw)
          .attr('y', ySc(yMax)).attr('height', ySc(1.5) - ySc(yMax))
          .attr('fill','rgba(239,68,68,0.05)');

      g.append('g').attr('class','grid')
       .call(d3.axisLeft(ySc).tickSize(-iw).tickFormat('').ticks(5));

      [{ v:0.8, c:'#10B981' }, { v:1.0, c:'#94A3B8' }, { v:1.5, c:'#EF4444' }].forEach(ref => {
        if (ref.v < yMin || ref.v > yMax) return;
        g.append('line')
          .attr('x1',0).attr('x2',iw).attr('y1',ySc(ref.v)).attr('y2',ySc(ref.v))
          .attr('stroke', ref.c).attr('stroke-dasharray','4,3').attr('opacity',0.5);
        g.append('text').attr('x', iw+3).attr('y', ySc(ref.v)+3)
          .attr('fill', ref.c).attr('font-size','9px').attr('font-family','var(--mono)')
          .text(d3.format('.1f')(ref.v));
      });

      const line = d3.line()
        .x(d => xSc(d.date)).y(d => ySc(d.val)).curve(d3.curveMonotoneX);

      series.forEach(s => {
        if (!s.values.length) return;
        g.append('path').datum(s.values)
          .attr('fill','none').attr('stroke', PART_COLORS[s.part])
          .attr('stroke-width', 2).attr('d', line);
        g.selectAll(null).data(s.values).join('circle')
          .attr('cx', d => xSc(d.date)).attr('cy', d => ySc(d.val))
          .attr('r', 3.5).attr('fill', PART_COLORS[s.part])
          .attr('stroke','var(--surface)').attr('stroke-width',1.5)
          .on('mouseover', (ev, d) => {
            const sig = d.val < 0.8 ? '🟢 Bullish' : d.val > 1.5 ? '🔴 Bearish' : '🟡 Neutral';
            showTip(`<b style="color:${PART_COLORS[s.part]}">${s.part}</b><br/>
              ${d.date}<br/>Index PCR: <b>${d3.format('.3f')(d.val)}</b><br/>${sig}`, ev);
          })
          .on('mousemove', moveTip).on('mouseout', hideTip);
      });

      g.append('g').attr('class','axis').attr('transform',`translate(0,${ih})`)
       .call(d3.axisBottom(xSc).tickSize(3));
      g.append('g').attr('class','axis')
       .call(d3.axisLeft(ySc).ticks(5).tickFormat(v => d3.format('.2f')(v)));

      const lg = document.getElementById('hist-pcr-legend');
      lg.innerHTML = PARTS.map(p =>
        `<div class="legend-item">
           <div class="legend-line" style="background:${PART_COLORS[p]}"></div>${p}
         </div>`).join('');
    })();

    /* ── H3: DoD Changes Summary ── */
    (function() {
      const pairs = dodData.pairs || [];
      if (!pairs.length) return;

      const el = document.getElementById('chart-hist-dod');
      const W  = Math.max(el.parentElement.clientWidth - 32, 860);
      const H  = 210;
      const mg = { top: 8, right: 20, bottom: 40, left: 72 };
      const iw = W - mg.left - mg.right;
      const ih = H - mg.top  - mg.bottom;

      el.setAttribute('viewBox', `0 0 ${W} ${H}`);
      el.setAttribute('height', H);
      const svg = d3.select(el); svg.selectAll('*').remove();
      const g   = svg.append('g').attr('transform', `translate(${mg.left},${mg.top})`);

      const pairLabels = pairs.map(([d0, d1]) => `${d0.slice(-6)}→${d1.slice(-6)}`);
      const rows = [];
      pairs.forEach(([d0, d1], pi) => {
        const pKey = `${d0}|${d1}`;
        const pData = (dodData.data || {})[pKey] || {};
        PARTS.forEach(p => {
          rows.push({
            pairIdx: pi,
            pairLabel: pairLabels[pi],
            part: p,
            val: (pData[p] || {}).d_net_total ?? 0,
          });
        });
      });

      const xBand  = d3.scaleBand().domain(pairLabels).range([0, iw]).padding(0.28);
      const subBand= d3.scaleBand().domain(PARTS).range([0, xBand.bandwidth()]).padding(0.08);
      const absMax = Math.max(d3.max(rows, r => Math.abs(r.val)), 1);
      const ySc    = d3.scaleLinear().domain([-absMax*1.15, absMax*1.15]).range([ih, 0]);

      g.append('g').attr('class','grid')
       .call(d3.axisLeft(ySc).tickSize(-iw).tickFormat('').ticks(5));
      g.append('line')
        .attr('x1',0).attr('x2',iw).attr('y1',ySc(0)).attr('y2',ySc(0))
        .attr('stroke','var(--border2)').attr('stroke-width',1.5);

      pairLabels.forEach(pl => {
        PARTS.forEach(p => {
          const row = rows.find(r => r.pairLabel === pl && r.part === p);
          const val = row ? row.val : 0;
          const pos = val >= 0;
          const bH  = Math.abs(ySc(val) - ySc(0));
          g.append('rect')
            .attr('x', xBand(pl) + subBand(p))
            .attr('y', pos ? ySc(val) : ySc(0))
            .attr('width', subBand.bandwidth())
            .attr('height', Math.max(bH, 1))
            .attr('fill', PART_COLORS[p])
            .attr('opacity', pos ? 0.85 : 0.4).attr('rx', 2)
            .on('mouseover', ev => showTip(
              `<b style="color:${PART_COLORS[p]}">${p}</b> · ${pl}<br/>
               Δ Net Total OI: <b>${fmtN(val)}</b><br/>
               ${pos ? '▲ Net long added' : '▼ Net short added'}`, ev))
            .on('mousemove', moveTip).on('mouseout', hideTip);
        });
      });

      g.append('g').attr('class','axis').attr('transform',`translate(0,${ih})`)
       .call(d3.axisBottom(xBand).tickSize(3));
      g.append('g').attr('class','axis')
       .call(d3.axisLeft(ySc).ticks(5).tickFormat(fmtAxis));

      const lg = document.getElementById('hist-dod-legend');
      lg.innerHTML = PARTS.map(p =>
        `<div class="legend-item">
           <div class="legend-dot" style="background:${PART_COLORS[p]}"></div>${p}
         </div>`).join('');
    })();

    /* ── H4: Key Signals Table ── */
    (function() {
      const thead = document.getElementById('hist-thead');
      const tbody = document.getElementById('hist-tbody');

      // Header: Participant | Metric | Date1 | Date2 | ...
      thead.innerHTML = `<tr>
        <th style="background:var(--surface2);padding:8px 12px;text-align:left;
          font-size:11px;color:var(--muted);font-weight:600;
          border-bottom:1px solid var(--border);text-transform:uppercase;letter-spacing:.04em">
          Participant</th>
        <th style="background:var(--surface2);padding:8px 12px;text-align:left;
          font-size:11px;color:var(--muted);font-weight:600;
          border-bottom:1px solid var(--border);text-transform:uppercase;letter-spacing:.04em">
          Metric</th>
        ${dates.map(d => `<th style="background:var(--surface2);padding:8px 12px;
          text-align:right;font-size:11px;color:var(--muted);font-weight:600;
          border-bottom:1px solid var(--border)">${d}</th>`).join('')}
      </tr>`;

      const METRICS = [
        { lbl: 'Net Fut Idx OI', key: 'net_fut_idx',  fmt: v => fmtN(v) },
        { lbl: 'Index OI PCR',   key: 'idx_oi_pcr',   fmt: v => v != null ? d3.format('.3f')(v) : '—' },
        { lbl: 'L/S Ratio',      key: 'ls_ratio',     fmt: v => v != null ? d3.format('.2f')(v)+'x' : '—' },
        { lbl: 'Signal',         key: '_signal',       fmt: null },
      ];

      const cellStyle = 'padding:6px 12px;border-top:1px solid rgba(42,51,71,0.5);font-size:11px;';
      const rows = [];
      PARTS.forEach((p, pi) => {
        METRICS.forEach((m, mi) => {
          let cells = dates.map(d => {
            const rec = ((oiData[d] || {})[p]) || {};
            const sig = ((sentData[d] || {})[p] || {}).signal || 'Neutral';
            if (m.key === '_signal') {
              const col  = SIG_COL[sig]  || '#94A3B8';
              const icon = SIG_ICON[sig] || '🟡';
              return `<td style="${cellStyle}text-align:right;color:${col};font-weight:600">
                ${icon} ${sig}</td>`;
            }
            const val = rec[m.key];
            let color = 'var(--text)';
            if (m.key === 'net_fut_idx') color = (val||0) >= 0 ? '#10B981' : '#EF4444';
            if (m.key === 'idx_oi_pcr')  color = val < 0.8 ? 'var(--call)' : val > 1.5 ? 'var(--put)' : 'var(--text)';
            return `<td style="${cellStyle}text-align:right;color:${color}">${m.fmt(val)}</td>`;
          }).join('');

          const isFirst  = mi === 0;
          const rowspan  = METRICS.length;
          const partCell = isFirst
            ? `<td rowspan="${rowspan}" style="${cellStyle}font-weight:700;
                color:${PART_COLORS[p]};vertical-align:middle;
                border-right:1px solid var(--border)">${p}</td>`
            : '';
          rows.push(`<tr>${partCell}<td style="${cellStyle}color:var(--muted)">${m.lbl}</td>${cells}</tr>`);
        });
      });
      tbody.innerHTML = rows.join('');
    })();
  }

  /* ── Call chart functions ── */
  const oiData = data.oi || {};
  drawTrend(oiData.data || {}, dates);
  drawBreakdown(oiData.data || {}, latest);
  drawDoD(data.dod || {}, (data.dod || {}).pairs || []);


  renderParticipantTab(data);
  renderPCRTab(data);
  renderBiasTab(data);

  window._dashData = data;

  /* ── Enable all tabs now data is ready ── */
  document.querySelectorAll('.tab-btn[disabled]').forEach(b => b.removeAttribute('disabled'));

  /* ── Footer ── */
  const genAt = (meta.generated_at || '').slice(0, 19).replace('T', ' ') + ' UTC';
  document.getElementById('footer-gen').textContent     = 'Generated ' + genAt;
  document.getElementById('footer-sessions').textContent= dates.length + ' sessions';
  document.getElementById('footer-strikes').textContent = ((chain.data||[]).length) + ' strikes';
  document.getElementById('site-footer').style.display  = 'flex';

  /* ── Debounced resize handler ── */
  let _resizeTimer;
  window.addEventListener('resize', () => {
    clearTimeout(_resizeTimer);
    _resizeTimer = setTimeout(() => {
      if (!window._dashData) return;
      const renderMap = {
        overview:    () => { drawTrend(window._dashData.oi.data, window._dashData.meta.oi_dates);
                             drawBreakdown(window._dashData.oi.data, window._dashData.meta.latest_date);
                             drawDoD(window._dashData.dod, (window._dashData.dod||{}).pairs||[]); },
        participant: () => renderParticipantTab(window._dashData),
        pcr:         () => renderPCRTab(window._dashData),
        bias:        () => renderBiasTab(window._dashData),
        chain:       () => renderChainTab(window._dashData),
        maxpain:     () => renderMaxPainTab(window._dashData),
        ivskew:      () => renderIVSkewTab(window._dashData),
        history:     () => renderHistoryTab(window._dashData),
      };
      if (renderMap[_activeTab]) renderMap[_activeTab]();
    }, 250);
  });

  console.log('[FAO Claude] All tabs rendered. D3 version:', d3.version);
}

loadData();
</script>
</body>
</html>
"""

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="FAO Website Generator")
    parser.add_argument("--data", required=True, help="Path to data.json")
    parser.add_argument("--out",  required=True, help="Path to write index.html")
    args = parser.parse_args()

    if not Path(args.data).exists():
        raise FileNotFoundError(f"data.json not found: {args.data}")

    with open(args.data, encoding="utf-8") as f:
        data = json.load(f)

    m = _meta_from_data(data)

    # Inject server-side values into the HTML (nav bar baked in, everything
    # else is fetched and rendered client-side by D3 / JS)
    html = HTML
    html = html.replace("LATEST_DATE",  m["latest"])
    html = html.replace("MAX_PAIN",     _fmt(m["max_pain"]))
    html = html.replace("OI_PCR",       _fmt(m["oi_pcr"], ".3f"))
    html = html.replace("CALL_WALL",    _fmt(m["call_wall"]))
    html = html.replace("PUT_WALL",     _fmt(m["put_wall"]))
    html = html.replace("EXCEL_FILE",   m["excel"])

    Path(args.out).parent.mkdir(parents=True, exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"✓ index.html written: {args.out}")
    print(f"  Latest session : {m['latest']}")
    print(f"  Max Pain       : {_fmt(m['max_pain'])}")
    print(f"  OI PCR         : {_fmt(m['oi_pcr'], '.3f')}")
    print(f"  Sessions       : {m['n_sessions']}  |  Strikes: {m['n_strikes']}")


if __name__ == "__main__":
    main()
