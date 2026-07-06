"""
ST+GENERATEWEBSITE~01.py
─────────────────────────
NSE FAO Options Dashboard — Website Generator
Reads docs/data.json, writes docs/index.html.

Step 2 placeholder: confirms pipeline is wired correctly.
Full D3 dashboard will be built in subsequent steps.
"""

import argparse
import json
from pathlib import Path
from datetime import datetime, timezone


def main():
    parser = argparse.ArgumentParser(description="FAO Website Generator")
    parser.add_argument("--data", required=True, help="Path to data.json")
    parser.add_argument("--out",  required=True, help="Path to write index.html")
    args = parser.parse_args()

    if not Path(args.data).exists():
        raise FileNotFoundError(f"data.json not found: {args.data}")

    with open(args.data, encoding="utf-8") as f:
        data = json.load(f)

    meta      = data.get("meta", {})
    chain     = data.get("chain") or {}
    sentiment = data.get("sentiment", {})
    latest    = meta.get("latest_date", "—")
    dates     = meta.get("oi_dates", [])
    excel_fn  = meta.get("excel_filename", "")
    generated = datetime.now(timezone.utc).strftime("%d %b %Y %H:%M UTC")

    # Quick sentiment table rows
    sent_latest = sentiment.get("data", {}).get(latest, {})
    sent_rows = ""
    signal_colour = {"Bullish": "#10B981", "Cautious": "#F97316",
                     "Bearish": "#EF4444", "Neutral":  "#94A3B8"}
    signal_emoji  = {"Bullish": "🟢", "Cautious": "🟠",
                     "Bearish": "🔴", "Neutral":  "🟡"}
    for part, s in sent_latest.items():
        sig   = s.get("signal", "Neutral")
        color = signal_colour.get(sig, "#94A3B8")
        emoji = signal_emoji.get(sig, "🟡")
        sent_rows += (
            f'<tr><td>{part}</td>'
            f'<td>{s.get("net_fut_idx_dir","—")}</td>'
            f'<td>{s.get("idx_oi_pcr") or "—"}</td>'
            f'<td>{s.get("ls_ratio") or "—"}</td>'
            f'<td style="color:{color};font-weight:600">{emoji} {sig}</td></tr>'
        )

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1.0"/>
  <title>FAO Claude — NSE Options Dashboard</title>
  <style>
    body{{font-family:system-ui,sans-serif;background:#0B0F1A;color:#F1F5F9;
          margin:0;padding:40px 24px;}}
    h1{{font-size:28px;margin-bottom:4px;}}
    .sub{{color:#94A3B8;font-size:14px;margin-bottom:40px;}}
    .grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:16px;margin-bottom:40px;}}
    .card{{background:#1E2535;border:1px solid #2A3347;border-radius:10px;padding:20px;}}
    .card-val{{font-size:28px;font-weight:700;margin-bottom:4px;}}
    .card-lbl{{font-size:12px;color:#94A3B8;text-transform:uppercase;letter-spacing:.05em;}}
    .blue{{color:#3B82F6;}} .amber{{color:#F59E0B;}} .green{{color:#10B981;}} .orange{{color:#F97316;}}
    table{{width:100%;border-collapse:collapse;background:#1E2535;
           border:1px solid #2A3347;border-radius:10px;overflow:hidden;}}
    th{{background:#374151;padding:10px 16px;text-align:left;font-size:12px;
        text-transform:uppercase;letter-spacing:.05em;color:#94A3B8;}}
    td{{padding:10px 16px;border-top:1px solid #2A3347;font-size:14px;}}
    .dl{{margin-top:32px;}}
    .btn{{display:inline-block;background:#3B82F6;color:#fff;padding:10px 22px;
          border-radius:8px;text-decoration:none;font-weight:600;font-size:14px;}}
    .notice{{background:#1E2535;border:1px solid #374151;border-radius:8px;
             padding:16px 20px;color:#94A3B8;font-size:13px;margin-top:40px;}}
    .dates{{font-family:monospace;font-size:12px;color:#4B5563;margin-top:8px;}}
  </style>
</head>
<body>
  <h1>FAO Claude — NSE Options Dashboard</h1>
  <div class="sub">Latest session: <strong>{latest}</strong> &nbsp;·&nbsp; Generated: {generated}</div>

  <!-- Key metrics -->
  <div class="grid">
    <div class="card">
      <div class="card-val amber">{chain.get('max_pain', '—'):,.0f}</div>
      <div class="card-lbl">Max Pain Strike</div>
    </div>
    <div class="card">
      <div class="card-val blue">{chain.get('atm', '—'):,.0f}</div>
      <div class="card-lbl">ATM Strike</div>
    </div>
    <div class="card">
      <div class="card-val green">{chain.get('oi_pcr', '—')}</div>
      <div class="card-lbl">Aggregate OI PCR</div>
    </div>
    <div class="card">
      <div class="card-val orange">{chain.get('call_walls', [{}])[0].get('strike', '—'):,.0f}</div>
      <div class="card-lbl">Top Call Wall</div>
    </div>
    <div class="card">
      <div class="card-val orange">{chain.get('put_walls', [{}])[0].get('strike', '—'):,.0f}</div>
      <div class="card-lbl">Top Put Wall</div>
    </div>
    <div class="card">
      <div class="card-val blue">{len(dates)}</div>
      <div class="card-lbl">Sessions loaded</div>
    </div>
  </div>

  <!-- Sentiment table -->
  <h2 style="margin-bottom:16px;font-size:18px;">Composite Sentiment — {latest}</h2>
  <table>
    <thead>
      <tr>
        <th>Participant</th>
        <th>Fut Index Direction</th>
        <th>Index OI PCR</th>
        <th>L/S Ratio</th>
        <th>Signal</th>
      </tr>
    </thead>
    <tbody>{sent_rows}</tbody>
  </table>

  <!-- Excel download -->
  <div class="dl">
    <a class="btn" href="../outputs/{excel_fn}" download>
      ⬇ Download Excel Dashboard ({excel_fn})
    </a>
  </div>

  <div class="notice">
    <strong>🚧 Full interactive D3 dashboard coming in Step 3.</strong><br/>
    This placeholder confirms the GitHub Actions pipeline is wired correctly —
    data flows from CSVs → data.json → website on every push to <code>data/</code>.
    <div class="dates">Sessions: {" · ".join(dates)}</div>
  </div>
</body>
</html>"""

    Path(args.out).parent.mkdir(parents=True, exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"✓ index.html written: {args.out}")


if __name__ == "__main__":
    main()
