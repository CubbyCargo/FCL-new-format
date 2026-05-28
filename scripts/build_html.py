"""
build_html.py
Reads the FCL rate tariff Excel (TT sheet) and generates a card-style HTML
grouped by origin country → destination combo (trade lane).
"""

import pandas as pd
import json
import sys
from pathlib import Path
from datetime import datetime, timezone
from collections import defaultdict

EXCEL_PATH = next(Path(".").glob("Customer Rate Tariff Template_Week *.xlsx"))
OUTPUT_PATH = Path("docs/index.html")

SURCHARGE_COLS = [
    "OF", "THC ", "LAC", "ISPS", "Other Port Charges",
    "GRI", "Dredging Fee", "Terminal Lease Surcharge",
    "Destination Terminal Handling Charge", "Local Handling", "Admin"
]

SURCHARGE_LABELS = {
    "OF": "Ocean Freight",
    "THC ": "THC",
    "LAC": "LAC",
    "ISPS": "ISPS",
    "Other Port Charges": "Other Port Charges",
    "GRI": "GRI",
    "Dredging Fee": "Dredging Fee",
    "Terminal Lease Surcharge": "Terminal Lease",
    "Destination Terminal Handling Charge": "Dest. THC",
    "Local Handling": "Local Handling",
    "Admin": "Admin",
}

# Map POD values to clean destination labels
def get_destination_label(pod):
    pod = str(pod).strip()
    tt_pods = {"Port of Spain", "Point Lisas", "Port-of-Spain"}
    gy_pods = {"Georgetown"}
    sr_pods = {"Paramaribo"}
    co_pods = {"Buenaventura", "Buenaventua "}
    if pod in tt_pods or "Port of Spain" in pod or "Point Lisas" in pod:
        return "Trinidad & Tobago"
    if pod in gy_pods:
        return "Guyana"
    if pod in sr_pods:
        return "Suriname"
    if any(p in pod for p in co_pods):
        return "Colombia (Buenaventura)"
    return pod

def clean_val(v):
    try:
        f = float(v)
        return f if f > 0 else None
    except (TypeError, ValueError):
        return None

def fmt_usd(v):
    if v is None:
        return "-"
    return f"${v:,.2f}".rstrip("0").rstrip(".")

def build():
    print(f"Reading {EXCEL_PATH}...")
    # Headers are on row index 2 (two blank rows at top)
    df = pd.read_excel(EXCEL_PATH, sheet_name="TT", header=None)
    # Find the header row by locating 'Country'
    header_row = None
    for i, row in df.iterrows():
        if str(row.iloc[1]).strip() == "Country":
            header_row = i
            break
    if header_row is None:
        print("ERROR: Could not find header row containing 'Country'", file=sys.stderr)
        sys.exit(1)
    df.columns = [str(c).strip() for c in df.iloc[header_row]]
    df = df.iloc[header_row + 1:].reset_index(drop=True)
    # Drop the unnamed first column (column 0 is blank)
    df = df.loc[:, df.columns != "nan"]
    df.columns = [str(c).strip() for c in df.columns]

    # Drop rows that are notes or completely empty
    df = df.dropna(subset=["Country", "POL", "POD", "Container"])
    df = df[~df["Country"].astype(str).str.startswith("Notes")]
    df = df[~df["Country"].astype(str).str.startswith("Rate")]

    df["_dest_label"] = df["POD"].apply(get_destination_label)
    df["_lane"] = df["Country"].astype(str).str.strip() + " → " + df["_dest_label"]

    # Group: lane → list of rate rows
    lanes = defaultdict(list)
    for _, row in df.iterrows():
        lanes[row["_lane"]].append(row)

    # Sort lanes: TT first, then GY, SR, CO, then alphabetical
    def lane_sort_key(lane):
        order = ["Trinidad & Tobago", "Guyana", "Suriname", "Colombia"]
        for i, dest in enumerate(order):
            if dest in lane:
                return (i, lane)
        return (99, lane)

    sorted_lanes = sorted(lanes.keys(), key=lane_sort_key)

    # Build cards data
    cards = []
    for lane in sorted_lanes:
        rows = lanes[lane]
        # Group by POL within the lane
        pol_groups = defaultdict(list)
        for row in rows:
            pol_groups[str(row["POL"]).strip()].append(row)

        # Collect unique carriers, validity, transit
        carriers = sorted(set(str(r.get("Carrier", "")).strip() for r in rows if str(r.get("Carrier", "")).strip()))
        validities = sorted(set(str(r.get("Validity ", str(r.get("Validity", "")))).strip() for r in rows if str(r.get("Validity ", str(r.get("Validity", "")))).strip() not in ["nan", ""]))
        transits = sorted(set(str(r.get("Transit Time ", str(r.get("Transit Time", "")))).strip() for r in rows if str(r.get("Transit Time ", str(r.get("Transit Time", "")))).strip() not in ["nan", ""]))

        rate_rows = []
        for pol, pol_rows in sorted(pol_groups.items()):
            for r in pol_rows:
                container = str(r.get("Container", "")).strip()
                commodity = str(r.get("Commodity", "")).strip()
                pod = str(r.get("POD", "")).strip()
                total_no_ins = clean_val(r.get("Total without Insurance"))
                total_with_ins = clean_val(r.get("Total with Insurance"))
                insurance = clean_val(r.get("Insurance "))
                carrier = str(r.get("Carrier", "")).strip()
                agent = str(r.get("Agent", "")).strip()
                transit = str(r.get("Transit Time ", str(r.get("Transit Time", "")))).strip()
                validity = str(r.get("Validity ", str(r.get("Validity", "")))).strip()

                surcharges = {}
                for col in SURCHARGE_COLS:
                    v = clean_val(r.get(col))
                    if v:
                        surcharges[SURCHARGE_LABELS.get(col, col)] = v

                rate_rows.append({
                    "pol": pol,
                    "pod": pod,
                    "container": container,
                    "commodity": commodity,
                    "carrier": carrier,
                    "agent": agent if agent not in ["nan", ""] else "",
                    "transit": transit if transit != "nan" else "",
                    "validity": validity if validity != "nan" else "",
                    "surcharges": surcharges,
                    "total_no_ins": total_no_ins,
                    "total_with_ins": total_with_ins,
                    "insurance": insurance,
                })

        origin, dest = lane.split(" → ", 1)
        cards.append({
            "lane": lane,
            "origin": origin,
            "dest": dest,
            "carriers": carriers,
            "validities": validities,
            "transits": transits,
            "rate_rows": rate_rows,
        })

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    html = render_html(cards)
    OUTPUT_PATH.write_text(html, encoding="utf-8")
    print(f"✅ Written to {OUTPUT_PATH} ({len(cards)} trade lanes, {sum(len(c['rate_rows']) for c in cards)} rate rows)")

def render_html(cards):
    generated = datetime.now(timezone.utc).strftime("%d %b %Y %H:%M UTC")

    # Build destination filter options
    dests = sorted(set(c["dest"] for c in cards))
    dest_options = "".join(f'<button class="filter-btn" data-dest="{d}">{d}</button>' for d in dests)

    # Build cards HTML
    cards_html = ""
    for c in cards:
        origin_slug = c["origin"].lower().replace(" ", "-").replace("&", "and")
        dest_slug = c["dest"].lower().replace(" ", "-").replace("&", "and").replace("(", "").replace(")", "")

        # Build rate table rows
        table_rows = ""
        for r in c["rate_rows"]:
            surcharge_breakdown = " + ".join(
                f'<span class="surcharge-item">{k}: {fmt_usd(v)}</span>'
                for k, v in r["surcharges"].items()
            )
            agent_note = f'<div class="agent-note">{r["agent"]}</div>' if r["agent"] else ""
            table_rows += f"""
            <tr>
              <td><span class="tag">{r["container"]}</span></td>
              <td class="pol-cell">{r["pol"]}<div class="pod-sub">→ {r["pod"]}</div></td>
              <td class="commodity-cell">{r["commodity"]}</td>
              <td class="carrier-cell">{r["carrier"]}{agent_note}</td>
              <td class="transit-cell">{r["transit"]}</td>
              <td class="validity-cell">{r["validity"]}</td>
              <td class="surcharge-cell">{surcharge_breakdown}</td>
              <td class="total-cell"><span class="total-no-ins">{fmt_usd(r["total_no_ins"])}</span></td>
              <td class="total-cell total-with"><strong>{fmt_usd(r["total_with_ins"])}</strong></td>
            </tr>"""

        carriers_badges = " ".join(f'<span class="badge">{cr}</span>' for cr in c["carriers"])
        transit_range = " / ".join(set(c["transits"])) if c["transits"] else "—"
        validity_range = " / ".join(set(c["validities"])) if c["validities"] else "—"

        cards_html += f"""
    <div class="card" data-origin="{c['origin']}" data-dest="{c['dest']}" data-dest-slug="{dest_slug}">
      <div class="card-header">
        <div class="lane-title">
          <span class="origin-label">{c['origin']}</span>
          <span class="arrow">→</span>
          <span class="dest-label">{c['dest']}</span>
        </div>
        <div class="card-meta">
          <span class="meta-item">🚢 {carriers_badges}</span>
          <span class="meta-item">⏱ {transit_range}</span>
          <span class="meta-item">📅 Valid: {validity_range}</span>
        </div>
      </div>
      <div class="card-body">
        <div class="table-wrapper">
          <table>
            <thead>
              <tr>
                <th>Container</th>
                <th>POL → POD</th>
                <th>Commodity</th>
                <th>Carrier</th>
                <th>Transit</th>
                <th>Validity</th>
                <th>Surcharge Breakdown</th>
                <th>Total (no ins.)</th>
                <th>Total (with ins.)</th>
              </tr>
            </thead>
            <tbody>{table_rows}</tbody>
          </table>
        </div>
      </div>
    </div>"""

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>Cubby Cargo – FCL Rate Tariff</title>
<style>
  :root {{
    --cubby-blue: #1a3a5c;
    --cubby-orange: #f4821f;
    --cubby-light: #f7f9fc;
    --border: #dde3ec;
    --text: #1e2a3a;
    --muted: #6b7a8d;
    --card-shadow: 0 2px 12px rgba(26,58,92,0.08);
  }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'Segoe UI', Arial, sans-serif; background: var(--cubby-light); color: var(--text); font-size: 13px; }}
  header {{ background: var(--cubby-blue); color: white; padding: 18px 32px; display: flex; align-items: center; justify-content: space-between; }}
  header h1 {{ font-size: 1.3rem; font-weight: 700; letter-spacing: 0.3px; }}
  header .generated {{ font-size: 11px; opacity: 0.7; }}
  .subtitle {{ background: var(--cubby-orange); color: white; padding: 6px 32px; font-size: 11px; font-weight: 600; letter-spacing: 0.5px; }}
  .controls {{ padding: 16px 32px; background: white; border-bottom: 1px solid var(--border); display: flex; gap: 10px; flex-wrap: wrap; align-items: center; }}
  .controls label {{ font-size: 11px; font-weight: 600; color: var(--muted); text-transform: uppercase; letter-spacing: 0.5px; margin-right: 4px; }}
  .filter-btn {{
    background: var(--cubby-light); border: 1px solid var(--border); border-radius: 20px;
    padding: 5px 14px; font-size: 12px; cursor: pointer; color: var(--text); transition: all 0.15s;
  }}
  .filter-btn:hover, .filter-btn.active {{ background: var(--cubby-blue); color: white; border-color: var(--cubby-blue); }}
  .filter-btn.all-btn.active {{ background: var(--cubby-orange); border-color: var(--cubby-orange); }}
  .main {{ padding: 24px 32px; display: flex; flex-direction: column; gap: 20px; }}
  .card {{ background: white; border-radius: 10px; box-shadow: var(--card-shadow); overflow: hidden; border: 1px solid var(--border); }}
  .card-header {{ background: var(--cubby-blue); color: white; padding: 14px 20px; display: flex; justify-content: space-between; align-items: flex-start; flex-wrap: wrap; gap: 8px; }}
  .lane-title {{ display: flex; align-items: center; gap: 10px; font-size: 1rem; font-weight: 700; }}
  .origin-label {{ font-size: 1rem; }}
  .dest-label {{ color: #f4d03f; font-size: 1rem; }}
  .arrow {{ color: var(--cubby-orange); font-size: 1.1rem; }}
  .card-meta {{ display: flex; gap: 14px; font-size: 11px; opacity: 0.9; flex-wrap: wrap; align-items: center; }}
  .meta-item {{ display: flex; align-items: center; gap: 4px; }}
  .badge {{ background: rgba(255,255,255,0.18); border-radius: 4px; padding: 2px 7px; font-size: 10px; font-weight: 600; }}
  .card-body {{ padding: 0; }}
  .table-wrapper {{ overflow-x: auto; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
  thead tr {{ background: #eef2f8; }}
  th {{ padding: 9px 12px; text-align: left; font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.4px; color: var(--muted); border-bottom: 2px solid var(--border); white-space: nowrap; }}
  td {{ padding: 8px 12px; border-bottom: 1px solid #f0f3f7; vertical-align: top; }}
  tr:last-child td {{ border-bottom: none; }}
  tr:hover td {{ background: #f7faff; }}
  .tag {{ background: var(--cubby-blue); color: white; border-radius: 4px; padding: 2px 8px; font-size: 10px; font-weight: 700; white-space: nowrap; }}
  .pol-cell {{ font-weight: 600; }}
  .pod-sub {{ font-size: 10px; color: var(--muted); font-weight: 400; margin-top: 2px; }}
  .commodity-cell {{ color: var(--muted); }}
  .carrier-cell {{ font-weight: 600; color: var(--cubby-blue); }}
  .agent-note {{ font-size: 10px; color: var(--muted); font-weight: 400; margin-top: 2px; }}
  .transit-cell, .validity-cell {{ white-space: nowrap; color: var(--muted); }}
  .surcharge-cell {{ max-width: 280px; }}
  .surcharge-item {{ display: inline-block; background: #f0f3f7; border-radius: 3px; padding: 1px 5px; margin: 1px; font-size: 10px; color: var(--muted); }}
  .total-no-ins {{ color: var(--muted); }}
  .total-with {{ background: #fff8f0; }}
  .total-with strong {{ color: var(--cubby-orange); font-size: 13px; }}
  .notes {{ background: #fffbe6; border: 1px solid #ffe082; border-radius: 8px; padding: 14px 20px; margin: 0 32px 24px; font-size: 11px; color: #7a6000; line-height: 1.6; }}
  .notes strong {{ display: block; margin-bottom: 4px; }}
  .hidden {{ display: none !important; }}
  @media (max-width: 700px) {{
    header {{ padding: 14px 16px; }}
    .main {{ padding: 16px; }}
    .controls {{ padding: 12px 16px; }}
  }}
</style>
</head>
<body>
<header>
  <h1>🚢 Cubby Cargo — FCL Rate Tariff</h1>
  <span class="generated">Generated: {generated}</span>
</header>
<div class="subtitle">ALL-IN RATES (USD) · FCL ONLY · SUBJECT TO SPACE & EQUIPMENT AVAILABILITY</div>
<div class="controls">
  <label>Filter by Destination:</label>
  <button class="filter-btn all-btn active" data-dest="all">All Destinations</button>
  {dest_options}
</div>
<div class="main" id="cards-container">
{cards_html}
</div>
<div class="notes">
  <strong>📋 Important Notes</strong>
  Rates are subject to space and equipment validity. Cargo must be ingated on or before the specified validity date; updated rates may apply if container is not ingated by said date.
  Insurance covers a C&amp;F value of USD $30,000.00. Values greater than USD $30,000 as well as restricted commodities must be quoted on a case-by-case basis.
  Should client refuse Marine Insurance provided by Ramps Logistics, Ramps Logistics shall not be held liable for any claims, loss or damages resulting from the execution of services quoted and accepted by client.
</div>
<script>
  const btns = document.querySelectorAll('.filter-btn');
  const cards = document.querySelectorAll('.card');
  btns.forEach(btn => {{
    btn.addEventListener('click', () => {{
      btns.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      const dest = btn.dataset.dest;
      cards.forEach(card => {{
        if (dest === 'all' || card.dataset.dest === dest) {{
          card.classList.remove('hidden');
        }} else {{
          card.classList.add('hidden');
        }}
      }});
    }});
  }});
</script>
</body>
</html>"""

if __name__ == "__main__":
    build()
