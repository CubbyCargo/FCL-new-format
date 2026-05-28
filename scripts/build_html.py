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
OUTPUT_PATH       = Path("docs/index.html")
OUTPUT_JSON_PATH  = Path("docs/rates.json")
OUTPUT_QUOTE_PATH = Path("docs/quote.html")

SURCHARGE_COLS = [
    "OF", "THC", "LAC", "ISPS", "Other Port Charges",
    "GRI", "Dredging Fee", "Terminal Lease Surcharge",
    "Destination Terminal Handling Charge", "Local Handling", "Admin"
]

SURCHARGE_LABELS = {
    "OF": "Ocean Freight",
    "THC": "THC",
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
    build_json(cards)
    build_quote(cards)

def render_html(cards):
    generated = datetime.now(timezone.utc).strftime("%d %b %Y %H:%M UTC")

    # Build destination filter options
    dests = sorted(set(c["dest"] for c in cards))
    dest_options = "".join(f'<button class="filter-btn" data-dest="{d}">{d}</button>' for d in dests)

    # Build cards HTML
    cards_html = ""
    for c in cards:
        dest_slug = c["dest"].lower().replace(" ", "-").replace("&", "and").replace("(", "").replace(")", "")

        table_rows = ""
        for r in c["rate_rows"]:
            surcharge_breakdown = "".join(
                f'<div class="surcharge-item"><span class="surcharge-label">{k}</span><span class="surcharge-val">{fmt_usd(v)}</span></div>'
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
              <td class="surcharge-cell"><div class="surcharge-grid">{surcharge_breakdown}</div></td>
              <td class="total-cell">
                <div class="rate-block">
                  <div class="rate-no-ins">{fmt_usd(r["total_no_ins"])}<span class="rate-label">excl. insurance</span></div>
                  <div class="rate-with-ins">{fmt_usd(r["total_with_ins"])}<span class="rate-label ins-label">🛡 insured</span></div>
                </div>
              </td>
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
          <span class="meta-item">📅 {validity_range}</span>
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
                <th>Rate</th>
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
    --purple: #6b22d3;
    --purple-dark: #4e12a8;
    --purple-light: #f3eeff;
    --green: #8bea98;
    --green-dark: #2db84b;
    --green-bg: #f0fdf3;
    --white: #ffffff;
    --bg: #f8f7fc;
    --border: #e4ddf5;
    --text: #1a1030;
    --muted: #7a6e8a;
    --card-shadow: 0 2px 16px rgba(107,34,211,0.08);
  }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'Segoe UI', Arial, sans-serif; background: var(--bg); color: var(--text); font-size: 13px; }}

  /* ── Header ── */
  header {{ background: var(--purple); color: white; padding: 18px 32px; display: flex; align-items: center; justify-content: space-between; }}
  header h1 {{ font-size: 1.25rem; font-weight: 800; letter-spacing: 0.3px; }}
  header .generated {{ font-size: 11px; opacity: 0.65; }}
  .subtitle {{
    background: linear-gradient(90deg, var(--purple-dark), var(--purple));
    color: var(--green); padding: 5px 32px; font-size: 11px; font-weight: 700;
    letter-spacing: 1px; text-transform: uppercase;
  }}

  /* ── Filters ── */
  .controls {{
    padding: 14px 32px; background: var(--white); border-bottom: 1px solid var(--border);
    display: flex; gap: 8px; flex-wrap: wrap; align-items: center;
  }}
  .controls label {{ font-size: 10px; font-weight: 700; color: var(--muted); text-transform: uppercase; letter-spacing: 0.6px; margin-right: 4px; }}
  .filter-btn {{
    background: var(--white); border: 1.5px solid var(--border); border-radius: 20px;
    padding: 5px 14px; font-size: 12px; cursor: pointer; color: var(--muted);
    transition: all 0.15s; font-weight: 500;
  }}
  .filter-btn:hover {{ border-color: var(--purple); color: var(--purple); }}
  .filter-btn.active {{ background: var(--purple); color: white; border-color: var(--purple); font-weight: 600; }}
  .filter-btn.all-btn.active {{ background: var(--purple-dark); }}

  /* ── Cards ── */
  .main {{ padding: 24px 32px; display: flex; flex-direction: column; gap: 18px; }}
  .card {{ background: var(--white); border-radius: 12px; box-shadow: var(--card-shadow); overflow: hidden; border: 1.5px solid var(--border); }}
  .card-header {{
    background: linear-gradient(135deg, var(--purple) 0%, var(--purple-dark) 100%);
    color: white; padding: 14px 20px;
    display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 8px;
  }}
  .lane-title {{ display: flex; align-items: center; gap: 10px; font-size: 1rem; font-weight: 800; }}
  .dest-label {{ color: var(--green); }}
  .arrow {{ color: rgba(255,255,255,0.5); font-size: 1rem; }}
  .card-meta {{ display: flex; gap: 16px; font-size: 11px; opacity: 0.9; flex-wrap: wrap; align-items: center; }}
  .badge {{ background: rgba(255,255,255,0.15); border-radius: 4px; padding: 2px 8px; font-size: 10px; font-weight: 700; }}

  /* ── Table ── */
  .card-body {{ padding: 0; }}
  .table-wrapper {{ overflow-x: auto; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
  thead tr {{ background: var(--purple-light); }}
  th {{
    padding: 9px 12px; text-align: left; font-size: 10px; font-weight: 700;
    text-transform: uppercase; letter-spacing: 0.5px; color: var(--purple);
    border-bottom: 2px solid var(--border); white-space: nowrap;
  }}
  td {{ padding: 10px 12px; border-bottom: 1px solid #f0ecfa; vertical-align: top; }}
  tr:last-child td {{ border-bottom: none; }}
  tr:hover td {{ background: #faf8ff; }}

  /* ── Cell styles ── */
  .tag {{
    background: var(--purple); color: white; border-radius: 5px;
    padding: 3px 9px; font-size: 10px; font-weight: 700; white-space: nowrap;
  }}
  .pol-cell {{ font-weight: 600; color: var(--text); }}
  .pod-sub {{ font-size: 10px; color: var(--muted); font-weight: 400; margin-top: 2px; }}
  .commodity-cell {{ color: var(--muted); font-size: 11px; }}
  .carrier-cell {{ font-weight: 700; color: var(--purple); }}
  .agent-note {{ font-size: 10px; color: var(--muted); font-weight: 400; margin-top: 2px; }}
  .transit-cell, .validity-cell {{ white-space: nowrap; color: var(--muted); font-size: 11px; }}

  /* ── Surcharge breakdown ── */
  .surcharge-cell {{ min-width: 220px; max-width: 320px; }}
  .surcharge-grid {{ display: flex; flex-direction: column; gap: 3px; }}
  .surcharge-item {{
    display: flex; justify-content: space-between; align-items: center;
    background: var(--purple-light); border-radius: 4px; padding: 3px 8px;
    font-size: 11px;
  }}
  .surcharge-label {{ color: var(--muted); font-weight: 500; }}
  .surcharge-val {{ color: var(--purple); font-weight: 700; margin-left: 8px; }}

  /* ── Rate block (side by side totals) ── */
  .rate-block {{ display: flex; flex-direction: column; gap: 6px; min-width: 130px; }}
  .rate-no-ins {{
    display: flex; flex-direction: column;
    font-size: 14px; font-weight: 700; color: var(--muted);
    background: #f5f3fa; border-radius: 6px; padding: 6px 10px;
  }}
  .rate-with-ins {{
    display: flex; flex-direction: column;
    font-size: 16px; font-weight: 800; color: var(--green-dark);
    background: var(--green-bg); border: 1.5px solid var(--green);
    border-radius: 6px; padding: 6px 10px;
  }}
  .rate-label {{
    font-size: 9px; font-weight: 600; text-transform: uppercase;
    letter-spacing: 0.5px; color: var(--muted); margin-top: 2px;
  }}
  .ins-label {{ color: var(--green-dark); }}

  /* ── Notes ── */
  .notes {{
    background: var(--purple-light); border: 1.5px solid var(--border);
    border-radius: 10px; padding: 16px 22px; margin: 0 32px 28px;
    font-size: 11px; color: var(--purple-dark); line-height: 1.7;
  }}
  .notes strong {{ display: block; margin-bottom: 6px; font-size: 12px; color: var(--purple); }}

  .hidden {{ display: none !important; }}

  @media (max-width: 700px) {{
    header {{ padding: 14px 16px; }}
    .main {{ padding: 14px; }}
    .controls {{ padding: 12px 16px; }}
    .notes {{ margin: 0 14px 20px; }}
  }}
</style>
</head>
<body>
<header>
  <h1>🚢 Cubby Cargo — FCL Rate Tariff</h1>
  <span class="generated">Generated: {generated}</span>
</header>
<div class="subtitle">All-in rates (USD) · FCL only · Subject to space &amp; equipment availability</div>
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
  Rates are subject to space and equipment validity. Cargo must be ingated on or before the specified validity date; updated rates may apply if container is not ingated by said date.<br/>
  Marine Insurance covers a C&amp;F value of USD $30,000.00 at an additional <strong>$200</strong>. Values greater than USD $30,000 as well as restricted commodities must be quoted on a case-by-case basis.<br/>
  Should client decline Marine Insurance provided by Ramps Logistics, Ramps Logistics shall not be held liable for any claims, loss or damages arising from the execution of services.
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



# ─────────────────────────────────────────────
# rates.json
# ─────────────────────────────────────────────
def build_json(cards):
    generated = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    # Map dest label → short key matching old repo convention
    dest_key_map = {
        "Trinidad & Tobago": "TT",
        "Guyana":            "GUY",
        "Suriname":          "SUR",
        "Colombia (Buenaventura)": "COL",
        "Bridgetown":        "TT",
        "Kingston":          "TT",
        "Caucedo":           "TT",
    }

    destinations = {}
    for c in cards:
        dest_key = dest_key_map.get(c["dest"], c["dest"])
        lane_key = c["lane"].replace(" → ", " / ").replace(" & Tobago", "").upper()

        if dest_key not in destinations:
            destinations[dest_key] = {}

        rate_list = []
        for r in c["rate_rows"]:
            rate_list.append({
                "pol":                    r["pol"],
                "pod":                    r["pod"],
                "size":                   r["container"],
                "commodity":              r["commodity"],
                "carrier":                r["carrier"],
                "agent":                  r["agent"],
                "transit_time":           r["transit"],
                "validity":               r["validity"],
                "surcharges":             r["surcharges"],
                "total_without_insurance": r["total_no_ins"],
                "insurance":              200,
                "total_with_insurance":   r["total_with_ins"],
            })

        if lane_key not in destinations[dest_key]:
            destinations[dest_key][lane_key] = []
        destinations[dest_key][lane_key].extend(rate_list)

    payload = {
        "generated_at": generated,
        "source_file":  str(EXCEL_PATH.name),
        "destinations": destinations,
    }

    OUTPUT_JSON_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_JSON_PATH.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    total_rows = sum(len(v) for d in destinations.values() for v in d.values())
    print(f"✅ Written to {OUTPUT_JSON_PATH} ({len(destinations)} destinations, {total_rows} rate rows)")


# ─────────────────────────────────────────────
# quote.html
# ─────────────────────────────────────────────
def build_quote(cards):
    lanes_js = json.dumps([
        {
            "lane":   c["lane"],
            "origin": c["origin"],
            "dest":   c["dest"],
            "rates":  [
                {
                    "pol":            r["pol"],
                    "pod":            r["pod"],
                    "container":      r["container"],
                    "commodity":      r["commodity"],
                    "carrier":        r["carrier"],
                    "agent":          r["agent"],
                    "transit":        r["transit"],
                    "validity":       r["validity"],
                    "surcharges":     r["surcharges"],
                    "total_no_ins":   r["total_no_ins"],
                    "total_with_ins": r["total_with_ins"],
                }
                for r in c["rate_rows"]
            ],
        }
        for c in cards
    ], indent=2)

    dests = sorted(set(c["dest"] for c in cards))

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>Cubby Cargo — FCL Quote Generator</title>
<style>
  :root {{
    --purple: #6b22d3; --purple-dark: #4e12a8; --purple-light: #f3eeff;
    --green: #8bea98; --green-dark: #2db84b; --green-bg: #f0fdf3;
    --white: #ffffff; --bg: #f8f7fc; --border: #e4ddf5;
    --text: #1a1030; --muted: #7a6e8a; --shadow: 0 2px 16px rgba(107,34,211,0.10);
  }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'Segoe UI', Arial, sans-serif; background: var(--bg); color: var(--text); font-size: 13px; }}

  /* ── Screen-only elements ── */
  @media screen {{
    header {{ background: var(--purple); color: white; padding: 16px 32px; display: flex; align-items: center; justify-content: space-between; }}
    header h1 {{ font-size: 1.2rem; font-weight: 800; }}
    header a {{ color: var(--green); font-size: 12px; text-decoration: none; }}
    .subtitle {{ background: linear-gradient(90deg, var(--purple-dark), var(--purple)); color: var(--green); padding: 5px 32px; font-size: 11px; font-weight: 700; letter-spacing: 1px; text-transform: uppercase; }}
    .main {{ max-width: 900px; margin: 28px auto; padding: 0 20px; display: flex; flex-direction: column; gap: 20px; }}
    .form-card {{ background: var(--white); border-radius: 12px; border: 1.5px solid var(--border); box-shadow: var(--shadow); padding: 24px; }}
    .form-card h2 {{ font-size: 14px; font-weight: 700; color: var(--purple); margin-bottom: 18px; }}
    .form-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 14px; }}
    .field {{ display: flex; flex-direction: column; gap: 5px; }}
    .field label {{ font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; color: var(--muted); }}
    .field input, .field select {{
      border: 1.5px solid var(--border); border-radius: 7px; padding: 8px 11px;
      font-size: 13px; color: var(--text); background: var(--white); outline: none; transition: border-color 0.15s;
    }}
    .field input:focus, .field select:focus {{ border-color: var(--purple); }}
    .field select:disabled {{ background: #f5f3fa; color: var(--muted); }}
    .btn-search {{
      margin-top: 6px; background: var(--purple); color: white; border: none;
      border-radius: 8px; padding: 10px 28px; font-size: 13px; font-weight: 700; cursor: pointer;
    }}
    .btn-search:hover {{ background: var(--purple-dark); }}
    .results-card {{ background: var(--white); border-radius: 12px; border: 1.5px solid var(--border); box-shadow: var(--shadow); padding: 24px; display: none; }}
    .results-card h2 {{ font-size: 14px; font-weight: 700; color: var(--purple); margin-bottom: 16px; }}
    .rate-row {{
      border: 1.5px solid var(--border); border-radius: 10px; padding: 14px 16px; margin-bottom: 10px;
      display: grid; grid-template-columns: 1fr 1fr 1.5fr auto; gap: 12px; align-items: start; cursor: pointer;
    }}
    .rate-row:hover {{ border-color: var(--purple); background: var(--purple-light); }}
    .rate-row.selected {{ border-color: var(--green-dark); background: var(--green-bg); }}
    .btn-pdf {{
      background: var(--purple); color: white; border: none; border-radius: 8px;
      padding: 10px 24px; font-size: 13px; font-weight: 700; cursor: pointer;
    }}
    .btn-pdf:hover {{ background: var(--purple-dark); }}
    .quote-card {{ background: var(--white); border-radius: 12px; border: 1.5px solid var(--border); box-shadow: var(--shadow); padding: 28px; display: none; }}
  }}

  /* ── Print styles — quote only, one page, no nav ── */
  @media print {{
    @page {{ size: A4; margin: 18mm 18mm 14mm 18mm; }}
    header, .subtitle, .form-card, .results-card, .btn-pdf {{ display: none !important; }}
    body {{ background: white; font-size: 12px; }}
    .main {{ margin: 0; padding: 0; max-width: 100%; gap: 0; }}
    .quote-card {{
      display: block !important; border: none; box-shadow: none; padding: 0;
      border-radius: 0; page-break-inside: avoid;
    }}
    .ins-push, .quote-notes {{ -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
    .quote-lane-banner {{ -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
    .total-box {{ -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
    .detail-block {{ -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
    .print-logo {{ display: block !important; }}
  }}

  /* ── Shared styles (screen + print) ── */
  .print-logo {{ display: none; margin-bottom: 16px; }}
  .print-logo h2 {{ font-size: 1.1rem; font-weight: 800; color: var(--purple); }}
  .print-logo p {{ font-size: 10px; color: var(--muted); margin-top: 2px; }}
  .rate-meta {{ font-size: 11px; color: var(--muted); margin-top: 3px; }}
  .rate-carrier {{ font-weight: 700; color: var(--purple); }}
  .surcharge-list {{ font-size: 11px; }}
  .surcharge-line {{ display: flex; justify-content: space-between; padding: 2px 0; border-bottom: 1px dashed #ede8f8; }}
  .surcharge-line:last-child {{ border-bottom: none; }}
  .surcharge-line span:last-child {{ font-weight: 600; color: var(--purple); }}
  .totals {{ text-align: right; }}
  .total-no {{ font-size: 13px; color: var(--muted); margin-bottom: 6px; }}
  .total-with {{ font-size: 18px; font-weight: 800; color: var(--green-dark); background: var(--green-bg); border: 1.5px solid var(--green); border-radius: 7px; padding: 6px 12px; display: inline-block; }}
  .ins-note {{ font-size: 9px; color: var(--green-dark); display: block; margin-top: 2px; }}
  .tag {{ background: var(--purple); color: white; border-radius: 4px; padding: 2px 8px; font-size: 10px; font-weight: 700; }}
  .no-results {{ color: var(--muted); font-size: 13px; padding: 12px 0; }}
  .quote-header {{ display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 16px; flex-wrap: wrap; gap: 10px; }}
  .quote-title {{ font-size: 1.1rem; font-weight: 800; color: var(--purple); }}
  .quote-meta {{ font-size: 11px; color: var(--muted); margin-top: 4px; }}
  .quote-ref {{ font-size: 11px; color: var(--muted); text-align: right; }}
  .quote-lane-banner {{ background: linear-gradient(135deg, var(--purple), var(--purple-dark)); color: white; border-radius: 8px; padding: 11px 16px; display: flex; align-items: center; gap: 10px; font-size: 1rem; font-weight: 800; margin-bottom: 16px; }}
  .quote-lane-banner .dest {{ color: var(--green); }}
  .quote-details {{ display: grid; grid-template-columns: 1fr 1fr; gap: 14px; margin-bottom: 16px; }}
  .detail-block {{ background: var(--purple-light); border-radius: 8px; padding: 11px 13px; }}
  .detail-block h4 {{ font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; color: var(--purple); margin-bottom: 7px; }}
  .detail-line {{ display: flex; justify-content: space-between; font-size: 11px; padding: 3px 0; border-bottom: 1px solid #e4ddf5; }}
  .detail-line:last-child {{ border-bottom: none; }}
  .detail-line .val {{ font-weight: 600; }}
  .quote-totals {{ display: flex; gap: 12px; margin-bottom: 14px; }}
  .total-box {{ flex: 1; border-radius: 9px; padding: 12px 16px; text-align: center; }}
  .total-box.no-ins {{ background: #f5f3fa; border: 1.5px solid var(--border); }}
  .total-box.with-ins {{ background: var(--green-bg); border: 2px solid var(--green); }}
  .total-box .label {{ font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; color: var(--muted); margin-bottom: 5px; }}
  .total-box.with-ins .label {{ color: var(--green-dark); }}
  .total-box .amount {{ font-size: 20px; font-weight: 800; }}
  .total-box.no-ins .amount {{ color: var(--muted); }}
  .total-box.with-ins .amount {{ color: var(--green-dark); }}
  .ins-push {{ background: linear-gradient(90deg, var(--green-bg), #e8fded); border: 1.5px solid var(--green); border-radius: 8px; padding: 10px 14px; font-size: 11px; color: var(--green-dark); margin-bottom: 14px; }}
  .ins-push strong {{ font-size: 12px; }}
  .quote-notes {{ font-size: 10px; color: var(--muted); line-height: 1.6; background: var(--purple-light); border-radius: 8px; padding: 10px 14px; }}

  @media (max-width: 600px) {{
    .form-grid {{ grid-template-columns: 1fr; }}
    .rate-row {{ grid-template-columns: 1fr; }}
    .quote-details {{ grid-template-columns: 1fr; }}
    .quote-totals {{ flex-direction: column; }}
  }}
</style>
</head>
<body>
<header>
  <h1>🚢 Cubby Cargo — FCL Quote Generator</h1>
  <a href="index.html">← View Full Tariff</a>
</header>
<div class="subtitle">FCL All-in Rates (USD) · Subject to space &amp; equipment availability</div>

<div class="main">
  <!-- Search form -->
  <div class="form-card">
    <h2>📋 Shipment Details</h2>
    <div class="form-grid">
      <div class="field">
        <label>Prepared for</label>
        <input type="text" id="customer" placeholder="Customer name"/>
      </div>
      <div class="field">
        <label>Quote Reference</label>
        <input type="text" id="quoteref" placeholder="Auto-generated"/>
      </div>
      <div class="field">
        <label>Destination</label>
        <select id="sel-dest">
          <option value="">Select destination</option>
          {"".join(f'<option value="{d}">{d}</option>' for d in dests)}
        </select>
      </div>
      <div class="field">
        <label>Trade Lane</label>
        <select id="sel-lane" disabled>
          <option value="">Select lane</option>
        </select>
      </div>
      <div class="field">
        <label>Container Size</label>
        <select id="sel-size" disabled>
          <option value="">Select size</option>
        </select>
      </div>
    </div>
    <button class="btn-search" onclick="doSearch()">Search Rates</button>
  </div>

  <!-- Results -->
  <div class="results-card" id="results-card">
    <h2>Available Rates <span id="results-count" style="font-weight:400;color:var(--muted);font-size:12px;"></span></h2>
    <div id="results-list"></div>
  </div>

  <!-- Quote preview (screen + print) -->
  <div class="quote-card" id="quote-card">
    <!-- Print-only logo block -->
    <div class="print-logo">
      <h2>🚢 Cubby Cargo — FCL Shipping Quote</h2>
      <p>Ramps Logistics Ltd · ramps.co.tt · All rates in USD</p>
    </div>
    <div class="quote-header">
      <div>
        <div class="quote-title">FCL Shipping Quote</div>
        <div class="quote-meta" id="q-customer"></div>
      </div>
      <div class="quote-ref" id="q-ref"></div>
    </div>
    <div class="quote-lane-banner">
      <span id="q-origin"></span>
      <span style="color:rgba(255,255,255,0.5)">→</span>
      <span class="dest" id="q-dest"></span>
    </div>
    <div class="quote-details">
      <div class="detail-block">
        <h4>Shipment Info</h4>
        <div id="q-shipment-lines"></div>
      </div>
      <div class="detail-block">
        <h4>Surcharge Breakdown</h4>
        <div id="q-surcharge-lines"></div>
      </div>
    </div>
    <div class="quote-totals">
      <div class="total-box no-ins">
        <div class="label">Base Rate (excl. insurance)</div>
        <div class="amount" id="q-total-no"></div>
      </div>
      <div class="total-box with-ins">
        <div class="label">🛡 Total with Insurance</div>
        <div class="amount" id="q-total-with"></div>
      </div>
    </div>
    <div class="ins-push">
      🛡 <strong>Protect your cargo for just USD $200.</strong> Marine Insurance covers a C&amp;F value of up to USD $30,000. Ask your Cubby representative to include it in your booking.
    </div>
    <div class="quote-notes">
      Rates are subject to space and equipment validity. Cargo must be ingated on or before the validity date; updated rates may apply thereafter.
      Values greater than USD $30,000 and restricted commodities must be quoted on a case-by-case basis.
      Should client decline Marine Insurance, Ramps Logistics shall not be held liable for any claims, loss or damages arising from the execution of services.
    </div>
    <button class="btn-pdf" onclick="window.print()">🖨 Print / Save as PDF</button>
  </div>
</div>

<script>
const LANES = {lanes_js};

// Auto-generate quote ref on load
(function() {{
  const now = new Date();
  const yy = now.getFullYear();
  const mm = String(now.getMonth()+1).padStart(2,'0');
  const dd = String(now.getDate()).padStart(2,'0');
  const seq = String(Math.floor(Math.random()*900)+100);
  document.getElementById('quoteref').value = `QT-${{yy}}${{mm}}${{dd}}-${{seq}}`;
}})();

// When destination changes → populate lanes
document.getElementById('sel-dest').addEventListener('change', function() {{
  const dest = this.value;
  const laneEl = document.getElementById('sel-lane');
  const sizeEl = document.getElementById('sel-size');
  laneEl.innerHTML = '<option value="">Select lane</option>';
  sizeEl.innerHTML = '<option value="">Select size</option>';
  laneEl.disabled = true;
  sizeEl.disabled = true;
  if (!dest) return;
  const matching = LANES.filter(l => l.dest === dest);
  matching.forEach(l => {{
    const opt = document.createElement('option');
    opt.value = l.lane;
    opt.textContent = l.lane;
    laneEl.appendChild(opt);
  }});
  laneEl.disabled = false;
}});

// When lane changes → populate only available sizes
document.getElementById('sel-lane').addEventListener('change', function() {{
  const lane = this.value;
  const sizeEl = document.getElementById('sel-size');
  sizeEl.innerHTML = '<option value="">All sizes</option>';
  sizeEl.disabled = true;
  if (!lane) return;
  const laneData = LANES.find(l => l.lane === lane);
  if (!laneData) return;
  const sizes = [...new Set(laneData.rates.map(r => r.container))].sort();
  sizes.forEach(s => {{
    const opt = document.createElement('option');
    opt.value = s.startsWith('20') ? '20' : '40';
    opt.textContent = s;
    sizeEl.appendChild(opt);
  }});
  sizeEl.disabled = false;
}});

let selectedRate = null;

function doSearch() {{
  const dest = document.getElementById('sel-dest').value;
  const lane = document.getElementById('sel-lane').value;
  const size = document.getElementById('sel-size').value;

  let rates = [];
  LANES.forEach(l => {{
    if (dest && l.dest !== dest) return;
    if (lane && l.lane !== lane) return;
    l.rates.forEach(r => {{
      if (size === '20' && !r.container.startsWith('20')) return;
      if (size === '40' && !r.container.startsWith('40')) return;
      rates.push({{ ...r, _lane: l.lane, _origin: l.origin, _dest: l.dest }});
    }});
  }});

  const rc = document.getElementById('results-card');
  const list = document.getElementById('results-list');
  const count = document.getElementById('results-count');
  rc.style.display = 'block';
  count.textContent = `(${{rates.length}} result${{rates.length !== 1 ? 's' : ''}})`;

  if (rates.length === 0) {{
    list.innerHTML = '<div class="no-results">No rates found for that selection.</div>';
    return;
  }}

  list.innerHTML = rates.map((r, i) => {{
    const surchargeLines = Object.entries(r.surcharges || {{}}).map(([k,v]) =>
      `<div class="surcharge-line"><span>${{k}}</span><span>$${{v.toFixed(2)}}</span></div>`
    ).join('');
    return `
    <div class="rate-row" id="rr-${{i}}" onclick="selectRate(${{i}})">
      <div>
        <span class="tag">${{r.container}}</span>
        <div style="font-weight:600;margin-top:4px;">${{r.pol}}</div>
        <div class="rate-meta">→ ${{r.pod}}</div>
        <div class="rate-meta" style="margin-top:4px;">${{r.commodity}}</div>
      </div>
      <div>
        <div class="rate-carrier">${{r.carrier}}</div>
        ${{r.agent ? `<div class="rate-meta">${{r.agent}}</div>` : ''}}
        <div class="rate-meta" style="margin-top:4px;">⏱ ${{r.transit}} days</div>
        <div class="rate-meta">📅 Valid: ${{r.validity}}</div>
      </div>
      <div class="surcharge-list">${{surchargeLines}}</div>
      <div class="totals">
        <div class="total-no">$${{(r.total_no_ins||0).toLocaleString('en-US',{{minimumFractionDigits:2,maximumFractionDigits:2}})}}<br><small>excl. ins.</small></div>
        <div class="total-with">$${{(r.total_with_ins||0).toLocaleString('en-US',{{minimumFractionDigits:2,maximumFractionDigits:2}})}}<span class="ins-note">🛡 insured</span></div>
      </div>
    </div>`;
  }}).join('');

  window._searchRates = rates;
  selectedRate = null;
  document.getElementById('quote-card').style.display = 'none';
}}

function selectRate(i) {{
  document.querySelectorAll('.rate-row').forEach(el => el.classList.remove('selected'));
  document.getElementById('rr-' + i).classList.add('selected');
  selectedRate = window._searchRates[i];
  renderQuote(selectedRate);
}}

function fmt(v) {{
  if (!v && v !== 0) return '—';
  return '$' + v.toLocaleString('en-US', {{minimumFractionDigits:2, maximumFractionDigits:2}});
}}

function renderQuote(r) {{
  const customer = document.getElementById('customer').value || '—';
  const ref = document.getElementById('quoteref').value;
  const today = new Date().toLocaleDateString('en-GB', {{day:'2-digit',month:'short',year:'numeric'}});

  document.getElementById('q-customer').textContent = `Prepared for: ${{customer}} · Date: ${{today}}`;
  document.getElementById('q-ref').textContent = `Quote Ref: ${{ref}}`;
  document.getElementById('q-origin').textContent = r._origin;
  document.getElementById('q-dest').textContent = r._dest;

  document.getElementById('q-shipment-lines').innerHTML = [
    ['POL', r.pol], ['POD', r.pod], ['Container', r.container],
    ['Commodity', r.commodity || '—'],
    ['Carrier', r.carrier + (r.agent ? ` / ${{r.agent}}` : '')],
    ['Transit Time', r.transit ? r.transit + ' days' : '—'],
    ['Validity', r.validity || '—'],
  ].map(([k,v]) => `<div class="detail-line"><span>${{k}}</span><span class="val">${{v}}</span></div>`).join('');

  document.getElementById('q-surcharge-lines').innerHTML =
    Object.entries(r.surcharges || {{}}).map(([k,v]) =>
      `<div class="detail-line"><span>${{k}}</span><span class="val">$${{v.toFixed(2)}}</span></div>`
    ).join('') || '<div style="color:var(--muted);font-size:11px;">No breakdown available</div>';

  document.getElementById('q-total-no').textContent = fmt(r.total_no_ins);
  document.getElementById('q-total-with').textContent = fmt(r.total_with_ins);

  const qc = document.getElementById('quote-card');
  qc.style.display = 'block';
  qc.scrollIntoView({{behavior:'smooth', block:'start'}});
}}
</script>
</body>
</html>"""

    OUTPUT_QUOTE_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_QUOTE_PATH.write_text(html, encoding="utf-8")
    print(f"✅ Written to {OUTPUT_QUOTE_PATH}")

if __name__ == "__main__":
    build()
