# cubby-fcl-rates

Automated FCL rate tariff publisher for Cubby Cargo.

**Pipeline:**
```
You push data/rates.xlsx to GitHub
        ↓  GitHub Actions triggers automatically
        ↓  scripts/build_html.py reads the Excel
        ↓  Generates docs/index.html (card-style, grouped by trade lane)
        ↓  Deploys to GitHub Pages
        ↓  Respond.io syncs as Cubby Booking Buddy knowledge source
```

---

## Repo Structure

```
cubby-fcl-rates/
├── .github/
│   └── workflows/
│       └── build.yml        # GitHub Actions pipeline
├── scripts/
│   └── build_html.py        # Reads Excel → builds card-style HTML
├── data/
│   └── rates.xlsx           # ← your tariff file lives here
├── requirements.txt
└── README.md
```

---

## One-Time Setup

### 1. Create the repo on GitHub
- Go to github.com → **New repository** → name it `cubby-fcl-rates`
- Set to Public (required for free GitHub Pages)
- Don't initialise with any files — just create it empty

### 2. Upload these files
Push the following structure to the `main` branch:
```
.github/workflows/build.yml
scripts/build_html.py
data/rates.xlsx
requirements.txt
README.md
```

### 3. Enable GitHub Pages
- Go to repo **Settings → Pages**
- Source: **Deploy from a branch**
- Branch: `gh-pages` / `/ (root)`
- Save

### 4. Run the workflow once
- Go to **Actions → Build FCL Rate Tariff → Run workflow**
- After it completes, your tariff will be live at:
  `https://{your-github-username}.github.io/cubby-fcl-rates/`

### 5. Add to Respond.io
- In Respond.io, add the GitHub Pages URL above as a knowledge source
- Respond.io will re-crawl it on its sync schedule whenever rates update

---

## Updating Rates (Weekly)

1. Save your updated Excel as `rates.xlsx`
2. Push it to `data/rates.xlsx` in the repo (replace the existing file)
3. GitHub Actions detects the change and rebuilds automatically — nothing else needed

---

## Excel Format

The build script reads the **`TT` sheet** and expects these columns:
`Country, POL, POD, Container, Commodity, OF, THC, LAC, ISPS, Other Port Charges, GRI, Dredging Fee, Terminal Lease Surcharge, Destination Terminal Handling Charge, Local Handling, Admin, Total without Insurance, Insurance, Total with Insurance, Transit Time, Validity, Carrier, Agent`

If columns are added or renamed, update `SURCHARGE_COLS` and `SURCHARGE_LABELS` in `scripts/build_html.py`.

---

## Manual Rebuild

Go to repo **Actions → Build FCL Rate Tariff → Run workflow** at any time.
