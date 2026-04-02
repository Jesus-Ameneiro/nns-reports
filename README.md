# NNS Evidence Report Generator

Automated Evidence Report generation for Trimble SketchUp License Compliance cases.
Supports MCC (Mexico Centro Caribe) and CS (Cono Sur) regions.

## Project Structure

```
nns_app/
├── app.py              # Streamlit UI
├── processor.py        # Core data processing engine (VBA → Python)
├── report_writer.py    # Template filling logic (MCC + CS)
├── config.json         # Region/country mapping (edit freely)
├── requirements.txt    # Python dependencies
└── README.md
```

## Local Setup

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploying to Streamlit Community Cloud (Free)

1. Create a free account at https://share.streamlit.io
2. Push this folder to a GitHub repository
3. In Streamlit Cloud → "New app" → select your repo → set `app.py` as the entry point
4. Click Deploy — you'll get a shareable URL instantly

## How to Use

1. **Entity Name** — Full organization name
2. **Case ID(s)** — One or more external case IDs (click ＋ to add more)
3. **Country** — Select from dropdown; region (MCC/CS) is auto-detected
4. **Primary Domain** — Main website domain (e.g. `company.com`)
5. **Additional Domains** — Any other company domains for email/computer domain matching
6. **Machine Files** — Upload one or more `Exported Machines` Excel exports
7. **Case Event Files** — Upload one or more `Exported Case Events` Excel exports
8. **Template File** — Upload the correct regional template for the selected country
9. Click **Generate** → preview results → **Download** the filled report

## Config File (config.json)

Edit `config.json` to add/remove countries or regions at any time.
No code changes required.

```json
{
  "regions": {
    "MCC": {
      "name": "Mexico Centro Caribe",
      "countries": ["Mexico", "Guatemala", ...]
    },
    "CS": {
      "name": "Cono Sur",
      "countries": ["Argentina", "Colombia", ...]
    }
  }
}
```

## Processing Logic

- **Multiple machine files** → merged by Machine ID, values combined
- **Multiple event files** → merged, deduplicated by (Machine ID + Timestamp + Event Type + Product)
- **Machines sharing the same Active MAC** → grouped into one report row
- **Email selection** → Client Email (priority) → Additional Email, filtered by provided domains
- **Excluded machines** → 100% Education/Commercial/Evaluation events → flagged, not counted in aggregates
- **Empty columns** → Computer Domains and Client Email columns auto-deleted if all values are `-`
