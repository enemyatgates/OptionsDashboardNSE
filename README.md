# OptionsDashboardNSE

**Live dashboard:** https://enemyatgates.github.io/OptionsDashboardNSE

Automated NSE FAO options analysis — participant-wise OI, trading volume, and option chain data → Excel workbook + interactive website, generated on every push.

---

## How to use

1. Download your daily NSE FAO files and rename them:

| File | Rename to |
|---|---|
| Participant OI CSV | `FAOOIYYYYMMDD.csv` |
| Participant TV CSV | `FAOTVYYYYMMDD.csv` |
| Option Chain CSV | `FAOOCYYYYMMDD.csv` |

2. Push the renamed files into the `data/` folder.
3. GitHub Actions triggers automatically — Excel + website generated within ~60 seconds.

---

## Outputs

| Output | Location |
|---|---|
| Excel workbook | `outputs/FAOCLAUDEYYYYMMDD.xlsx` |
| Interactive website | https://enemyatgates.github.io/OptionsDashboardNSE |
| Raw data contract | `docs/data.json` |

---

## Repo structure

```
OptionsDashboardNSE/
├── data/                          ← push CSVs here
├── outputs/                       ← Excel files committed here after each run
├── docs/                          ← GitHub Pages (data.json + index.html)
├── ST+GENERATEDASHBOARD~01.py     ← Excel + data.json generator
├── ST+GENERATEWEBSITE~01.py       ← Website generator
└── .github/workflows/build.yml   ← CI/CD pipeline
```

---

## GitHub Pages setup

Go to **Settings → Pages → Source** and set:
- Branch: `main`
- Folder: `/docs`

---

## Requirements (local run)

```bash
pip install pandas openpyxl
python "ST+GENERATEDASHBOARD~01.py" --folder data
python "ST+GENERATEWEBSITE~01.py" --data docs/data.json --out docs/index.html
```
