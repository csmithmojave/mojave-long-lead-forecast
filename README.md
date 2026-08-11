# Mojave HVAC — Long Lead Parts Forecast

Self-contained HTML dashboard for tracking 120-day demand on parts with lead times > 59 days.

## Folder structure

```
Long lead planning/
├── build.py               ← run this to rebuild the dashboard
├── requirements.txt
├── .github/
│   └── workflows/
│       └── build.yml      ← auto-rebuilds when data/ files change
├── data/                  ← drop source files here
│   ├── ToExcel_Items*.csv
│   ├── CRM demand planning tool*.xlsx
│   └── ToExcel_PlanningDetail*.csv
└── output/
    └── demand-forecast.html  ← open this in Chrome
```

## Setup (one time)

### 1. Install Python dependencies
```bash
pip install -r requirements.txt
```

### 2. Add your source files
Copy your three export files into the `data/` folder:
| File type | Filename pattern |
|---|---|
| Items | `ToExcel_Items*.csv` or `items*.csv` |
| CRM forecast | `CRM demand planning tool*.xlsx` |
| Planning detail | `ToExcel_PlanningDetail*.csv` or `planning detail*.csv` |

The script always picks the **most recently modified** file matching each pattern, so you can drop new files in without deleting old ones.

### 3. Run the build
```bash
python build.py
```

Open `output/demand-forecast.html` in Chrome.

---

## Weekly update workflow

1. Export fresh files from ERP/CRM
2. Drop them into `data/`
3. Run `python build.py`
4. Open (or refresh) `output/demand-forecast.html`

If using GitHub:
- `git add data/ && git commit -m "Update 8.10" && git push`
- GitHub Actions rebuilds the HTML automatically
- No need to run build.py locally

---

## Commodity mapping

Edit `COMM_MAP` in `build.py` to add or rename commodities:

```python
COMM_MAP = {
    'Compressor':          ['compressor'],
    'Fans':                ['fans', 'fan'],
    'Evaporator Coils':    ['evaporator coil', 'evap co'],
    'Micro Channel Coils': ['micro channel', 'xcond', 'regen coil'],
    'Liquid Desiccant':    ['liquid desiccant', 'salt'],
    'VFD':                 ['vfd'],
}
```

Parts with unrecognized commodities appear in the **Missing Commodity** tab.

## Manual demand overrides

Edit `MANUAL_OVERRIDES` in `build.py`:

```python
MANUAL_OVERRIDES = {
    'P-002488': 7.0,
}
```
