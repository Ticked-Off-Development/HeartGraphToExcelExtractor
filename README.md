# HeartGraph Screenshot → Excel Extractor

Reads **HeartGraph Premium** screenshots and extracts heart rate zone data into an Excel spreadsheet matching Carol's HRM Daily format.

This repository contains two tools:

| Tool | File | Purpose |
|------|------|---------|
| **Extractor** | `heartgraph_extractor.py` | OCR screenshots → Excel (HRM Daily format) |
| **Dashboard** | `hrm_dashboard.py` | Interactive Streamlit dashboard for the Excel file |

## Overview

HeartGraph tracks your heart rate continuously and shows time spent in each of five heart rate zones. This script uses OCR to read screenshots exported from the app and compiles them into a structured Excel file for easy review and analysis.

## Requirements

- Python 3.8+
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) installed on your system
- Python packages:

```bash
pip install pytesseract Pillow openpyxl
```

### Installing Tesseract

- **macOS**: `brew install tesseract`
- **Ubuntu/Debian**: `sudo apt install tesseract-ocr`
- **Windows**: Download the installer from the [Tesseract GitHub releases](https://github.com/tesseract-ocr/tesseract/releases)

## Usage

```bash
python heartgraph_extractor.py <screenshot_folder> [output.xlsx]
```

| Argument | Description |
|---|---|
| `screenshot_folder` | Folder containing HeartGraph screenshots |
| `output.xlsx` | Output Excel file (default: `heartgraph_data.xlsx`) |

**Example:**

```bash
python heartgraph_extractor.py ~/Screenshots/heartgraph heartgraph_data.xlsx
```

## Screenshot Requirements

- **Supported formats**: `.png`, `.jpg`, `.jpeg`, `.bmp`, `.tiff`
- **File naming**: Files can be named anything — the date is read from the screenshot header (e.g., `20 Feb 22:46`)
- **Screenshot types**: Each day requires two screenshots:
  - **Zones screenshot** — shows time spent in each zone with percentages
  - **Summary screenshot** — shows heart rate range and mean heart rate
- Screenshots from the same date are automatically paired

## Date Handling

HeartGraph sessions that start in the evening (at or after 8:00 PM) are offset by **+1 day** in the spreadsheet. This reflects that an overnight tracking session started on the evening of one day corresponds to the following calendar day.

Sessions that start during the day are recorded on the same date shown in the screenshot.

## Output Format

The generated Excel file (`HRM Daily`) contains one row per day with the following columns:

| Column | Label | Description |
|---|---|---|
| A | Date | Full date (e.g., `Monday, February 21, 2025`) |
| B | red | Time in Zone 5 (highest HR) — H:MM:SS |
| C | orange | Time in Zone 4 — H:MM:SS |
| D | yellow | Time in Zone 3 — H:MM:SS |
| E | green | Time in Zone 2 — H:MM:SS |
| F | blue | Time in Zone 1 (lowest HR) — H:MM:SS |
| G | Max HR (bpm) | Maximum heart rate recorded |
| H | Mean 24 hr HR (bpm) | Mean heart rate over the session |

Zone columns are colour-coded to match their zone colour. The spreadsheet header row uses a blue theme consistent with Carol's original HRM Daily format.

## Multiple Sessions Per Day

If a day has more than one pair of screenshots (e.g., two separate tracking sessions), the script aggregates them as follows:

- **Zone times**: summed across all sessions
- **Mean HR**: averaged across sessions
- **Max HR**: highest value across sessions

Multi-session days are reported in the console output after processing.

## How It Works

1. **OCR extraction** — Tesseract reads text from each screenshot
2. **Date parsing** — The header region is cropped and searched for a date pattern (`DD Mon HH:MM`)
3. **Type classification** — Screenshots are identified as either a zones view or a summary view based on their content
4. **Data extraction** — Zone times and HR statistics are parsed using regex
5. **Aggregation** — Multiple sessions per day are combined
6. **Excel generation** — An `.xlsx` file is created with styled headers and colour-coded zone columns

## Troubleshooting

**"Could not extract date, skipping"**
The OCR could not find a recognisable date in the screenshot header. Ensure the screenshot includes the HeartGraph header bar and is not cropped.

**"Only found N zone values"**
The script found fewer than 5 zone time values. This may happen if the screenshot is blurry, low-contrast, or the zones view uses an unusual layout. Try taking a higher-resolution screenshot.

**"Unknown screenshot type"**
The script could not determine whether the screenshot is a zones or summary view. Check that it is an unmodified HeartGraph export.

**No data extracted at all**
Verify that Tesseract is installed and accessible from the command line (`tesseract --version`), and that the screenshots are valid HeartGraph exports in a supported image format.

---

## HRM Pacing Dashboard

An interactive local web dashboard that reads the HRM Daily Excel file and visualises heart rate zone data, pacing patterns, PEM/STABLE status, and medication events. Charts support zooming, hovering, and drill-down.

### Dashboard Requirements

```bash
pip install -r requirements_dashboard.txt
```

Dependencies: `streamlit`, `plotly`, `pandas`, `openpyxl`, `numpy`

### Running the Dashboard

```bash
streamlit run hrm_dashboard.py
```

The dashboard opens in your browser at `http://localhost:8501`. Use the sidebar to upload your HRM Daily `.xlsx` file.

### Sidebar Controls

| Control | Description |
|---------|-------------|
| File uploader | Upload your HRM Daily `.xlsx` file |
| Date range | Presets: Last 30/60/90 days, Last 6 months, This year, All time, or Custom |
| Tag filter | Multi-select to include/exclude PEM, STABLE, STABLE/PEM days |

### Dashboard Tabs

#### 📊 Overview
- KPI cards: days tracked, PEM days, STABLE days, average mean HR, average safe-zone %
- Mean HR over time with 7-day rolling average overlay; PEM days shaded in red
- Average zone distribution donut chart
- Recent 14-day data table

#### 🔵 Zone Analysis
- Stacked bar chart of daily zone time — toggle between **Hours** and **% of session**
- Drill-down selector: pick any zone to view its individual trend line with area fill and 7-day rolling average
- Zone summary table (average, max, and total time per zone)

#### 📈 HR Trends
- Two-panel chart (shared x-axis): Mean 24hr HR with 7-day average (top) and Max HR (bottom)
- PEM days shaded on both panels
- HR summary table grouped by status tag

#### 🛡️ Pacing & PEM
- Safe zone % over time: daily scatter + 7-day rolling mean line, with PEM shading
- Metric cards comparing safe-zone % overall vs PEM days vs STABLE days
- Grouped bar chart: average time in each zone broken down by tag (PEM / STABLE / etc.)
- Monthly summary table: days, PEM days, PEM %, average mean HR, average safe zone %
- Pre-PEM zone pattern table: average zone hours in the 3 days before PEM events vs baseline (shown automatically when ≥3 PEM days are in the selected range)

#### 💊 Events
- Mean HR timeline with colour-coded vertical annotations for each event in the **Other Events** column
- Event log table
- Before / After comparison: average mean HR in the 7 days before and after each event type, with delta

### Excel Columns Used by the Dashboard

The dashboard reads all columns the extractor produces (A–H) plus any additional columns you maintain manually:

| Column | Header | Source |
|--------|--------|--------|
| A | Date | Extractor |
| B | red | Extractor — Zone 5 (highest HR) |
| C | orange | Extractor — Zone 4 |
| D | yellow | Extractor — Zone 3 |
| E | green | Extractor — Zone 2 |
| F | blue | Extractor — Zone 1 (lowest HR) |
| G | Max HR (bpm) | Extractor |
| H | Mean 24 hr HR (bpm) | Extractor |
| I | Tags | Manual — e.g. `PEM`, `STABLE`, `STABLE/PEM` |
| J | Other Events | Manual — e.g. `Tirzepatide`, `Sirolimus` |
| K | 7-day average 24 hr HR | Manual or computed by dashboard if absent |

Column names are detected by substring match, so minor naming variations (e.g. `Tag` vs `Tags`) are handled automatically.
