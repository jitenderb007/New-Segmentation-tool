# SegmentPro

**Customer Segmentation for Market Research** — a Windows desktop tool that turns any survey Excel/CSV file into a clean segmentation report and stakeholder-ready PowerPoint deck.

---

## What it does

1. **Upload** any Excel (`.xlsx`, `.xls`, `.xlsm`) or CSV file.
2. **Auto-detects** Likert-style attitudinal variables as segmentation basis (you can override).
3. **Cleans** the data — drops empty rows/cols, imputes missing with column median, removes zero-variance columns.
4. **Clusters** respondents using K-Means with silhouette-based k selection (3 to 7), 50 random initializations, reproducible seed.
5. **Names** segments automatically from their distinguishing attitudes (Outlet Devotees, Fashion-Forward Trendsetters, Premium Quality Seekers, etc.).
6. **Exports** a 7-sheet Excel workbook and a PowerPoint deck with title, overview, per-segment deep-dives, and methodology slides.

All in a 4-step wizard interface — no Python knowledge required for end users.

---

## Installing on a team member's Windows machine

### Option 1 — Distribute as a single `.exe` (recommended for end users)

On any machine with Python 3.9+ installed:

```powershell
pip install -r requirements.txt
pip install pyinstaller
pyinstaller --onefile --windowed --name SegmentPro SegmentPro.py
```

The bundled `SegmentPro.exe` in the `dist/` folder can be copied to any Windows machine and run directly — no Python install required on the end user's machine.

### Option 2 — Run from source (for analysts who already have Python)

```powershell
pip install -r requirements.txt
python SegmentPro.py
```

---

## Requirements

- Python 3.9 or newer (with Tcl/Tk — the default Windows Python installer includes it)
- The packages listed in `requirements.txt`

---

## Using the app

### Step 1 — Load Data
Click **Browse for file…** and pick your survey file. The app detects the most populated sheet (or picks the sheet named `Data`). A preview of the first 5 rows appears along with dataset stats (rows, columns, missing cells).

### Step 2 — Choose Variables
The app auto-selects numeric survey items with 3–10 response categories (standard Likert scales). Review the list:
- Click the checkmark column to toggle a variable in/out of the basis.
- Use **Filter** to search within the variable list.
- Use **Auto-detect** to reset to the recommended selection.

At least 3 variables must be selected to proceed.

### Step 3 — Run Analysis
- **Auto-select k** (recommended): tests k=3 to 7 and picks the best silhouette score.
- **Fixed k**: specify a value if you already know how many segments you want.

Click **▶ Run Segmentation**. Progress bar shows the stages.

### Step 4 — Results & Export
Review segments on-screen: size, description, and top 3 differentiators. Export to:
- **📊 Excel Workbook** — full 7-sheet report
- **🎯 PowerPoint Deck** — stakeholder presentation
- **📦 Export Both** — saves both to a folder you choose

---

## Excel output — what's in each sheet

| Sheet | Content |
|---|---|
| **Executive Summary** | Segment sizes, percentages, and narrative descriptions |
| **Demographics** | Age, gender, income, parent status by segment (uses auto-detected demographic columns) |
| **Segment Profiles** | Mean score on every basis variable by segment, with overall mean |
| **Key Differentiators** | Top 10 most-deviating variables per segment, with z-scores |
| **Z-Score Heatmap** | Full profile heatmap — green = above average, red = below |
| **Respondent Assignments** | Every respondent's segment — for downstream targeting |
| **Methodology** | Full technical write-up of the approach, k selection, silhouette scores |

---

## PowerPoint output — deck structure

1. **Title slide** — dataset summary
2. **Segment Overview** — all segments in card layout with size percentages
3. **One deep-dive slide per segment** — who they are, key differentiators, size
4. **Methodology** — process steps and silhouette score comparison
5. **Next Steps** — recommended follow-up actions

---

## How segment naming works

The tool matches the top-deviating variables of each segment against a library of common market-research archetypes:

| Archetype | Trigger |
|---|---|
| Outlet Devotees | Strong agreement with "I shop outlet stores regularly" |
| Fashion-Forward Trendsetters | "First to try trendy runway fashions" |
| Premium Quality Seekers | "I invest in higher quality fabrics" |
| Comfort-First Practicals | "Comfort is the most important factor" |
| Value-Conscious Deal-Seekers | "I actively seek out deals and coupons" |
| Brand-Loyal Classics | "I am loyal to the brands I buy" |
| Minimalist Needs-Based | "I only buy clothing when I need it" |
| Individualist Stylists | "I like my clothes to stand out from the crowd" |

If no archetype matches, the segment is named by its single strongest-deviating variable (so you still get something descriptive, not just "Segment 1").

---

## Troubleshooting

**"Missing Dependencies" error on launch**
Install the listed packages: `pip install scikit-learn openpyxl python-pptx`

**Preview doesn't show after loading Excel**
The app reads the sheet named `Data` first, or falls back to the sheet with the most rows. If your data is on a sheet with a different name and no `Data` sheet exists, it still works — check the log at the bottom of the app for which sheet was read.

**Too few variables detected**
Click the **Select All** button in Step 2 and manually un-check the columns that aren't attitudinal (IDs, dates, open-text fields, etc.).

**k auto-selection picks a number you don't like**
Switch to **Fixed k** in Step 3 and specify the number you want. The silhouette scores for every tested k are shown in the methodology slide so you can see how far off your choice is from the best-scoring one.

---

## Files in this package

- `SegmentPro.py` — the app (single file, runs directly)
- `requirements.txt` — Python dependencies
- `run_segmentpro.bat` — Windows launcher script (double-click to run from source)
- `README.md` — this file
- `test_headless.py` — non-UI test harness (for CI / verification)
- `sample_output/` — example Excel report and PowerPoint deck generated from the GAP survey

---

*Generated by SegmentPro v1.0*
