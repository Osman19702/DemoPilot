# DemoPilot Deck Generator

DemoPilot converts Jira-style CSV exports into polished PI demo decks (PowerPoint) in minutes. It blends heuristic and ML-inspired scoring with a design-forward layout system so product or engineering teams can showcase the most demo-worthy work without spending hours building slides by hand.

## What You Get

- Automated deck assembly: title, KPIs, multi-page PI roadmap, deep-dive slides, and optional appendix.
- Smart ranking that looks at story points, keywords, user impact language, UI cues, and priority to surface the best demo candidates.
- Adaptive design system with modern typography, color psychology, text autoshrink, and responsive card layout.
- Continuous numbering and pagination on the PI Demo Roadmap when more than five items are included.
- Two entry points: a command-line interface and a lightweight Flask web UI for drag-and-drop CSV uploads.

## Project Layout

- `make_deck.py` — core engine that scores issues and builds PowerPoint slides.
- `webapp/` — Flask frontend (`app.py`, templates, static assets) to generate decks through a browser.
- `requirements.txt` — consolidated dependency list for both the CLI and web UI.
- `*.csv` — sample Jira exports for testing.

## Prerequisites

- Python 3.10 or newer.
- PowerPoint or compatible viewer to open generated `.pptx` files.
- Windows PowerShell commands shown below; adjust activation paths if you use a different shell.

### Create and Activate a Virtual Environment

```powershell
python -m venv .venv
\.venv\Scripts\Activate
```

### Install Dependencies

```powershell
pip install -r requirements.txt
```

## Command-Line Usage

```powershell
\.venv\Scripts\python.exe make_deck.py --csv <input.csv> --out <output.pptx>
```

### Common Flags

- `--csv` (required): Jira export CSV path.
- `--out` (required): Output PPTX filename.
- `--pi` / `--sprint`: Filter to a specific PI name.
- `--pi-number`: Override the PI label shown on slides (e.g., `28` -> `PI 28`).
- `--top`: Limit the number of ranked items (default 8).
- `--template`: Apply a `.potx`/`.pptx` brand template.
- `--title`: Custom deck title.
- `--include-appendix`: Append a table summarizing all items.

### Example

```powershell
\.venv\Scripts\python.exe make_deck.py `
  --csv demopilot_detection_ready_6sprints.csv `
  --pi "PI 6" `
  --pi-number 28 `
  --top 8 `
  --template BrandTheme.potx `
  --out DemoPilot_PI6.pptx `
  --include-appendix
```

## Web UI Usage

The Flask frontend offers a quick way to upload CSV files, adjust top-N, and download the generated deck.

```powershell
\.venv\Scripts\python.exe webapp/app.py
```

Navigate to http://127.0.0.1:5000 and follow the form prompts:

1. Choose a Jira CSV export.
2. Provide an output file name (`.pptx` extension is added automatically if omitted).
3. Set the `Top-N` limit and optional title.
4. Submit to download the generated presentation.

Stop the development server with `Ctrl+C` when you are finished.

## CSV Expectations

DemoPilot expects Jira-like columns but fails gracefully when data is missing:

- `Issue key` (generated automatically if absent)
- `Summary`
- `Issue Type`
- `Priority`
- `Status`
- `Story Points`
- `Description`
- Optional: `Assignee`, `Components`, `PI` / `Sprint`, `Updated`

Missing numeric fields default to zero, and text fields are blank-safe to prevent crashes. Filtering by PI or sprint is only applied when matching columns exist.

## How Scoring Works

- Story points provide a base weight (higher SP = more impactful).
- Priority and issue type adjust the score; stories with user impact language get a boost.
- Keyword detectors look for UI, customer-facing terms, performance improvements, and metric callouts.
- Items with cancelled or rejected statuses are excluded automatically.
- Scores are rounded for display and printed to stdout when running the CLI, helping you audit ranking decisions.

## Slide Breakdown

1. **Title** — Branded hero slide with optional template integration.
2. **KPIs** — Cards summarizing total story points, item count, and averages, plus breakdowns by issue type and assignee/component.
3. **PI Demo Roadmap** — Automatically paginated overview; up to five items per slide with continuous numbering, color-coded priority strips, and issue metadata.
4. **Item Detail Slides** — One per ranked item, including why-it-matters callouts, metadata cards, and bullet highlights from the description.
5. **Appendix** (optional) — Table view of the export with alternating row shading and priority/type color accents.

### Text Fitting

Every slide passes through a text autoshrink routine to avoid overflow. Font sizes adjust within sensible bounds, preserving readability while fitting the available space.

## Troubleshooting

- **ModuleNotFoundError** — Ensure the virtual environment is activated and dependencies are installed (`pip install -r requirements.txt`).
- **File permission errors on Windows** — Close the target PPTX before re-running the generator.
- **No items selected** — Check that the CSV has valid statuses and story points; the script prints a list of fallback candidates if the demo-worthiness threshold filters everything out.
- **Flask server exits immediately** — Confirm no other process is using port 5000, or export `FLASK_RUN_PORT=<port>` and update the command accordingly.

## Extending DemoPilot

- Adjust scoring weights in `make_deck.py` to tune what qualifies as demo-worthy.
- Customize colors and typography via the `DesignTheme` class.
- Add new CSV-derived metrics or slide types by extending `add_kpi_slide`, `add_overview_slide`, or `add_item_slide`.

## License

Provided as-is under an internal-use license. Modify and integrate it into your PI demo workflow as needed.
