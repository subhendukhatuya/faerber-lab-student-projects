# Färber Lab – Student Projects (WS 2026/27)

Live page: https://subhendukhatuya.github.io/faerber-lab-student-projects/

## Updating topics
1. Edit `Student_Projects.xlsx` (one row per topic; set **Availability** to `Open` or `Filled`) and place it next to `build.py`.
2. Run `pip install openpyxl && python build.py` – this regenerates `projects.json`.
3. Commit and push. GitHub Pages updates within a minute.

Small changes (e.g. marking a topic as filled) can also be made directly in `projects.json` on GitHub.
Dates, the deadline and the registration-form link live in the `CONFIG` block at the bottom of `index.html`.

`fallback_content.json` holds text for rows whose spreadsheet cells are empty or say "same content as the old portal".
