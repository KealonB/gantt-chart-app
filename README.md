# Gantt

A static, single-page Gantt chart creator you can run entirely in the browser. There is no backend and no build step, so the repo root can be served as-is — including from GitHub Pages.

This is a **prerelease** for early testing.

## Features

- Spreadsheet-like task list: name, start date, duration (calendar days), end date, progress, finish-to-start dependencies (by task ID), assignee, and color
- Add, delete, and reorder rows (drag the handle, or use up/down buttons)
- Indent / outdent to build group and subtask outlines; group rows span their children and roll up progress
- **Live autofill:** entering a start and duration fills the end date (and the timeline bar); entering a start and end fills the duration. Changes appear as you type.
- Dependencies automatically push dependent tasks so they start the day after their predecessor ends, and the shift cascades
- If a task has no start date, it defaults to the day after the latest predecessor, or today
- Timeline with day, week, and month zoom, weekend shading, a today line, and dependency arrows
- Drag bars to move tasks; drag either edge to resize. Group bars move their children. The table stays in sync.
- Autosave to `localStorage` in this browser
- Export / import JSON, export a PNG of the chart
- Load a sample project, or start a new one
- Light and dark themes, plus undo / redo

## Usage

1. Click **Add task** (or load the sample project).
2. Type a name, then set a start date and a duration — the end date and the bar on the right fill in immediately.
3. Point **Depends** at earlier task IDs (`4` or `2, 4`). The dependent task will not start until those predecessors finish.
4. Use the indent button to nest a row under the one above it; the parent becomes a group bar.
5. Drag a bar to shift its dates, or drag the left/right edge to change duration.
6. Your work is saved automatically in this browser. Use **Export JSON** for a portable backup, or **Export PNG** for a snapshot.

### Keyboard

- `Ctrl+Z` / `⌘Z` — undo
- `Ctrl+Shift+Z` or `Ctrl+Y` — redo

## Run locally

Serve the repo root with any static file server, for example:

```bash
python3 -m http.server 8080
```

Then open `http://localhost:8080`. You can also open `index.html` as a file, but a local server is the more reliable option.

## GitHub Pages

The site is published from the `main` branch, from the repository root (`index.html` at `/`). An empty `.nojekyll` file is included so GitHub Pages will not process the site with Jekyll.

Live prerelease (once Pages is enabled):

https://kealonb.github.io/gantt-chart-app/

## Project files

| File | Role |
| --- | --- |
| `index.html` | App shell |
| `styles.css` | Layout, light/dark theme |
| `schedule.js` | Date math, dependency shifting, group rollups |
| `app.js` | UI, persistence, import/export |
| `test/run.mjs` | Headless unit and browser checks |

## Known limitations

- Only finish-to-start dependencies are supported (no start-to-start, lag, or lead).
- Duration is in **calendar days**, including weekends. Weekend shading is visual only.
- Circular dependencies are ignored until you break the cycle (a warning appears in the status bar).
- Data lives in this browser’s `localStorage` only — it is not synced across devices unless you export/import JSON.
- There is no multi-user editing, critical-path analysis, or resource leveling.
- Group start, end, duration, and progress are computed from children and are not edited directly (drag a group bar to move the whole block).
- Prerelease: expect rough edges; please file issues when something feels off.
