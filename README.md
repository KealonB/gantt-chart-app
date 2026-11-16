# Gantt

A static, single-page Gantt chart creator you can run entirely in the browser. There is no backend and no build step, so the repo root can be served as-is — including from GitHub Pages.

This is a **prerelease** for early testing.

## Features

- Spreadsheet-like task list: name, start, duration, end, progress, dependencies, assignee, allocation %, milestone, variance, and color
- Add, delete, and reorder rows (drag the handle, or use up/down buttons)
- Indent / outdent to build group and subtask outlines; group rows span their children and roll up progress
- **Live autofill:** entering a start and duration fills the end date (and the timeline bar); entering a start and end fills the duration. Changes appear as you type.
- Dependencies automatically push dependent tasks, and the shift cascades. If a task has no start date, it defaults to the day after the latest predecessor, or today.
- **Dependency types** FS, SS, FF, and SF, with lag or lead. Type them in the Depends cell, for example `3`, `4FS+2`, `5SS-1`, `6FF`, or `7SF`.
- **Workdays-only scheduling** (optional): durations skip weekends and an editable holiday list. Calendar-days mode remains the default, and existing projects keep their original dates until you switch.
- **Critical path:** a CPM forward/backward pass computes total float and highlights critical bars and arrows (toggle it on from the toolbar).
- **Milestones** are zero-duration tasks drawn as diamonds.
- **Baselines:** Set baseline saves a snapshot of planned dates. Ghost bars show the plan under the actual bars, and the Var column shows finish slip (or pull-in) in the current day unit. Clear baseline removes the snapshot.
- Timeline with day, week, and month zoom, a today line, weekend shading, and dependency arrows. Drag bars to move them, or drag either edge to resize.
- **Resource workload view:** a heat strip per assignee, per day, with overallocation flagged when a person is over 100% (allocation % per task is optional; it defaults to 100).
- Autosave to `localStorage` in this browser (older saved projects are migrated automatically)
- Export / import JSON, **Excel (.xlsx)** (a formatted Tasks sheet plus a cell-colored Timeline sheet), CSV import, and PNG export
- Load a sample project, or start a new one
- Light and dark themes, plus undo / redo

## Usage

1. Click **Add task** (or load the sample project).
2. Type a name, then set a start date and a duration — the end date and the bar on the right fill in immediately.
3. Point **Depends** at earlier task IDs. Plain `4` is finish-to-start with no lag. `4FS+2` waits two extra days; `5SS-1` is start-to-start with one day of lead.
4. Use the indent button to nest a row under the one above it; the parent becomes a group bar.
5. Check **MS** to make a milestone, or set duration to 0.
6. Switch **Workdays** if duration should skip weekends. Open **Holidays** to add dates such as Thanksgiving.
7. Turn on **Critical path** to see the chain that sets the finish date. **Set baseline** once a plan is agreed, then watch **Var** as dates move.
8. Open the **Resources** view to see each assignee's load and any days over 100%.
9. Your work is saved automatically in this browser. Use **Export JSON** or **Export Excel** for a portable backup, or **Export PNG** for a snapshot.

### Keyboard

- `Ctrl+Z` / `⌘Z` — undo
- `Ctrl+Shift+Z` or `Ctrl+Y` — redo

## Run locally

Serve the repo root with any static file server, for example:

```bash
python3 -m http.server 8080
```

Then open `http://localhost:8080`. You can also open `index.html` as a file, but a local server is the more reliable option.

Headless checks:

```bash
node test/run.mjs
python3 -m http.server 8080
node test/run.mjs --url http://127.0.0.1:8080/
```

## GitHub Pages

The site is published from the `main` branch, from the repository root (`index.html` at `/`). An empty `.nojekyll` file is included so GitHub Pages will not process the site with Jekyll.

Live prerelease:

https://kealonb.github.io/gantt-chart-app/

## Project files

| File | Role |
| --- | --- |
| `index.html` | App shell |
| `styles.css` | Layout, light/dark theme |
| `schedule.js` | Date math, calendars, dependencies, CPM, group rollups |
| `excel.js` | Excel and CSV import/export |
| `app.js` | UI, persistence, views |
| `vendor/xlsx.bundle.js` | SheetJS-compatible library with cell styling (xlsx-js-style 1.2.0) |
| `test/run.mjs` | Headless unit and browser checks |

## Known limitations

- Duration in workdays mode counts weekdays minus holidays; remaining calendar shading is visual only.
- Switching between calendar days and workdays keeps each task's start and duration, then recomputes the end in the new unit.
- Lag and lead are whole days in the current calendar (no hours).
- Critical path is computed from the current plan and dependency network. Group deadlines are pushed down to children; unusual networks can still look noisy.
- Variance is finish variance (actual end minus baseline end) in the current day unit. A positive number means the task finished later than planned.
- Overallocation is flagged when the sum of allocation % on a day is greater than 100. Unassigned tasks are omitted from the resource view.
- Excel Timeline coloring uses one column per day (or per week if the span is long). Very long projects make a wide sheet.
- Data lives in this browser's `localStorage` only — it is not synced across devices unless you export/import.
- There is no multi-user editing or resource leveling (the resource view is diagnostic only).
- Prerelease: expect rough edges; please file issues when something feels off.
