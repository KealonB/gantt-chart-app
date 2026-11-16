#!/usr/bin/env node
/**
 * Headless checks for the Gantt schedule engine, Excel I/O, and (optionally) the UI.
 * Usage:
 *   node test/run.mjs
 *   node test/run.mjs --url http://127.0.0.1:8080/
 */
import { createRequire } from "node:module";
import { spawn, spawnSync } from "node:child_process";
import { existsSync } from "node:fs";
import { fileURLToPath } from "node:url";
import path from "node:path";
import process from "node:process";

const require = createRequire(import.meta.url);
const S = require("../schedule.js");
const Excel = require("../excel.js");
const XLSX = require("../vendor/xlsx.bundle.js");

const __dirname = path.dirname(fileURLToPath(import.meta.url));

let failed = 0;
let passed = 0;

function assert(cond, msg) {
  if (cond) {
    passed++;
    console.log("  ok  " + msg);
  } else {
    failed++;
    console.error("  FAIL  " + msg);
  }
}

function eq(a, b, msg) {
  assert(a === b, msg + " (expected " + JSON.stringify(b) + ", got " + JSON.stringify(a) + ")");
}

function task(partial) {
  return Object.assign(
    {
      number: 1,
      name: "",
      start: null,
      end: null,
      duration: null,
      progress: 0,
      dependencies: [],
      indent: 0,
    },
    partial
  );
}

console.log("Schedule engine");

/* start + duration => end */
{
  const tasks = [task({ number: 1, start: "2026-11-15", duration: 5 })];
  S.reschedule(tasks, { today: "2026-11-15" });
  eq(tasks[0].end, "2026-11-19", "start + duration 5 computes inclusive end");
  eq(tasks[0].duration, 5, "duration stays 5");
}

/* start + end => duration */
{
  const tasks = [task({ number: 1, start: "2026-11-01", end: "2026-11-10" })];
  S.reschedule(tasks, { today: "2026-11-15" });
  eq(tasks[0].duration, 10, "start + end computes inclusive duration");
}

/* default start is today when no predecessor */
{
  const tasks = [task({ number: 1, duration: 3 })];
  S.reschedule(tasks, { today: "2026-11-15" });
  eq(tasks[0].start, "2026-11-15", "missing start defaults to today");
  eq(tasks[0].end, "2026-11-17", "end follows from default start + duration");
}

/* FS dependency shifts the dependent task, and cascades */
{
  const tasks = [
    task({ number: 1, name: "A", start: "2026-11-15", duration: 5 }),
    task({ number: 2, name: "B", start: "2026-11-15", duration: 3, dependencies: [1] }),
    task({ number: 3, name: "C", start: "2026-11-15", duration: 2, dependencies: [2] }),
  ];
  S.reschedule(tasks, { today: "2026-11-15" });
  eq(tasks[0].end, "2026-11-19", "predecessor A ends on Nov 19");
  eq(tasks[1].start, "2026-11-20", "B is pushed to the day after A ends");
  eq(tasks[1].end, "2026-11-22", "B keeps a 3-day duration after the shift");
  eq(tasks[2].start, "2026-11-23", "C cascades to the day after B ends");
  eq(tasks[2].end, "2026-11-24", "C keeps a 2-day duration");
}

/* dependent that already starts later keeps slack */
{
  const tasks = [
    task({ number: 1, start: "2026-11-15", duration: 2 }),
    task({ number: 2, start: "2026-11-25", duration: 2, dependencies: [1] }),
  ];
  S.reschedule(tasks, { today: "2026-11-15" });
  eq(tasks[1].start, "2026-11-25", "later start is kept when the FS constraint is already satisfied");
}

/* default start uses latest predecessor end */
{
  const tasks = [
    task({ number: 1, start: "2026-11-10", duration: 5 }),
    task({ number: 2, duration: 3, dependencies: [1] }),
  ];
  S.reschedule(tasks, { today: "2026-11-15" });
  eq(tasks[1].start, "2026-11-15", "missing start defaults to day after predecessor end");
  eq(tasks[1].end, "2026-11-17", "duration still applied after default start");
}

/* groups span children and roll up progress */
{
  const tasks = [
    task({ number: 1, name: "Group", indent: 0 }),
    task({ number: 2, name: "Child A", start: "2026-11-02", duration: 4, progress: 100, indent: 1 }),
    task({ number: 3, name: "Child B", start: "2026-11-04", duration: 7, progress: 0, indent: 1 }),
  ];
  S.reschedule(tasks, { today: "2026-11-15" });
  eq(tasks[0].isGroup, true, "parent is marked as a group");
  eq(tasks[0].start, "2026-11-02", "group start is the earliest child start");
  eq(tasks[0].end, "2026-11-10", "group end is the latest child end");
  eq(tasks[0].duration, 9, "group duration spans first start through last end");
  assert(tasks[0].progress === 36, "group progress is duration-weighted (got " + tasks[0].progress + ")");
}

/* depending on a group waits until the group (all children) finish */
{
  const tasks = [
    task({ number: 1, name: "Group", indent: 0 }),
    task({ number: 2, name: "Child", start: "2026-11-15", duration: 5, indent: 1 }),
    task({ number: 3, name: "After", duration: 2, dependencies: [1] }),
  ];
  S.reschedule(tasks, { today: "2026-11-15" });
  eq(tasks[0].end, "2026-11-19", "group end follows the child");
  eq(tasks[2].start, "2026-11-20", "task depending on a group starts after the group ends");
}

/* shifting a predecessor later also shifts dependents */
{
  const tasks = [
    task({ number: 1, start: "2026-11-15", duration: 3 }),
    task({ number: 2, start: "2026-11-18", duration: 2, dependencies: [1] }),
  ];
  S.reschedule(tasks, { today: "2026-11-15" });
  eq(tasks[1].start, "2026-11-18", "B starts the day after A before the move");
  S.shiftTask(tasks[0], 4);
  tasks[0].duration = S.inclusiveDuration(tasks[0].start, tasks[0].end);
  S.reschedule(tasks, { today: "2026-11-15" });
  eq(tasks[0].start, "2026-11-19", "A moved 4 days later");
  eq(tasks[1].start, "2026-11-22", "B follows A after the move");
}

/* parse dependency input (legacy numbers + typed tokens) */
{
  const ids = S.parseDependencyInput("#2, 4 5;7");
  eq(S.formatDependencies(ids), "2, 4, 5, 7", "legacy parser still accepts #, commas, spaces, and semicolons");
  const typed = S.parseDependencyInput("3FS+2, 4SS-1, 5, #6FF+0");
  eq(S.formatDependencies(typed), "3FS+2, 4SS-1, 5, 6FF", "parser accepts type and lag tokens");
  const migrated = S.normalizeDeps([1, 2]);
  eq(migrated[0].pred, 1, "numeric dependency 1 migrates to pred 1");
  eq(migrated[0].type, "FS", "numeric dependency migrates as FS");
  eq(migrated[0].lag, 0, "numeric dependency migrates with lag 0");
}

/* cycle detection does not hang */
{
  const tasks = [
    task({ number: 1, start: "2026-11-15", duration: 2, dependencies: [2] }),
    task({ number: 2, start: "2026-11-15", duration: 2, dependencies: [1] }),
  ];
  const result = S.reschedule(tasks, { today: "2026-11-15" });
  assert(result.cyclic[1] && result.cyclic[2], "cycle is detected between 1 and 2");
  eq(tasks[0].start, "2026-11-15", "cyclic tasks keep their dates instead of looping");
}

console.log("\nWorkdays and holidays");
{
  const tasks = [task({ number: 1, start: "2026-11-16", duration: 5 })];
  S.reschedule(tasks, { today: "2026-11-16", mode: "workdays" });
  eq(tasks[0].end, "2026-11-20", "5 workdays starting Monday end on Friday");
}
{
  const tasks = [task({ number: 1, start: "2026-11-13", duration: 3 })];
  S.reschedule(tasks, { today: "2026-11-16", mode: "workdays" });
  eq(tasks[0].end, "2026-11-17", "3 workdays starting Friday end on Tuesday");
}
{
  const tasks = [task({ number: 1, start: "2026-11-23", duration: 5 })];
  S.reschedule(tasks, { today: "2026-11-16", mode: "workdays", holidays: ["2026-11-26"] });
  eq(tasks[0].end, "2026-11-30", "Thanksgiving is skipped when counting 5 workdays");
}
{
  const tasks = [
    task({ number: 1, start: "2026-11-16", duration: 5 }),
    task({ number: 2, duration: 2, dependencies: [1] }),
  ];
  S.reschedule(tasks, { today: "2026-11-16", mode: "workdays" });
  eq(tasks[0].end, "2026-11-20", "predecessor ends Friday in workdays mode");
  eq(tasks[1].start, "2026-11-23", "FS successor starts the next Monday, not Saturday");
}
{
  const tasks = [task({ number: 1, start: "2026-11-16", end: "2026-11-20" })];
  S.reschedule(tasks, { today: "2026-11-16", mode: "workdays" });
  eq(tasks[0].duration, 5, "start + end in workdays mode counts only working days");
}

console.log("\nDependency types and lag");
{
  const tasks = [
    task({ number: 1, name: "A", start: "2026-11-16", duration: 3 }),
    task({ number: 2, name: "B", duration: 2, dependencies: "1FS+2" }),
  ];
  S.reschedule(tasks, { today: "2026-11-16" });
  eq(tasks[0].end, "2026-11-18", "A is 3 calendar days, ending Nov 18");
  eq(tasks[1].start, "2026-11-21", "FS+2 starts 2 days after the day after A ends");
}
{
  const tasks = [
    task({ number: 1, name: "A", start: "2026-11-16", duration: 5 }),
    task({ number: 2, name: "B", duration: 2, dependencies: "1SS+2" }),
  ];
  S.reschedule(tasks, { today: "2026-11-16" });
  eq(tasks[1].start, "2026-11-18", "SS+2 starts 2 days after A's start");
}
{
  const tasks = [
    task({ number: 1, name: "A", start: "2026-11-16", duration: 5 }),
    task({ number: 2, name: "B", duration: 2, dependencies: "1FF" }),
  ];
  S.reschedule(tasks, { today: "2026-11-16" });
  eq(tasks[1].end, "2026-11-20", "FF+0 ends no earlier than A");
  eq(tasks[1].start, "2026-11-19", "FF+0 keeps B's 2-day duration");
}
{
  const tasks = [
    task({ number: 1, name: "A", start: "2026-11-20", duration: 3 }),
    task({ number: 2, name: "B", duration: 2, dependencies: "1SF" }),
  ];
  S.reschedule(tasks, { today: "2026-11-16" });
  eq(tasks[1].end, "2026-11-20", "SF+0 ends no earlier than A's start");
  eq(tasks[1].start, "2026-11-19", "SF+0 keeps B's 2-day duration");
}
{
  const tasks = [
    task({ number: 1, name: "A", start: "2026-11-16", duration: 3 }),
    task({ number: 2, name: "B", duration: 2, dependencies: "1FS-1" }),
  ];
  S.reschedule(tasks, { today: "2026-11-16" });
  eq(tasks[1].start, "2026-11-18", "FS-1 (lead) may start on A's finish date");
}

console.log("\nCritical path (CPM) and float");
{
  const tasks = [
    task({ number: 1, name: "A", start: "2026-11-16", duration: 5 }),
    task({ number: 2, name: "B", duration: 3, dependencies: [1] }),
    task({ number: 3, name: "C", duration: 2, dependencies: [1] }),
  ];
  S.reschedule(tasks, { today: "2026-11-16" });
  eq(tasks[0].end, "2026-11-20", "A ends Nov 20");
  eq(tasks[1].end, "2026-11-23", "B ends Nov 23 and sets the project finish");
  eq(tasks[2].end, "2026-11-22", "C ends Nov 22, one day of float");
  assert(tasks[0].critical && tasks[1].critical, "A and B are on the critical path");
  assert(!tasks[2].critical, "C is not critical");
  eq(tasks[2].totalFloat, 1, "C has one day of total float");
  eq(tasks[0].totalFloat, 0, "A has zero total float");
}

console.log("\nMilestones and baseline variance");
{
  const tasks = [
    task({ number: 1, name: "Go live", start: "2026-11-20", duration: 0, milestone: true }),
  ];
  S.reschedule(tasks, { today: "2026-11-16" });
  eq(tasks[0].duration, 0, "milestone duration is 0");
  eq(tasks[0].end, "2026-11-20", "milestone end equals start");
  assert(tasks[0].milestone, "task is marked as a milestone");
}
{
  const tasks = [
    task({
      number: 1,
      start: "2026-11-16",
      duration: 5,
      baselineStart: "2026-11-16",
      baselineEnd: "2026-11-18",
    }),
  ];
  S.reschedule(tasks, { today: "2026-11-16" });
  eq(tasks[0].end, "2026-11-20", "actual end is Nov 20");
  eq(tasks[0].variance, 2, "finish slipped 2 calendar days versus the baseline");
}
{
  const tasks = [
    task({
      number: 1,
      start: "2026-11-16",
      duration: 5,
      baselineStart: "2026-11-16",
      baselineEnd: "2026-11-20",
    }),
  ];
  S.reschedule(tasks, { today: "2026-11-16", mode: "workdays", holidays: ["2026-11-19"] });
  // actual: Mon 16, 17, 18, 20, 23 (skip 19 holiday) ends 23
  eq(tasks[0].end, "2026-11-23", "holiday stretches the actual finish");
  eq(tasks[0].variance, 1, "workday variance vs baseline Friday is 1 (Monday)");
}

console.log("\nExcel export / import round-trip");
{
  const original = {
    projectName: "Harbor test",
    calendarMode: "workdays",
    holidays: ["2026-11-26", "2026-12-25"],
    zoom: "week",
    tasks: [
      task({ number: 1, name: "Kickoff", start: "2026-11-02", duration: 0, milestone: true, assignee: "Alex", color: "#0f766e" }),
      task({
        number: 2,
        name: "Build",
        duration: 5,
        assignee: "Quinn",
        allocation: 80,
        dependencies: "1FS+1",
        indent: 0,
        baselineStart: "2026-11-03",
        baselineEnd: "2026-11-09",
        color: "#0369a1",
      }),
      task({ number: 3, name: "Design", duration: 4, dependencies: "1SS+2", assignee: "Sam", allocation: 100, color: "#7c3aed" }),
    ],
  };
  S.reschedule(original.tasks, { today: "2026-11-16", mode: "workdays", holidays: original.holidays });
  const buf = Excel.writeNodeBuffer(original, XLSX, S);
  assert(buf && buf.length > 100, "xlsx buffer is a non-empty file (" + (buf && buf.length) + " bytes)");
  // ZIP/xlsx magic
  assert(buf[0] === 0x50 && buf[1] === 0x4b, "xlsx buffer starts with a ZIP signature");
  const parsed = Excel.parseWorkbook(buf, XLSX, S);
  eq(parsed.projectName, "Harbor test", "round-trip keeps the project name");
  eq(parsed.calendarMode, "workdays", "round-trip keeps workdays mode");
  eq(parsed.holidays.join(","), "2026-11-26,2026-12-25", "round-trip keeps holidays");
  eq(parsed.tasks.length, 3, "round-trip keeps all tasks");
  const kick = parsed.tasks[0];
  eq(kick.milestone, true, "round-trip keeps milestone flag");
  eq(kick.start, original.tasks[0].start, "round-trip keeps milestone date " + kick.start);
  const build = parsed.tasks[1];
  eq(S.formatDependencies(build.dependencies), "1FS+1", "round-trip keeps typed lag");
  eq(build.allocation, 80, "round-trip keeps allocation %");
  eq(build.baselineEnd, "2026-11-09", "round-trip keeps baseline end");
  const design = parsed.tasks[2];
  eq(S.formatDependencies(design.dependencies), "1SS+2", "round-trip keeps SS+2");

  const csv = "Name,Start,Duration,Dependencies,Assignee\nAlpha,2026-11-16,3,,Sam\nBeta,2026-11-16,2,1FS+2,Alex\n";
  const fromCsv = Excel.parseCSV(csv, S);
  eq(fromCsv.tasks.length, 2, "CSV import reads two rows");
  eq(fromCsv.tasks[0].name, "Alpha", "CSV import maps the Name column");
  eq(fromCsv.tasks[1].dependencies[0].type, "FS", "CSV import parses a typed dependency");
  eq(fromCsv.tasks[1].dependencies[0].lag, 2, "CSV import parses lag");
}

console.log("\nEngine results: " + passed + " passed, " + failed + " failed");

function chromePath() {
  if (process.env.CHROME_PATH) return process.env.CHROME_PATH;
  const candidates = [
    "google-chrome",
    "google-chrome-stable",
    "chromium",
    "chromium-browser",
    "/opt/google/chrome/google-chrome",
    "/usr/bin/google-chrome",
    "/usr/bin/google-chrome-stable",
  ];
  for (const c of candidates) {
    if (c.startsWith("/") && existsSync(c)) return c;
    const which = spawnSync("which", [c], { encoding: "utf8" });
    if (which.status === 0 && which.stdout.trim()) return which.stdout.trim();
  }
  return "google-chrome";
}

async function runBrowser(url) {
  console.log("\nBrowser UI @ " + url);

  let puppeteer;
  try {
    puppeteer = require("/tmp/gantt-test/node_modules/puppeteer-core");
  } catch (err) {
    console.log("Installing puppeteer-core in /tmp for the headless check...");
    await new Promise((resolve, reject) => {
      const npm = spawn("npm", ["install", "--prefix", "/tmp/gantt-test", "puppeteer-core@23"], {
        stdio: "inherit",
      });
      npm.on("exit", (code) => (code === 0 ? resolve() : reject(new Error("npm install failed: " + code))));
    });
    puppeteer = require("/tmp/gantt-test/node_modules/puppeteer-core");
  }

  const browser = await puppeteer.launch({
    executablePath: chromePath(),
    headless: "new",
    args: ["--no-sandbox", "--disable-gpu", "--disable-dev-shm-usage"],
  });

  const page = await browser.newPage();
  const errors = [];
  page.on("pageerror", (err) => errors.push(String(err)));
  page.on("console", (msg) => {
    if (msg.type() === "error") errors.push("console.error: " + msg.text());
  });

  const res = await page.goto(url, { waitUntil: "networkidle0", timeout: 30000 });
  assert(res && res.ok(), "app page returns " + (res && res.status()));

  const boot = await page.evaluate(() => ({
    hasApp: typeof window.GanttApp === "object",
    hasSchedule: typeof window.GanttSchedule === "object",
    hasExcel: typeof window.GanttExcel === "object",
    hasXLSX: typeof window.XLSX === "object",
    version: document.querySelector('meta[name="gantt-app-version"]')?.content,
  }));
  assert(boot.hasApp, "window.GanttApp is exposed");
  assert(boot.hasSchedule, "window.GanttSchedule is exposed");
  assert(boot.hasExcel && boot.hasXLSX, "Excel library is loaded in the page");
  eq(boot.version, "2", "page reports app version 2");

  await page.evaluate(() => window.GanttApp.newProject());

  const added = await page.evaluate(() => {
    const t = window.GanttApp.addTaskSilent({
      name: "Design",
      start: "2026-11-15",
      duration: 5,
    });
    const bar = document.querySelector('.bar[data-id="' + t.id + '"]');
    const state = window.GanttApp.getState();
    const task = state.tasks[0];
    return {
      id: t.id,
      start: task.start,
      end: task.end,
      duration: task.duration,
      hasBar: !!bar,
      barWidth: bar ? bar.getBoundingClientRect().width : 0,
      rowCount: document.querySelectorAll("#task-body tr").length,
    };
  });

  eq(added.start, "2026-11-15", "UI task keeps the entered start date");
  eq(added.end, "2026-11-19", "UI start + duration fills the end date");
  eq(added.duration, 5, "UI duration is 5");
  assert(added.hasBar, "a timeline bar is rendered for the new task");
  assert(added.barWidth > 8, "the rendered bar has a visible width (" + added.barWidth + "px)");
  eq(added.rowCount, 1, "the task table has one row");

  const deps = await page.evaluate(() => {
    const a = window.GanttApp.getState().tasks[0];
    window.GanttApp.addTaskSilent({
      name: "Build",
      start: "2026-11-15",
      duration: 3,
      dependencies: [a.number],
    });
    const state = window.GanttApp.getState();
    const b = state.tasks[1];
    const barB = document.querySelector('.bar[data-id="' + b.id + '"]');
    return {
      aEnd: state.tasks[0].end,
      bStart: b.start,
      bEnd: b.end,
      hasBar: !!barB,
      arrows: document.querySelectorAll(".arrow-layer path").length,
    };
  });

  eq(deps.aEnd, "2026-11-19", "predecessor still ends Nov 19");
  eq(deps.bStart, "2026-11-20", "UI dependency shift moves the dependent start");
  eq(deps.bEnd, "2026-11-22", "UI dependency shift keeps duration and fills the new end");
  assert(deps.hasBar, "dependent task has a timeline bar");
  assert(deps.arrows >= 1, "a dependency arrow is drawn");

  const sampleUi = await page.evaluate(() => {
    window.GanttApp.loadSample();
    window.GanttApp.setShowCritical(true);
    const state = window.GanttApp.getState();
    return {
      name: state.projectName,
      tasks: state.tasks.length,
      mode: state.calendarMode,
      holidays: state.holidays.length,
      diamonds: document.querySelectorAll(".milestone-diamond").length,
      criticalBars: document.querySelectorAll(".bar.is-critical, .milestone-diamond.is-critical").length,
      baselines: document.querySelectorAll(".bar-baseline").length,
    };
  });
  eq(sampleUi.name, "Harbor site launch", "sample project loads");
  eq(sampleUi.tasks, 16, "sample has 16 tasks");
  eq(sampleUi.mode, "workdays", "sample uses workdays scheduling");
  assert(sampleUi.holidays >= 2, "sample has holidays");
  assert(sampleUi.diamonds >= 2, "milestone diamonds render (" + sampleUi.diamonds + ")");
  assert(sampleUi.criticalBars >= 2, "critical bars or diamonds are highlighted (" + sampleUi.criticalBars + ")");
  assert(sampleUi.baselines >= 1, "baseline ghost bars render (" + sampleUi.baselines + ")");

  const resources = await page.evaluate(() => {
    window.GanttApp.setView("resources");
    const rows = document.querySelectorAll(".res-row");
    const over = document.querySelectorAll(".res-label.is-over, .res-cell.is-over");
    return {
      view: window.GanttApp.getState().view,
      rows: rows.length,
      over: over.length,
      canvasShown: !document.getElementById("resource-canvas").hidden,
    };
  });
  eq(resources.view, "resources", "resource view is selected");
  assert(resources.rows >= 3, "resource view renders assignee rows (" + resources.rows + ")");
  assert(resources.over >= 1, "overallocation is flagged in the resource view");
  assert(resources.canvasShown, "resource canvas is visible");

  const xlsx = await page.evaluate(() => {
    const buf = window.GanttApp.exportExcelBuffer();
    const bytes = Array.from(new Uint8Array(buf));
    return { length: bytes.length, magic: bytes.slice(0, 2) };
  });
  assert(xlsx.length > 1000, "UI Excel export returns a substantial file (" + xlsx.length + " bytes)");
  assert(xlsx.magic[0] === 0x50 && xlsx.magic[1] === 0x4b, "UI Excel export is a valid xlsx/ZIP file");

  // Parse that buffer in Node to confirm it is a real workbook.
  const pageBuf = await page.evaluate(() => Array.from(new Uint8Array(window.GanttApp.exportExcelBuffer())));
  const wb = XLSX.read(Buffer.from(pageBuf), { type: "buffer" });
  assert(wb.SheetNames.indexOf("Tasks") >= 0 && wb.SheetNames.indexOf("Timeline") >= 0, "exported workbook has Tasks and Timeline sheets (" + wb.SheetNames.join(", ") + ")");

  assert(errors.length === 0, "no page or console errors" + (errors.length ? ": " + errors.join(" | ") : ""));

  await browser.close();
}

const urlFlag = process.argv.indexOf("--url");
const url = urlFlag >= 0 ? process.argv[urlFlag + 1] : process.env.TEST_URL;

if (url) {
  try {
    await runBrowser(url);
  } catch (err) {
    failed++;
    console.error("  FAIL  browser check: " + (err.stack || err));
  }
} else {
  console.log("\nSkipping browser UI check (pass --url http://127.0.0.1:8080/).");
}

console.log("\nTotal: " + passed + " passed, " + failed + " failed");
process.exit(failed ? 1 : 0);
