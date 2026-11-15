#!/usr/bin/env node
/**
 * Headless checks for the Gantt schedule engine and (optionally) the UI.
 * Usage:
 *   node test/run.mjs
 *   node test/run.mjs --url http://127.0.0.1:8080/
 */
import { createRequire } from "node:module";
import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";
import path from "node:path";
import process from "node:process";

const require = createRequire(import.meta.url);
const S = require("../schedule.js");

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, "..");

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

/* parse dependency input */
{
  const ids = S.parseDependencyInput("#2, 4 5;7");
  assert(ids.join(",") === "2,4,5,7", "dependency parser accepts #, commas, spaces, and semicolons: " + ids);
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
  const { spawnSync } = require("node:child_process");
  const { existsSync } = require("node:fs");
  for (const c of candidates) {
    if (c.startsWith("/") && existsSync(c)) return c;
    const which = spawnSync("which", [c], { encoding: "utf8" });
    if (which.status === 0 && which.stdout.trim()) return which.stdout.trim();
  }
  return "google-chrome";
}

async function cdp(wsUrl, method, params, sessionId) {
  const WS = (await import("ws")).default;
  // we'll use a raw fetch-based HTTP /json and a tiny WS client if ws is available
  // fallback implemented below without the ws package
  void WS;
  void wsUrl;
  void method;
  void params;
  void sessionId;
}

async function runBrowser(url) {
  console.log("\nBrowser UI @ " + url);

  const puppeteerMod = await import("puppeteer-core").catch(() => null);
  if (!puppeteerMod) {
    console.log("Installing puppeteer-core in /tmp for the headless check...");
    await new Promise((resolve, reject) => {
      const npm = spawn("npm", ["install", "--prefix", "/tmp/gantt-test", "puppeteer-core@23"], {
        stdio: "inherit",
      });
      npm.on("exit", (code) => (code === 0 ? resolve() : reject(new Error("npm install failed: " + code))));
    });
  }

  const puppeteer = puppeteerMod || createRequire(import.meta.url)("/tmp/gantt-test/node_modules/puppeteer-core");

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
  }));
  assert(boot.hasApp, "window.GanttApp is exposed");
  assert(boot.hasSchedule, "window.GanttSchedule is exposed");

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
    console.error("  FAIL  browser check: " + err.stack || err);
  }
} else {
  console.log("\nSkipping browser UI check (pass --url http://127.0.0.1:8080/).");
}

console.log("\nTotal: " + passed + " passed, " + failed + " failed");
process.exit(failed ? 1 : 0);
