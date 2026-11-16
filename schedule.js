/**
 * Gantt scheduling engine (dates, calendars, dependencies, CPM, groups).
 * Works in the browser and in Node for tests.
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory();
  } else {
    root.GanttSchedule = factory();
  }
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  var DAY_MS = 86400000;
  var DEP_TYPES = { FS: true, SS: true, FF: true, SF: true };

  function todayISO() {
    return toISO(new Date());
  }

  function toISO(d) {
    var y = d.getFullYear();
    var m = String(d.getMonth() + 1).padStart(2, "0");
    var day = String(d.getDate()).padStart(2, "0");
    return y + "-" + m + "-" + day;
  }

  function parseISO(s) {
    if (!s || typeof s !== "string") return null;
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s.trim());
    if (!m) return null;
    var y = Number(m[1]);
    var mo = Number(m[2]);
    var da = Number(m[3]);
    var d = new Date(y, mo - 1, da);
    if (d.getFullYear() !== y || d.getMonth() !== mo - 1 || d.getDate() !== da) return null;
    return d;
  }

  function isISO(s) {
    return parseISO(s) != null;
  }

  function addDays(iso, n) {
    var d = parseISO(iso);
    if (!d) return null;
    d.setDate(d.getDate() + Number(n));
    return toISO(d);
  }

  function diffDays(a, b) {
    var da = parseISO(a);
    var db = parseISO(b);
    if (!da || !db) return null;
    return Math.round((db.getTime() - da.getTime()) / DAY_MS);
  }

  function maxISO() {
    var valid = Array.prototype.filter.call(arguments, isISO);
    if (!valid.length) return null;
    return valid.reduce(function (m, d) {
      return d > m ? d : m;
    });
  }

  function minISO() {
    var valid = Array.prototype.filter.call(arguments, isISO);
    if (!valid.length) return null;
    return valid.reduce(function (m, d) {
      return d < m ? d : m;
    });
  }

  function uniqueNums(arr) {
    var out = [];
    var seen = {};
    (arr || []).forEach(function (n) {
      var v = Number(n);
      if (!Number.isFinite(v) || v <= 0) return;
      v = Math.round(v);
      if (!seen[v]) {
        seen[v] = true;
        out.push(v);
      }
    });
    return out;
  }

  function uniqueSortedDates(arr) {
    var out = [];
    var seen = {};
    (arr || []).forEach(function (s) {
      if (!isISO(s)) return;
      if (!seen[s]) {
        seen[s] = true;
        out.push(s);
      }
    });
    out.sort();
    return out;
  }

  function isWeekend(iso) {
    var d = parseISO(iso);
    if (!d) return false;
    var day = d.getDay();
    return day === 0 || day === 6;
  }

  function makeCalendar(mode, holidays) {
    var holidayMap = {};
    uniqueSortedDates(holidays).forEach(function (d) {
      holidayMap[d] = true;
    });
    return {
      mode: mode === "workdays" ? "workdays" : "calendar",
      holidays: uniqueSortedDates(holidays),
      holidayMap: holidayMap,
    };
  }

  function isWorking(iso, cal) {
    if (!isISO(iso)) return false;
    if (!cal || cal.mode !== "workdays") return true;
    if (isWeekend(iso)) return false;
    if (cal.holidayMap && cal.holidayMap[iso]) return false;
    return true;
  }

  function addWorkdays(iso, n, cal) {
    if (!isISO(iso)) return null;
    n = Math.round(Number(n) || 0);
    if (!n) return iso;
    var step = n > 0 ? 1 : -1;
    var left = Math.abs(n);
    var cur = iso;
    var guard = 0;
    while (left > 0 && guard++ < 20000) {
      cur = addDays(cur, step);
      if (isWorking(cur, cal)) left--;
    }
    return cur;
  }

  function snapToWorkday(iso, cal, direction) {
    if (!isISO(iso)) return iso;
    if (!cal || cal.mode !== "workdays" || isWorking(iso, cal)) return iso;
    var step = direction < 0 ? -1 : 1;
    var cur = iso;
    for (var i = 0; i < 732; i++) {
      cur = addDays(cur, step);
      if (isWorking(cur, cal)) return cur;
    }
    return iso;
  }

  function addLag(iso, lag, cal) {
    if (!isISO(iso)) return null;
    lag = Math.round(Number(lag) || 0);
    if (!lag) return iso;
    if (cal && cal.mode === "workdays") return addWorkdays(iso, lag, cal);
    return addDays(iso, lag);
  }

  function nextUnit(iso, cal) {
    if (cal && cal.mode === "workdays") return addWorkdays(iso, 1, cal);
    return addDays(iso, 1);
  }

  function prevUnit(iso, cal) {
    if (cal && cal.mode === "workdays") return addWorkdays(iso, -1, cal);
    return addDays(iso, -1);
  }

  function shiftOffset(from, to, cal) {
    if (!isISO(from) || !isISO(to)) return 0;
    if (cal && cal.mode === "workdays") return workingDiff(from, to, cal);
    return diffDays(from, to);
  }

  function workingDiff(from, to, cal) {
    if (from === to) return 0;
    var sign = from < to ? 1 : -1;
    var a = sign > 0 ? from : to;
    var b = sign > 0 ? to : from;
    var n = 0;
    var cur = a;
    while (cur < b) {
      cur = addDays(cur, 1);
      if (isWorking(cur, cal)) n++;
    }
    return n * sign;
  }

  function endFromStartDuration(start, duration, cal) {
    if (!isISO(start)) return null;
    duration = Math.round(Number(duration) || 0);
    if (duration <= 0) return start;
    if (cal && cal.mode === "workdays") {
      start = snapToWorkday(start, cal, 1);
      return addWorkdays(start, duration - 1, cal);
    }
    return addDays(start, duration - 1);
  }

  function startFromEndDuration(end, duration, cal) {
    if (!isISO(end)) return null;
    duration = Math.round(Number(duration) || 0);
    if (duration <= 0) return end;
    if (cal && cal.mode === "workdays") {
      end = snapToWorkday(end, cal, -1);
      return addWorkdays(end, -(duration - 1), cal);
    }
    return addDays(end, -(duration - 1));
  }

  function inclusiveDuration(start, end, cal) {
    if (!isISO(start) || !isISO(end)) return null;
    if (end < start) return 0;
    if (cal && cal.mode === "workdays") {
      var n = 0;
      eachDay(start, end, function (iso) {
        if (isWorking(iso, cal)) n++;
      });
      return n;
    }
    return diffDays(start, end) + 1;
  }

  /* ---- dependencies ----------------------------------------------------- */

  function normalizeDep(raw) {
    if (raw == null || raw === "") return null;
    if (typeof raw === "number" || (typeof raw === "string" && /^\s*#?\d+\s*$/.test(raw))) {
      var n0 = Math.round(Number(String(raw).replace("#", "").trim()));
      if (!Number.isFinite(n0) || n0 <= 0) return null;
      return { pred: n0, type: "FS", lag: 0 };
    }
    if (typeof raw === "object") {
      var p = Math.round(Number(raw.pred != null ? raw.pred : raw.id != null ? raw.id : raw.number));
      if (!Number.isFinite(p) || p <= 0) return null;
      var type = String(raw.type || "FS").toUpperCase();
      if (!DEP_TYPES[type]) type = "FS";
      var lag = Math.round(Number(raw.lag) || 0);
      return { pred: p, type: type, lag: lag };
    }
    return null;
  }

  function parseOneDep(token) {
    token = String(token || "").trim();
    if (!token) return null;
    var m = /^#?(\d+)\s*(FS|SS|FF|SF)?\s*([+-]\d+)?$/i.exec(token);
    if (!m) return null;
    var pred = Number(m[1]);
    var type = m[2] ? m[2].toUpperCase() : "FS";
    var lag = m[3] ? Number(m[3]) : 0;
    return { pred: pred, type: type, lag: lag };
  }

  function parseDependencyInput(text) {
    if (text == null || text === "") return [];
    if (Array.isArray(text)) return normalizeDeps(text);
    var out = [];
    String(text).replace(/#?\d+\s*(?:FS|SS|FF|SF)?\s*(?:[+-]\d+)?/gi, function (token) {
      var one = parseOneDep(token);
      if (one) out.push(one);
      return token;
    });
    return uniqueDeps(out);
  }

  function uniqueDeps(arr) {
    var out = [];
    var seen = {};
    (arr || []).forEach(function (raw) {
      var d = normalizeDep(raw);
      if (!d) return;
      var key = d.pred + ":" + d.type + ":" + d.lag;
      if (seen[key]) return;
      seen[key] = true;
      out.push(d);
    });
    return out;
  }

  function normalizeDeps(arr) {
    if (typeof arr === "string") return parseDependencyInput(arr);
    return uniqueDeps(arr || []);
  }

  function formatDep(dep) {
    var d = normalizeDep(dep);
    if (!d) return "";
    var s = String(d.pred);
    if (d.type !== "FS" || d.lag) s += d.type;
    if (d.lag) s += (d.lag > 0 ? "+" : "") + d.lag;
    return s;
  }

  function formatDependencies(deps) {
    return normalizeDeps(deps)
      .map(formatDep)
      .filter(Boolean)
      .join(", ");
  }

  function depPreds(deps) {
    return normalizeDeps(deps).map(function (d) {
      return d.pred;
    });
  }

  /* ---- outline helpers -------------------------------------------------- */

  function findTaskByNumber(tasks, num) {
    var n = Number(num);
    for (var i = 0; i < tasks.length; i++) {
      if (tasks[i].number === n) return tasks[i];
    }
    return null;
  }

  function findIndexByNumber(tasks, num) {
    var n = Number(num);
    for (var i = 0; i < tasks.length; i++) {
      if (tasks[i].number === n) return i;
    }
    return -1;
  }

  function descendantRange(tasks, parentIndex) {
    var indent = tasks[parentIndex].indent || 0;
    var end = parentIndex;
    for (var i = parentIndex + 1; i < tasks.length; i++) {
      if ((tasks[i].indent || 0) <= indent) break;
      end = i;
    }
    return end > parentIndex ? { start: parentIndex + 1, end: end } : null;
  }

  function isGroup(tasks, index) {
    return descendantRange(tasks, index) != null;
  }

  function getDescendantIndexes(tasks, parentIndex) {
    var range = descendantRange(tasks, parentIndex);
    if (!range) return [];
    var idxs = [];
    for (var i = range.start; i <= range.end; i++) idxs.push(i);
    return idxs;
  }

  function getBlockIndexes(tasks, index) {
    return [index].concat(getDescendantIndexes(tasks, index));
  }

  function outlineNumber(tasks, index) {
    var parts = [];
    var counts = [0];
    for (var i = 0; i <= index; i++) {
      var indent = tasks[i].indent || 0;
      counts = counts.slice(0, indent + 1);
      while (counts.length <= indent) counts.push(0);
      counts[indent] = (counts[indent] || 0) + 1;
      if (i === index) {
        for (var k = 0; k <= indent; k++) parts.push(String(counts[k] || 1));
      }
    }
    return parts.join(".");
  }

  function detectCycle(tasks) {
    var byNum = {};
    tasks.forEach(function (t) {
      byNum[t.number] = t;
    });
    var visiting = {};
    var visited = {};
    var cyclic = {};

    function dfs(num) {
      if (visited[num]) return false;
      if (visiting[num]) return true;
      visiting[num] = true;
      var task = byNum[num];
      var deps = normalizeDeps(task && task.dependencies);
      for (var i = 0; i < deps.length; i++) {
        if (!byNum[deps[i].pred]) continue;
        if (dfs(deps[i].pred)) {
          cyclic[num] = true;
          return true;
        }
      }
      delete visiting[num];
      visited[num] = true;
      return false;
    }

    tasks.forEach(function (t) {
      dfs(t.number);
    });
    return cyclic;
  }

  function topologicalNumbers(tasks) {
    var byNum = {};
    tasks.forEach(function (t) {
      byNum[t.number] = t;
    });
    var cyclic = detectCycle(tasks);
    var visited = {};
    var order = [];

    function dfs(num) {
      if (visited[num]) return;
      visited[num] = true;
      var task = byNum[num];
      if (!task) return;
      if (!cyclic[num]) {
        normalizeDeps(task.dependencies).forEach(function (d) {
          if (byNum[d.pred]) dfs(d.pred);
        });
      }
      order.push(num);
    }

    tasks.forEach(function (t) {
      dfs(t.number);
    });
    return { order: order, cyclic: cyclic };
  }

  function shiftTask(task, deltaDays) {
    if (!deltaDays) return;
    if (isISO(task.start)) task.start = addDays(task.start, deltaDays);
    if (isISO(task.end)) task.end = addDays(task.end, deltaDays);
  }

  function shiftBlock(tasks, index, deltaDays) {
    if (!deltaDays) return;
    getBlockIndexes(tasks, index).forEach(function (i) {
      shiftTask(tasks[i], deltaDays);
    });
  }

  function clampProgress(v) {
    var n = Math.round(Number(v) || 0);
    if (n < 0) return 0;
    if (n > 100) return 100;
    return n;
  }

  function clampAllocation(v) {
    if (v == null || v === "") return 100;
    var n = Math.round(Number(v));
    if (!Number.isFinite(n)) return 100;
    if (n < 0) return 0;
    if (n > 400) return 400;
    return n;
  }

  function taskDuration(task) {
    if (task.milestone) return 0;
    var d = Math.round(Number(task.duration));
    if (!Number.isFinite(d) || d < 0) return 0;
    return d;
  }

  function constraintFromDep(pred, dep, cal) {
    var type = (dep && dep.type) || "FS";
    var lag = dep && dep.lag != null ? Number(dep.lag) : 0;
    var result = { minStart: null, minEnd: null };
    if (type === "FS" && isISO(pred.end)) {
      result.minStart = nextUnit(addLag(pred.end, lag, cal), cal);
    } else if (type === "SS" && isISO(pred.start)) {
      result.minStart = addLag(pred.start, lag, cal);
      if (cal && cal.mode === "workdays") result.minStart = snapToWorkday(result.minStart, cal, 1);
    } else if (type === "FF" && isISO(pred.end)) {
      result.minEnd = addLag(pred.end, lag, cal);
      if (cal && cal.mode === "workdays") result.minEnd = snapToWorkday(result.minEnd, cal, 1);
    } else if (type === "SF" && isISO(pred.start)) {
      result.minEnd = addLag(pred.start, lag, cal);
      if (cal && cal.mode === "workdays") result.minEnd = snapToWorkday(result.minEnd, cal, 1);
    }
    return result;
  }

  function earliestStartFromConstraints(task, constraints, cal) {
    var minStart = null;
    var minEnd = null;
    (constraints || []).forEach(function (c) {
      if (c.minStart) minStart = maxISO(minStart, c.minStart);
      if (c.minEnd) minEnd = maxISO(minEnd, c.minEnd);
    });
    var fromEnd = minEnd != null ? startFromEndDuration(minEnd, taskDuration(task), cal) : null;
    return maxISO(minStart, fromEnd);
  }

  function defaultStartFor(task, allTasks, today, cal, cyclic) {
    var constraints = [];
    normalizeDeps(task.dependencies).forEach(function (d) {
      if (cyclic && cyclic[d.pred]) return;
      var p = findTaskByNumber(allTasks, d.pred);
      if (p) constraints.push(constraintFromDep(p, d, cal));
    });
    var start = earliestStartFromConstraints(task, constraints, cal);
    if (start) return start;
    var t = today || todayISO();
    if (cal && cal.mode === "workdays") return snapToWorkday(t, cal, 1);
    return t;
  }

  function normalizeLeaf(task, allTasks, today, cal, cyclic) {
    var milestone = !!(task.milestone || Number(task.duration) === 0 && task.duration !== "" && task.duration != null && !isGroupPlaceholder(task));
    if (task.milestone) milestone = true;
    var start = isISO(task.start) ? task.start : null;
    var end = isISO(task.end) ? task.end : null;
    var duration =
      task.duration == null || task.duration === ""
        ? milestone
          ? 0
          : null
        : Math.round(Number(task.duration));
    if (duration != null && !Number.isFinite(duration)) duration = null;
    if (duration != null && duration < 0) duration = 0;
    if (milestone) {
      duration = 0;
      if (!start && end) start = end;
      if (!start) start = defaultStartFor(task, allTasks, today, cal, cyclic);
      if (cal && cal.mode === "workdays") start = snapToWorkday(start, cal, 1);
      task.milestone = true;
      task.start = start;
      task.end = start;
      task.duration = 0;
      return;
    }
    task.milestone = false;
    if (cal && cal.mode === "workdays") {
      if (start) start = snapToWorkday(start, cal, 1);
      if (end) end = snapToWorkday(end, cal, 1);
    }
    if (start && end && end < start) end = start;

    if (start && duration != null) {
      end = endFromStartDuration(start, duration, cal);
    } else if (start && end) {
      duration = inclusiveDuration(start, end, cal);
      if (duration < 1) duration = 1;
    } else if (end && duration != null) {
      start = startFromEndDuration(end, duration, cal);
    } else if (start) {
      duration = 1;
      end = endFromStartDuration(start, duration, cal);
    } else if (end) {
      duration = 1;
      start = startFromEndDuration(end, duration, cal);
    } else {
      start = defaultStartFor(task, allTasks, today, cal, cyclic);
      duration = duration != null ? Math.max(1, duration) : 1;
      end = endFromStartDuration(start, duration, cal);
    }

    task.start = start;
    task.end = end;
    task.duration = duration;
  }

  function isGroupPlaceholder(task) {
    return false;
  }

  function recomputeGroups(tasks, cal) {
    var i, j, desc, indent, sources, starts, ends, progNum, progDen, t, start, end, dur, prog;
    for (i = 0; i < tasks.length; i++) tasks[i].isGroup = false;

    for (i = tasks.length - 1; i >= 0; i--) {
      desc = getDescendantIndexes(tasks, i);
      if (!desc.length) continue;
      tasks[i].isGroup = true;
      tasks[i].milestone = false;
      indent = tasks[i].indent || 0;
      sources = desc.filter(function (idx) {
        return (tasks[idx].indent || 0) === indent + 1;
      });
      if (!sources.length) sources = desc;
      starts = [];
      ends = [];
      progNum = 0;
      progDen = 0;
      for (j = 0; j < sources.length; j++) {
        t = tasks[sources[j]];
        if (isISO(t.start)) starts.push(t.start);
        if (isISO(t.end)) ends.push(t.end);
        dur = Number(t.duration) || 0;
        if (t.milestone) dur = 0;
        prog = clampProgress(t.progress);
        progNum += prog * Math.max(dur, t.milestone ? 1 : 0);
        progDen += Math.max(dur, t.milestone ? 1 : 0);
      }
      start = minISO.apply(null, starts);
      end = maxISO.apply(null, ends);
      if (start && end) {
        tasks[i].start = start;
        tasks[i].end = end;
        tasks[i].duration = inclusiveDuration(start, end, cal);
      }
      if (progDen > 0) tasks[i].progress = Math.round(progNum / progDen);
    }
  }

  function applyDependencies(tasks, cyclic, cal) {
    var topo = topologicalNumbers(tasks);
    var order = topo.order;
    var changed = false;
    var k, num, idx, task, constraints, target, delta;

    for (k = 0; k < order.length; k++) {
      num = order[k];
      if (cyclic[num]) continue;
      idx = findIndexByNumber(tasks, num);
      if (idx < 0) continue;
      task = tasks[idx];
      constraints = [];
      normalizeDeps(task.dependencies).forEach(function (d) {
        if (cyclic[d.pred]) return;
        var p = findTaskByNumber(tasks, d.pred);
        if (p) constraints.push(constraintFromDep(p, d, cal));
      });
      target = earliestStartFromConstraints(task, constraints, cal);
      if (target && (!isISO(task.start) || task.start < target)) {
        if (isISO(task.start)) {
          delta = diffDays(task.start, target);
          shiftBlock(tasks, idx, delta);
          if (!task.isGroup) {
            task.start = target;
            task.end = endFromStartDuration(task.start, taskDuration(task), cal);
          }
        } else {
          task.start = target;
          task.end = endFromStartDuration(task.start, taskDuration(task) || (task.milestone ? 0 : 1), cal);
        }
        changed = true;
      }
    }
    return changed;
  }

  function finishVariance(task, cal) {
    if (!isISO(task.end) || !isISO(task.baselineEnd)) return null;
    return shiftOffset(task.baselineEnd, task.end, cal);
  }

  function computeCPM(tasks, cal) {
    var i, t, es, ef, ls, lf, projectEnd, constraints, target;
    var topo = topologicalNumbers(tasks);
    var cyclic = topo.cyclic;
    var byNum = {};
    tasks.forEach(function (task) {
      byNum[task.number] = task;
      task.earlyStart = null;
      task.earlyEnd = null;
      task.lateStart = null;
      task.lateEnd = null;
      task.totalFloat = null;
      task.critical = false;
    });

    topo.order.forEach(function (num) {
      t = byNum[num];
      if (!t) return;
      constraints = [];
      if (!cyclic[num]) {
        normalizeDeps(t.dependencies).forEach(function (d) {
          var p = byNum[d.pred];
          if (!p || !isISO(p.earlyStart)) return;
          constraints.push(
            constraintFromDep(
              { start: p.earlyStart, end: p.earlyEnd },
              d,
              cal
            )
          );
        });
      }
      target = earliestStartFromConstraints(t, constraints, cal);
      es = maxISO(target, t.start);
      if (!es) es = t.start;
      if (cal && cal.mode === "workdays" && es && !t.milestone) es = snapToWorkday(es, cal, 1);
      ef = t.milestone || taskDuration(t) === 0 ? es : endFromStartDuration(es, taskDuration(t), cal);
      if (t.isGroup && isISO(t.end) && t.end > ef) ef = t.end;
      t.earlyStart = es;
      t.earlyEnd = ef;
    });

    // Roll group early dates from children (bottom-up) so a summary's EF is the latest child.
    for (i = tasks.length - 1; i >= 0; i--) {
      if (!tasks[i].isGroup) continue;
      var desc = getDescendantIndexes(tasks, i);
      var starts = [];
      var ends = [];
      desc.forEach(function (idx) {
        if (isISO(tasks[idx].earlyStart)) starts.push(tasks[idx].earlyStart);
        if (isISO(tasks[idx].earlyEnd)) ends.push(tasks[idx].earlyEnd);
      });
      if (starts.length) tasks[i].earlyStart = minISO.apply(null, starts);
      if (ends.length) tasks[i].earlyEnd = maxISO.apply(null, ends);
    }

    projectEnd = null;
    tasks.forEach(function (task) {
      projectEnd = maxISO(projectEnd, task.earlyEnd, task.end);
    });

    var successors = {};
    tasks.forEach(function (task) {
      successors[task.number] = [];
    });
    tasks.forEach(function (task) {
      if (cyclic[task.number]) return;
      normalizeDeps(task.dependencies).forEach(function (d) {
        if (!byNum[d.pred]) return;
        successors[d.pred].push({ succ: task, dep: d });
      });
    });

    function latestFinishFromSucc(pred, link, cal) {
      var s = link.succ;
      var dep = link.dep;
      var type = dep.type || "FS";
      var lag = dep.lag || 0;
      var result = { lateStart: null, lateEnd: null };
      if (!isISO(s.lateStart) || !isISO(s.lateEnd)) return result;
      if (type === "FS") {
        // pred.LF <= prevUnit(succ.LS) lag inverted
        result.lateEnd = addLag(prevUnit(s.lateStart, cal), -lag, cal);
      } else if (type === "SS") {
        result.lateStart = addLag(s.lateStart, -lag, cal);
      } else if (type === "FF") {
        result.lateEnd = addLag(s.lateEnd, -lag, cal);
      } else if (type === "SF") {
        result.lateStart = addLag(s.lateEnd, -lag, cal);
      }
      return result;
    }

    var reverse = topo.order.slice().reverse();
    reverse.forEach(function (num) {
      t = byNum[num];
      if (!t) return;
      lf = projectEnd;
      ls = null;
      var links = successors[num] || [];
      if (links.length && !cyclic[num]) {
        lf = null;
        ls = null;
        links.forEach(function (link) {
          var bound = latestFinishFromSucc(t, link, cal);
          if (bound.lateEnd) lf = lf == null ? bound.lateEnd : minISO(lf, bound.lateEnd);
          if (bound.lateStart) ls = ls == null ? bound.lateStart : minISO(ls, bound.lateStart);
        });
      }
      var fromEnd = lf != null ? startFromEndDuration(lf, taskDuration(t), cal) : null;
      var fromStartEnd = ls != null ? endFromStartDuration(ls, taskDuration(t), cal) : null;
      if (ls == null) ls = fromEnd;
      else if (fromEnd) ls = minISO(ls, fromEnd);
      if (lf == null) lf = fromStartEnd;
      else if (fromStartEnd) lf = minISO(lf, fromStartEnd);
      if (ls == null && isISO(t.earlyStart) && projectEnd) {
        ls = startFromEndDuration(projectEnd, taskDuration(t), cal);
        lf = projectEnd;
      }
      if (t.milestone) {
        if (lf != null && ls == null) ls = lf;
        if (ls != null && lf == null) lf = ls;
      }
      t.lateStart = ls;
      t.lateEnd = lf;
    });

    // Push a group's successor-driven late finish down onto its children, then
    // roll group late dates back up without loosening a tighter group constraint.
    for (i = 0; i < tasks.length; i++) {
      if (!tasks[i].isGroup || !isISO(tasks[i].lateEnd)) continue;
      getDescendantIndexes(tasks, i).forEach(function (idx) {
        var child = tasks[idx];
        if (isISO(child.lateEnd)) child.lateEnd = minISO(child.lateEnd, tasks[i].lateEnd);
        else child.lateEnd = tasks[i].lateEnd;
        child.lateStart = startFromEndDuration(child.lateEnd, taskDuration(child), cal);
        if (child.milestone) child.lateStart = child.lateEnd;
      });
    }
    for (i = tasks.length - 1; i >= 0; i--) {
      if (!tasks[i].isGroup) continue;
      var dIdx = getDescendantIndexes(tasks, i);
      var lStarts = [];
      var lEnds = [];
      dIdx.forEach(function (idx) {
        if (isISO(tasks[idx].lateStart)) lStarts.push(tasks[idx].lateStart);
        if (isISO(tasks[idx].lateEnd)) lEnds.push(tasks[idx].lateEnd);
      });
      if (lStarts.length) {
        var rolledLS = minISO.apply(null, lStarts);
        tasks[i].lateStart = isISO(tasks[i].lateStart) ? minISO(tasks[i].lateStart, rolledLS) : rolledLS;
      }
      if (lEnds.length) {
        var rolledLF = maxISO.apply(null, lEnds);
        tasks[i].lateEnd = isISO(tasks[i].lateEnd) ? minISO(tasks[i].lateEnd, rolledLF) : rolledLF;
      }
    }

    tasks.forEach(function (task) {
      if (!isISO(task.earlyStart) || !isISO(task.lateStart)) {
        task.totalFloat = 0;
        task.critical = !!task.start;
        return;
      }
      var classic = shiftOffset(task.earlyStart, task.lateStart, cal);
      var remaining = isISO(task.start) ? shiftOffset(task.start, task.lateStart, cal) : classic;
      task.totalFloat = classic;
      task.remainingFloat = remaining;
      task.critical = remaining <= 0 && classic <= 0;
      if (task.isGroup) {
        var kids = getDescendantIndexes(tasks, findIndexByNumber(tasks, task.number));
        var any = kids.some(function (idx) {
          return tasks[idx].critical && !tasks[idx].isGroup;
        });
        if (any && classic <= 0) task.critical = true;
      }
    });

    return { projectEnd: projectEnd };
  }

  function isCriticalDep(pred, succ, dep, cal) {
    if (!pred || !succ || !pred.critical || !succ.critical) return false;
    var c = constraintFromDep(pred, dep, cal);
    if (c.minStart && isISO(succ.start) && succ.start === c.minStart) return true;
    if (c.minEnd && isISO(succ.end) && succ.end === c.minEnd) return true;
    return false;
  }

  function reschedule(tasks, options) {
    options = options || {};
    var today = options.today || todayISO();
    var cal = options.calendar || makeCalendar(options.mode || "calendar", options.holidays || []);
    var i, cyclic, guard, changed;

    for (i = 0; i < tasks.length; i++) {
      tasks[i].indent = Math.max(0, Math.round(Number(tasks[i].indent) || 0));
      tasks[i].dependencies = normalizeDeps(tasks[i].dependencies).filter(function (d) {
        return d.pred !== tasks[i].number;
      });
      tasks[i].progress = clampProgress(tasks[i].progress);
      tasks[i].allocation = clampAllocation(tasks[i].allocation);
      if (tasks[i].baselineStart && !isISO(tasks[i].baselineStart)) tasks[i].baselineStart = null;
      if (tasks[i].baselineEnd && !isISO(tasks[i].baselineEnd)) tasks[i].baselineEnd = null;
      if (tasks[i].milestone) {
        tasks[i].duration = 0;
      }
    }

    cyclic = detectCycle(tasks);

    for (i = 0; i < tasks.length; i++) {
      if (!isGroup(tasks, i)) normalizeLeaf(tasks[i], tasks, today, cal, cyclic);
    }

    guard = 0;
    while (guard++ < 30) {
      for (i = 0; i < tasks.length; i++) {
        if (!isGroup(tasks, i)) normalizeLeaf(tasks[i], tasks, today, cal, cyclic);
      }
      recomputeGroups(tasks, cal);
      changed = applyDependencies(tasks, cyclic, cal);
      recomputeGroups(tasks, cal);
      if (!changed) break;
    }

    computeCPM(tasks, cal);

    for (i = 0; i < tasks.length; i++) {
      tasks[i].variance = finishVariance(tasks[i], cal);
    }

    return { tasks: tasks, cyclic: cyclic, iterations: guard, calendar: cal };
  }

  function eachDay(start, end, fn) {
    if (!isISO(start) || !isISO(end) || end < start) return;
    var cur = start;
    while (cur <= end) {
      fn(cur, parseISO(cur));
      cur = addDays(cur, 1);
    }
  }

  function setBaseline(tasks) {
    tasks.forEach(function (t) {
      t.baselineStart = isISO(t.start) ? t.start : null;
      t.baselineEnd = isISO(t.end) ? t.end : null;
    });
  }

  function clearBaseline(tasks) {
    tasks.forEach(function (t) {
      t.baselineStart = null;
      t.baselineEnd = null;
      t.variance = null;
    });
  }

  function resourceUsage(tasks, rangeStart, rangeEnd, cal) {
    var byName = {};
    tasks.forEach(function (t) {
      if (t.isGroup) return;
      var name = (t.assignee || "").trim();
      if (!name) return;
      if (!byName[name]) byName[name] = { name: name, days: {}, overDays: [], totalAlloc: 0 };
      if (!isISO(t.start) || !isISO(t.end)) return;
      var alloc = clampAllocation(t.allocation);
      eachDay(t.start, t.end, function (iso) {
        if (rangeStart && iso < rangeStart) return;
        if (rangeEnd && iso > rangeEnd) return;
        if (cal && cal.mode === "workdays" && !isWorking(iso, cal)) return;
        if (t.milestone && iso !== t.start) return;
        byName[name].days[iso] = (byName[name].days[iso] || 0) + (t.milestone ? alloc : alloc);
      });
    });
    var rows = Object.keys(byName)
      .sort()
      .map(function (name) {
        var row = byName[name];
        var overDays = [];
        var peak = 0;
        Object.keys(row.days).forEach(function (iso) {
          if (row.days[iso] > peak) peak = row.days[iso];
          if (row.days[iso] > 100) overDays.push(iso);
        });
        overDays.sort();
        row.overDays = overDays;
        row.peak = peak;
        row.overallocated = overDays.length > 0;
        return row;
      });
    return rows;
  }

  return {
    todayISO: todayISO,
    toISO: toISO,
    parseISO: parseISO,
    isISO: isISO,
    addDays: addDays,
    diffDays: diffDays,
    maxISO: maxISO,
    minISO: minISO,
    uniqueNums: uniqueNums,
    uniqueSortedDates: uniqueSortedDates,
    isWeekend: isWeekend,
    makeCalendar: makeCalendar,
    isWorking: isWorking,
    addWorkdays: addWorkdays,
    snapToWorkday: snapToWorkday,
    addLag: addLag,
    nextUnit: nextUnit,
    prevUnit: prevUnit,
    shiftOffset: shiftOffset,
    workingDiff: workingDiff,
    endFromStartDuration: endFromStartDuration,
    startFromEndDuration: startFromEndDuration,
    inclusiveDuration: inclusiveDuration,
    normalizeDep: normalizeDep,
    parseDependencyInput: parseDependencyInput,
    normalizeDeps: normalizeDeps,
    formatDep: formatDep,
    formatDependencies: formatDependencies,
    depPreds: depPreds,
    findTaskByNumber: findTaskByNumber,
    findIndexByNumber: findIndexByNumber,
    descendantRange: descendantRange,
    isGroup: isGroup,
    getDescendantIndexes: getDescendantIndexes,
    getBlockIndexes: getBlockIndexes,
    outlineNumber: outlineNumber,
    detectCycle: detectCycle,
    topologicalNumbers: topologicalNumbers,
    shiftTask: shiftTask,
    shiftBlock: shiftBlock,
    clampProgress: clampProgress,
    clampAllocation: clampAllocation,
    taskDuration: taskDuration,
    constraintFromDep: constraintFromDep,
    defaultStartFor: defaultStartFor,
    normalizeLeaf: normalizeLeaf,
    recomputeGroups: recomputeGroups,
    applyDependencies: applyDependencies,
    finishVariance: finishVariance,
    computeCPM: computeCPM,
    isCriticalDep: isCriticalDep,
    reschedule: reschedule,
    eachDay: eachDay,
    setBaseline: setBaseline,
    clearBaseline: clearBaseline,
    resourceUsage: resourceUsage,
  };
});
