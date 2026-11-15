/**
 * Gantt scheduling engine (dates, dependencies, groups).
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

  function inclusiveDuration(start, end) {
    var d = diffDays(start, end);
    if (d == null) return null;
    return d + 1;
  }

  function endFromStartDuration(start, duration) {
    var dur = Math.max(1, Math.round(Number(duration)));
    return addDays(start, dur - 1);
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

  function parseDependencyInput(text) {
    if (text == null || text === "") return [];
    return uniqueNums(
      String(text)
        .split(/[,;\s]+/)
        .map(function (p) {
          var m = /^#?(\d+)$/.exec(p.trim());
          return m ? Number(m[1]) : null;
        })
    );
  }

  function formatDependencies(ids) {
    return uniqueNums(ids).join(", ");
  }

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
      var deps = (task && task.dependencies) || [];
      for (var i = 0; i < deps.length; i++) {
        if (!byNum[deps[i]]) continue;
        if (dfs(deps[i])) {
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
        (task.dependencies || []).forEach(function (d) {
          if (byNum[d]) dfs(d);
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

  function defaultStartFor(task, allTasks, today) {
    var predEnds = [];
    (task.dependencies || []).forEach(function (n) {
      var p = findTaskByNumber(allTasks, n);
      if (p && isISO(p.end)) predEnds.push(p.end);
    });
    var latest = maxISO.apply(null, predEnds);
    return latest ? addDays(latest, 1) : today;
  }

  function normalizeLeaf(task, allTasks, today) {
    var start = isISO(task.start) ? task.start : null;
    var end = isISO(task.end) ? task.end : null;
    var duration =
      task.duration == null || task.duration === ""
        ? null
        : Math.round(Number(task.duration));
    if (!Number.isFinite(duration) || duration < 1) duration = duration == null || task.duration === "" ? null : 1;
    if (start && end && end < start) end = start;

    if (start && duration != null) {
      end = endFromStartDuration(start, duration);
    } else if (start && end) {
      duration = inclusiveDuration(start, end);
    } else if (end && duration != null) {
      start = addDays(end, -(duration - 1));
    } else if (start) {
      duration = 1;
      end = start;
    } else if (end) {
      duration = 1;
      start = end;
    } else {
      start = defaultStartFor(task, allTasks, today);
      duration = duration != null ? duration : 1;
      end = endFromStartDuration(start, duration);
    }

    task.start = start;
    task.end = end;
    task.duration = duration;
  }

  function recomputeGroups(tasks) {
    var i, j, desc, indent, sources, starts, ends, progNum, progDen, t, start, end, dur, prog;
    for (i = 0; i < tasks.length; i++) tasks[i].isGroup = false;

    for (i = tasks.length - 1; i >= 0; i--) {
      desc = getDescendantIndexes(tasks, i);
      if (!desc.length) continue;
      tasks[i].isGroup = true;
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
        prog = clampProgress(t.progress);
        progNum += prog * dur;
        progDen += dur;
      }
      start = minISO.apply(null, starts);
      end = maxISO.apply(null, ends);
      if (start && end) {
        tasks[i].start = start;
        tasks[i].end = end;
        tasks[i].duration = inclusiveDuration(start, end);
      }
      if (progDen > 0) tasks[i].progress = Math.round(progNum / progDen);
    }
  }

  function applyDependencies(tasks, cyclic) {
    var topo = topologicalNumbers(tasks);
    var order = topo.order;
    var changed = false;
    var i, k, num, idx, task, mins, p, minStart, delta;

    for (k = 0; k < order.length; k++) {
      num = order[k];
      if (cyclic[num]) continue;
      idx = findIndexByNumber(tasks, num);
      if (idx < 0) continue;
      task = tasks[idx];
      mins = [];
      (task.dependencies || []).forEach(function (d) {
        if (cyclic[d]) return;
        p = findTaskByNumber(tasks, d);
        if (p && isISO(p.end)) mins.push(addDays(p.end, 1));
      });
      minStart = maxISO.apply(null, mins);
      if (minStart && (!isISO(task.start) || task.start < minStart)) {
        if (isISO(task.start)) {
          delta = diffDays(task.start, minStart);
          shiftBlock(tasks, idx, delta);
        } else {
          task.start = minStart;
          if (task.duration) task.end = endFromStartDuration(task.start, task.duration);
        }
        changed = true;
      }
    }
    return changed;
  }

  function reschedule(tasks, options) {
    options = options || {};
    var today = options.today || todayISO();
    var i, cyclic, guard, changed;

    for (i = 0; i < tasks.length; i++) {
      tasks[i].indent = Math.max(0, Math.round(Number(tasks[i].indent) || 0));
      tasks[i].dependencies = uniqueNums(tasks[i].dependencies).filter(function (n) {
        return n !== tasks[i].number;
      });
      tasks[i].progress = clampProgress(tasks[i].progress);
    }

    cyclic = detectCycle(tasks);

    for (i = 0; i < tasks.length; i++) {
      if (!isGroup(tasks, i)) normalizeLeaf(tasks[i], tasks, today);
    }

    guard = 0;
    while (guard++ < 30) {
      for (i = 0; i < tasks.length; i++) {
        if (!isGroup(tasks, i)) normalizeLeaf(tasks[i], tasks, today);
      }
      recomputeGroups(tasks);
      changed = applyDependencies(tasks, cyclic);
      recomputeGroups(tasks);
      if (!changed) break;
    }

    return { tasks: tasks, cyclic: cyclic, iterations: guard };
  }

  function eachDay(start, end, fn) {
    var cur = start;
    while (cur <= end) {
      fn(cur, parseISO(cur));
      cur = addDays(cur, 1);
    }
  }

  function isWeekend(iso) {
    var d = parseISO(iso);
    if (!d) return false;
    var day = d.getDay();
    return day === 0 || day === 6;
  }

  return {
    todayISO: todayISO,
    toISO: toISO,
    parseISO: parseISO,
    isISO: isISO,
    addDays: addDays,
    diffDays: diffDays,
    inclusiveDuration: inclusiveDuration,
    endFromStartDuration: endFromStartDuration,
    maxISO: maxISO,
    minISO: minISO,
    uniqueNums: uniqueNums,
    parseDependencyInput: parseDependencyInput,
    formatDependencies: formatDependencies,
    findTaskByNumber: findTaskByNumber,
    findIndexByNumber: findIndexByNumber,
    descendantRange: descendantRange,
    isGroup: isGroup,
    getDescendantIndexes: getDescendantIndexes,
    getBlockIndexes: getBlockIndexes,
    detectCycle: detectCycle,
    topologicalNumbers: topologicalNumbers,
    shiftTask: shiftTask,
    shiftBlock: shiftBlock,
    clampProgress: clampProgress,
    defaultStartFor: defaultStartFor,
    normalizeLeaf: normalizeLeaf,
    recomputeGroups: recomputeGroups,
    applyDependencies: applyDependencies,
    reschedule: reschedule,
    eachDay: eachDay,
    isWeekend: isWeekend,
  };
});
