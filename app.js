/* Gantt chart SPA — table, timeline, persistence, import/export */
(function () {
  "use strict";

  var S = window.GanttSchedule;
  var Excel = window.GanttExcel;
  var XLSX = window.XLSX;
  var STORAGE_KEY = "gantt-chart-app:v1";
  var THEME_KEY = "gantt-chart-app:theme";
  var ROW_H = 40;
  var RES_ROW_H = 44;
  var PALETTE = [
    "#0f766e",
    "#0369a1",
    "#7c3aed",
    "#be185d",
    "#c2410c",
    "#4d7c0f",
    "#0e7490",
    "#b45309",
    "#4338ca",
    "#115e59",
  ];
  var PX = { day: 36, week: 20, month: 8 };

  var state = {
    version: 2,
    projectName: "Untitled project",
    tasks: [],
    nextNumber: 1,
    zoom: "week",
    split: 620,
    selectedId: null,
    calendarMode: "calendar",
    holidays: [],
    showCritical: false,
    view: "timeline",
    mobilePane: "timeline",
    baselineSetAt: null,
  };

  var MOBILE_BP = 900;
  var pendingInstall = null;
  var swRefreshing = false;
  var pressTimer = null;

  var historyStack = [];
  var redoStack = [];
  var MAX_HISTORY = 60;
  var typingCheckpoint = false;
  var suppressSave = false;
  var hoverId = null;
  var drag = null;
  var rowDrag = null;
  var toastTimer = null;

  function $(id) {
    return document.getElementById(id);
  }

  function uid() {
    return "t" + Math.random().toString(36).slice(2, 9) + Date.now().toString(36).slice(-4);
  }

  function clone(obj) {
    return JSON.parse(JSON.stringify(obj));
  }

  function calendar() {
    return S.makeCalendar(state.calendarMode || "calendar", state.holidays || []);
  }

  function scheduleOptions() {
    return { today: S.todayISO(), mode: state.calendarMode || "calendar", holidays: state.holidays || [] };
  }

  function snapshot() {
    return {
      projectName: state.projectName,
      tasks: clone(state.tasks),
      nextNumber: state.nextNumber,
      zoom: state.zoom,
      calendarMode: state.calendarMode,
      holidays: clone(state.holidays || []),
      showCritical: !!state.showCritical,
      view: state.view,
      mobilePane: state.mobilePane,
      baselineSetAt: state.baselineSetAt,
    };
  }

  function pushHistory() {
    historyStack.push(snapshot());
    if (historyStack.length > MAX_HISTORY) historyStack.shift();
    redoStack = [];
    updateUndoButtons();
  }

  function restoreSnap(snap) {
    state.projectName = snap.projectName;
    state.tasks = clone(snap.tasks);
    state.nextNumber = snap.nextNumber;
    state.zoom = snap.zoom;
    state.calendarMode = snap.calendarMode || "calendar";
    state.holidays = clone(snap.holidays || []);
    state.showCritical = !!snap.showCritical;
    state.view = snap.view || "timeline";
    state.mobilePane = snap.mobilePane || (state.view === "resources" ? "resources" : "timeline");
    state.baselineSetAt = snap.baselineSetAt || null;
    S.reschedule(state.tasks, scheduleOptions());
    renderAll();
    persist();
  }

  function undo() {
    if (!historyStack.length) return;
    redoStack.push(snapshot());
    restoreSnap(historyStack.pop());
    updateUndoButtons();
    toast("Undid last change");
  }

  function redo() {
    if (!redoStack.length) return;
    historyStack.push(snapshot());
    restoreSnap(redoStack.pop());
    updateUndoButtons();
    toast("Redid last change");
  }

  function updateUndoButtons() {
    var noUndo = !historyStack.length;
    var noRedo = !redoStack.length;
    ["btn-undo", "btn-redo"].forEach(function (id, i) {
      var el = $(id);
      if (el) el.disabled = i === 0 ? noUndo : noRedo;
    });
    document.querySelectorAll("[data-action=undo]").forEach(function (el) { el.disabled = noUndo; });
    document.querySelectorAll("[data-action=redo]").forEach(function (el) { el.disabled = noRedo; });
  }

  function toast(msg, opts) {
    opts = opts || {};
    var el = $("toast");
    // Keep a sticky "update available" toast until the user acts on it.
    if (el.classList.contains("toast-action") && el.classList.contains("show") && !(opts.actionLabel && opts.onAction)) {
      return;
    }
    el.classList.remove("toast-action");
    if (opts.actionLabel && opts.onAction) {
      el.classList.add("toast-action");
      el.innerHTML = "";
      el.appendChild(document.createTextNode(msg));
      var btn = document.createElement("button");
      btn.type = "button";
      btn.textContent = opts.actionLabel;
      btn.addEventListener("click", opts.onAction);
      el.appendChild(btn);
      el.classList.add("show");
      clearTimeout(toastTimer);
      return;
    }
    el.textContent = msg;
    el.classList.add("show");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () {
      el.classList.remove("show");
    }, opts.ms || 2400);
  }

  function persistPayload() {
    return {
      version: 2,
      projectName: state.projectName,
      tasks: state.tasks,
      nextNumber: state.nextNumber,
      zoom: state.zoom,
      split: state.split,
      calendarMode: state.calendarMode,
      holidays: state.holidays,
      showCritical: state.showCritical,
      view: state.view,
      mobilePane: state.mobilePane,
      baselineSetAt: state.baselineSetAt,
    };
  }

  function persist() {
    if (suppressSave) return;
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(persistPayload()));
      $("status-saved").textContent = "Saved in this browser";
    } catch (err) {
      $("status-saved").textContent = "Could not save (storage full?)";
    }
  }

  function loadPersisted() {
    try {
      var raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return false;
      var data = JSON.parse(raw);
      applyProject(data, { resetHistory: true });
      if (typeof data.split === "number") state.split = data.split;
      return true;
    } catch (err) {
      return false;
    }
  }

  function applyProject(data, opts) {
    opts = opts || {};
    var tasks = [];
    if (Array.isArray(data)) tasks = data;
    else if (data && Array.isArray(data.tasks)) tasks = data.tasks;
    else throw new Error("JSON must be a project object with a tasks array, or a task array.");

    state.projectName = (data && data.projectName) || "Untitled project";
    state.calendarMode = data && data.calendarMode === "workdays" ? "workdays" : "calendar";
    state.holidays = S.uniqueSortedDates((data && data.holidays) || []);
    state.showCritical = !!(data && data.showCritical);
    if (data && (data.view === "timeline" || data.view === "resources")) state.view = data.view;
    else state.view = "timeline";
    if (data && (data.mobilePane === "table" || data.mobilePane === "timeline" || data.mobilePane === "resources")) {
      state.mobilePane = data.mobilePane;
    } else {
      state.mobilePane = state.view === "resources" ? "resources" : "timeline";
    }
    state.baselineSetAt = (data && data.baselineSetAt) || null;
    var next = 1;
    state.tasks = tasks.map(function (raw, i) {
      var num = Number(raw.number) > 0 ? Math.round(Number(raw.number)) : i + 1;
      if (num >= next) next = num + 1;
      return normalizeIncoming(raw, num);
    });
    state.nextNumber = Math.max(next, Number(data && data.nextNumber) || 1);
    if (data && (data.zoom === "day" || data.zoom === "week" || data.zoom === "month")) {
      state.zoom = data.zoom;
    }
    S.reschedule(state.tasks, scheduleOptions());
    if (!state.baselineSetAt && state.tasks.some(function (t) { return t.baselineStart || t.baselineEnd; })) {
      state.baselineSetAt = S.todayISO();
    }
    if (opts.resetHistory) {
      historyStack = [];
      redoStack = [];
    }
    renderAll();
    persist();
  }

  function normalizeIncoming(raw, num) {
    var deps = S.parseDependencyInput(
      Array.isArray(raw.dependencies) ? raw.dependencies : raw.dependencies || ""
    );
    var milestone = !!(raw.milestone || raw.duration === 0);
    return {
      id: raw.id || uid(),
      number: num,
      name: raw.name != null ? String(raw.name) : "",
      start: S.isISO(raw.start) ? raw.start : null,
      end: S.isISO(raw.end) ? raw.end : null,
      duration: milestone ? 0 : raw.duration == null || raw.duration === "" ? null : Number(raw.duration),
      progress: S.clampProgress(raw.progress),
      dependencies: deps,
      assignee: raw.assignee != null ? String(raw.assignee) : "",
      allocation: S.clampAllocation(raw.allocation),
      milestone: milestone,
      baselineStart: S.isISO(raw.baselineStart) ? raw.baselineStart : null,
      baselineEnd: S.isISO(raw.baselineEnd) ? raw.baselineEnd : null,
      color: /^#?[0-9a-fA-F]{6}$/.test(raw.color || "")
        ? raw.color[0] === "#"
          ? raw.color
          : "#" + raw.color
        : PALETTE[(num - 1) % PALETTE.length],
      indent: Math.max(0, Math.round(Number(raw.indent) || 0)),
      collapsed: !!raw.collapsed,
    };
  }

  function nextColor() {
    return PALETTE[state.tasks.length % PALETTE.length];
  }

  function createTask(partial) {
    partial = partial || {};
    var milestone = !!(partial.milestone || partial.duration === 0);
    return {
      id: uid(),
      number: state.nextNumber++,
      name: partial.name || "",
      start: partial.start || S.todayISO(),
      duration: milestone ? 0 : partial.duration != null ? partial.duration : 1,
      end: partial.end || null,
      progress: partial.progress || 0,
      dependencies: S.parseDependencyInput(partial.dependencies || []),
      assignee: partial.assignee || "",
      allocation: S.clampAllocation(partial.allocation == null ? 100 : partial.allocation),
      milestone: milestone,
      baselineStart: partial.baselineStart || null,
      baselineEnd: partial.baselineEnd || null,
      color: partial.color || nextColor(),
      indent: partial.indent || 0,
      collapsed: false,
    };
  }

  function getTask(id) {
    for (var i = 0; i < state.tasks.length; i++) {
      if (state.tasks[i].id === id) return state.tasks[i];
    }
    return null;
  }

  function indexById(id) {
    for (var i = 0; i < state.tasks.length; i++) {
      if (state.tasks[i].id === id) return i;
    }
    return -1;
  }

  function isRowVisible(index) {
    var indent = state.tasks[index].indent || 0;
    for (var i = index - 1; i >= 0; i--) {
      var pi = state.tasks[i].indent || 0;
      if (pi < indent) {
        if (state.tasks[i].collapsed) return false;
        indent = pi;
      }
    }
    return true;
  }

  function visibleEntries() {
    var out = [];
    for (var i = 0; i < state.tasks.length; i++) {
      if (isRowVisible(i)) out.push({ task: state.tasks[i], index: i });
    }
    return out;
  }

  /* ---- sample ----------------------------------------------------------- */

  function sampleProject() {
    return {
      projectName: "Harbor site launch",
      zoom: "week",
      calendarMode: "workdays",
      holidays: ["2026-11-26", "2026-12-25"],
      showCritical: true,
      view: "timeline",
      baselineSetAt: "2026-11-02",
      tasks: [
        { number: 1, name: "Discovery", indent: 0, color: "#0f766e", assignee: "" },
        { number: 2, name: "Kickoff", start: "2026-11-02", duration: 0, milestone: true, progress: 100, indent: 1, color: "#0f766e", assignee: "Alex", allocation: 100, baselineStart: "2026-11-02", baselineEnd: "2026-11-02" },
        { number: 3, name: "Stakeholder interviews", duration: 4, progress: 100, indent: 1, color: "#0f766e", assignee: "Alex", allocation: 100, dependencies: "2FS", baselineStart: "2026-11-03", baselineEnd: "2026-11-06" },
        { number: 4, name: "Requirements document", duration: 5, progress: 100, indent: 1, color: "#0f766e", assignee: "Jordan", allocation: 100, dependencies: "3FS", baselineStart: "2026-11-09", baselineEnd: "2026-11-13" },
        { number: 5, name: "Design", indent: 0, color: "#7c3aed", dependencies: "4FS" },
        { number: 6, name: "Wireframes", duration: 4, progress: 40, indent: 1, color: "#7c3aed", assignee: "Sam", allocation: 100, dependencies: "4FS", baselineStart: "2026-11-16", baselineEnd: "2026-11-19" },
        { number: 7, name: "Visual design", duration: 6, progress: 0, indent: 1, color: "#7c3aed", assignee: "Sam", allocation: 100, dependencies: "6SS+2", baselineStart: "2026-11-18", baselineEnd: "2026-11-25" },
        { number: 8, name: "Design review", duration: 2, progress: 0, indent: 1, color: "#7c3aed", assignee: "Alex", allocation: 100, dependencies: "7FS+1", baselineStart: "2026-11-27", baselineEnd: "2026-11-30" },
        { number: 9, name: "Build", indent: 0, color: "#0369a1", dependencies: "8FS" },
        { number: 10, name: "Frontend implementation", duration: 9, indent: 1, color: "#0369a1", assignee: "Riley", allocation: 100, dependencies: "8FS", baselineStart: "2026-12-01", baselineEnd: "2026-12-11" },
        { number: 11, name: "Backend APIs", duration: 8, indent: 1, color: "#0369a1", assignee: "Quinn", allocation: 100, dependencies: "8FS", baselineStart: "2026-12-01", baselineEnd: "2026-12-10" },
        { number: 12, name: "Integrations", duration: 4, indent: 1, color: "#0369a1", assignee: "Quinn", allocation: 100, dependencies: "11SS+3", baselineStart: "2026-12-04", baselineEnd: "2026-12-09" },
        { number: 13, name: "Launch", indent: 0, color: "#c2410c", dependencies: "9FS" },
        { number: 14, name: "QA & bug bash", duration: 3, indent: 1, color: "#c2410c", assignee: "Jordan", allocation: 50, dependencies: "9FS, 12FF", baselineStart: "2026-12-11", baselineEnd: "2026-12-15" },
        { number: 15, name: "Staging deploy", duration: 1, indent: 1, color: "#c2410c", assignee: "Riley", allocation: 100, dependencies: "14FS", baselineStart: "2026-12-16", baselineEnd: "2026-12-16" },
        { number: 16, name: "Go live", duration: 0, milestone: true, indent: 1, color: "#be185d", assignee: "Alex", allocation: 100, dependencies: "15FS", baselineStart: "2026-12-17", baselineEnd: "2026-12-17" },
      ],
    };
  }

  /* ---- schedule + field edits ------------------------------------------ */

  function runSchedule() {
    var result = S.reschedule(state.tasks, scheduleOptions());
    var cyclic = Object.keys(result.cyclic || {});
    var warn = $("status-cycle");
    if (cyclic.length) {
      warn.hidden = false;
      warn.textContent =
        "Circular dependency among tasks " +
        cyclic.join(", ") +
        " — those links are ignored until the cycle is removed.";
    } else {
      warn.hidden = true;
      warn.textContent = "";
    }
    return result;
  }

  function applyField(task, field, value) {
    if (field === "name") task.name = value;
    else if (field === "assignee") task.assignee = value;
    else if (field === "color") task.color = value;
    else if (field === "allocation") {
      if (value === "" || value === "-") return false;
      task.allocation = S.clampAllocation(value);
    } else if (field === "milestone") {
      var on = value === true || value === "true" || value === "on" || value === "1";
      task.milestone = on;
      if (on) {
        task.duration = 0;
        if (S.isISO(task.start)) task.end = task.start;
      } else if (!task.duration) {
        task.duration = 1;
      }
    } else if (field === "progress") {
      if (value === "" || value === "-") return false;
      task.progress = S.clampProgress(value);
    } else if (field === "dependencies") {
      task.dependencies = S.parseDependencyInput(value);
    } else if (field === "start") {
      if (value === "") {
        task.start = null;
      } else if (S.isISO(value)) {
        task.start = value;
      } else {
        return false;
      }
    } else if (field === "duration") {
      if (value === "" || value === "-") return false;
      var n = Number(value);
      if (!Number.isFinite(n)) return false;
      n = Math.round(n);
      if (n <= 0) {
        task.duration = 0;
        task.milestone = true;
      } else {
        task.duration = n;
        task.milestone = false;
      }
    } else if (field === "end") {
      if (value === "") {
        return false;
      } else if (S.isISO(value)) {
        if (task.milestone) {
          task.start = value;
          task.end = value;
        } else if (S.isISO(task.start) && value < task.start) {
          task.end = task.start;
          task.duration = 1;
        } else if (S.isISO(task.start)) {
          task.end = value;
          var dur = S.inclusiveDuration(task.start, task.end, calendar());
          task.duration = dur < 1 ? 1 : dur;
        } else {
          task.end = value;
        }
      } else {
        return false;
      }
    }
    return true;
  }

  function onFieldInput(id, field, value) {
    var task = getTask(id);
    if (!task) return;
    if (task.isGroup && (field === "start" || field === "end" || field === "duration" || field === "progress" || field === "milestone")) {
      return;
    }
    if (!applyField(task, field, value)) return;
    runSchedule();
    persist();
    syncComputedCells();
    renderTimeline();
    updateStatus();
  }

  /* ---- structural edits ------------------------------------------------- */

  function addTask(partial, opts) {
    opts = opts || {};
    if (!opts.silent) pushHistory();
    var task = createTask(partial);
    state.tasks.push(task);
    runSchedule();
    state.selectedId = task.id;
    if (!opts.silent) {
      renderAll();
      persist();
      focusCell(task.id, "name");
    }
    return task;
  }

  function deleteTask(id) {
    var idx = indexById(id);
    if (idx < 0) return;
    pushHistory();
    var num = state.tasks[idx].number;
    state.tasks.splice(idx, 1);
    normalizeIndents();
    state.tasks.forEach(function (t) {
      t.dependencies = S.normalizeDeps(t.dependencies).filter(function (d) {
        return d.pred !== num;
      });
    });
    runSchedule();
    renderAll();
    persist();
  }

  function normalizeIndents() {
    for (var i = 0; i < state.tasks.length; i++) {
      var maxIndent = i === 0 ? 0 : (state.tasks[i - 1].indent || 0) + 1;
      if ((state.tasks[i].indent || 0) > maxIndent) state.tasks[i].indent = maxIndent;
    }
  }

  function canIndent(index) {
    if (index <= 0) return false;
    var mine = state.tasks[index].indent || 0;
    var prev = state.tasks[index - 1].indent || 0;
    return mine < prev + 1;
  }

  function indentBlock(index) {
    if (!canIndent(index)) return;
    pushHistory();
    S.getBlockIndexes(state.tasks, index).forEach(function (i) {
      state.tasks[i].indent = (state.tasks[i].indent || 0) + 1;
    });
    runSchedule();
    renderAll();
    persist();
  }

  function outdentBlock(index) {
    if ((state.tasks[index].indent || 0) <= 0) return;
    pushHistory();
    S.getBlockIndexes(state.tasks, index).forEach(function (i) {
      state.tasks[i].indent = Math.max(0, (state.tasks[i].indent || 0) - 1);
    });
    normalizeIndents();
    runSchedule();
    renderAll();
    persist();
  }

  function moveBlock(from, insertAt) {
    var block = S.getBlockIndexes(state.tasks, from);
    var blockEnd = block[block.length - 1] + 1;
    if (insertAt >= from && insertAt <= blockEnd) return false;
    var items = block.map(function (i) {
      return state.tasks[i];
    });
    var next = state.tasks.filter(function (_, i) {
      return block.indexOf(i) === -1;
    });
    var dest = 0;
    if (insertAt >= state.tasks.length) dest = next.length;
    else {
      for (var i = 0; i < insertAt; i++) {
        if (block.indexOf(i) === -1) dest++;
      }
    }
    var args = [dest, 0].concat(items);
    Array.prototype.splice.apply(next, args);
    state.tasks = next;
    return true;
  }

  function moveBlockUp(index) {
    if (index === 0) return;
    var myIndent = state.tasks[index].indent || 0;
    var i = index - 1;
    while (i > 0 && (state.tasks[i].indent || 0) > myIndent) i--;
    if ((state.tasks[i].indent || 0) < myIndent) return;
    pushHistory();
    moveBlock(index, i);
    runSchedule();
    renderAll();
    persist();
  }

  function moveBlockDown(index) {
    var block = S.getBlockIndexes(state.tasks, index);
    var after = block[block.length - 1] + 1;
    if (after >= state.tasks.length) return;
    var myIndent = state.tasks[index].indent || 0;
    if ((state.tasks[after].indent || 0) < myIndent) return;
    var nextBlock = S.getBlockIndexes(state.tasks, after);
    var insertAt = nextBlock[nextBlock.length - 1] + 1;
    pushHistory();
    moveBlock(index, insertAt);
    runSchedule();
    renderAll();
    persist();
  }

  function toggleCollapse(index) {
    if (!S.isGroup(state.tasks, index)) return;
    state.tasks[index].collapsed = !state.tasks[index].collapsed;
    renderAll();
    persist();
  }

  function newProject() {
    if (state.tasks.length && !confirm("Clear the current project? This cannot be undone after you continue (it is not kept in undo).")) {
      return;
    }
    pushHistory();
    state.projectName = "Untitled project";
    state.tasks = [];
    state.nextNumber = 1;
    state.selectedId = null;
    state.calendarMode = "calendar";
    state.holidays = [];
    state.showCritical = false;
    state.view = "timeline";
    state.mobilePane = "timeline";
    state.baselineSetAt = null;
    renderAll();
    persist();
  }

  function loadSample() {
    if (state.tasks.length && !confirm("Replace the current project with the sample?")) return;
    pushHistory();
    applyProject(sampleProject());
    toast("Loaded sample project");
  }

  function setBaseline() {
    pushHistory();
    S.setBaseline(state.tasks);
    state.baselineSetAt = S.todayISO();
    runSchedule();
    renderAll();
    persist();
    toast("Baseline saved");
  }

  function clearBaseline() {
    if (!state.baselineSetAt && !state.tasks.some(function (t) { return t.baselineStart || t.baselineEnd; })) return;
    pushHistory();
    S.clearBaseline(state.tasks);
    state.baselineSetAt = null;
    runSchedule();
    renderAll();
    persist();
    toast("Baseline cleared");
  }

  /* ---- render table ----------------------------------------------------- */

  function icon(path) {
    return (
      '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6">' +
      path +
      "</svg>"
    );
  }

  function varianceHtml(task) {
    if (task.variance == null) return '<span class="var-val">—</span>';
    var v = task.variance;
    var cls = v > 0 ? " is-slip" : v < 0 ? " is-ahead" : "";
    var text = (v > 0 ? "+" : "") + v;
    return '<span class="var-val' + cls + '">' + text + "</span>";
  }

  function renderTable() {
    var body = $("task-body");
    var html = "";
    state.tasks.forEach(function (task, index) {
      var hidden = !isRowVisible(index);
      var group = !!task.isGroup;
      var pad = (task.indent || 0) * 18;
      var twist = group
        ? '<button type="button" class="twist' +
          (task.collapsed ? " is-collapsed" : "") +
          '" data-act="collapse" title="' +
          (task.collapsed ? "Expand" : "Collapse") +
          '" aria-label="' +
          (task.collapsed ? "Expand" : "Collapse") +
          ' group">' +
          icon('<path d="M4 6l4 4 4-4"/>') +
          "</button>"
        : '<span class="twist-spacer"></span>';
      var ro = group ? " readonly" : "";
      var title = group ? ' title="Group dates and progress are calculated from child tasks"' : "";
      html +=
        '<tr draggable="false" data-id="' +
        escapeAttr(task.id) +
        '" class="' +
        (hidden ? "is-hidden " : "") +
        (group ? "is-group " : "") +
        (task.critical && state.showCritical ? "is-critical " : "") +
        (state.selectedId === task.id ? "is-selected " : "") +
        (hoverId === task.id ? "is-hover " : "") +
        '">' +
        '<td class="col-handle"><button type="button" class="drag-handle" data-act="grab" draggable="true" title="Drag to reorder" aria-label="Reorder task">' +
        icon('<path d="M3 5h10M3 8h10M3 11h10"/>') +
        "</button></td>" +
        '<td class="col-id"><span class="num">#' +
        task.number +
        "</span></td>" +
        '<td class="col-name"><div class="name-cell">' +
        '<span class="indent-pad" style="width:' +
        pad +
        'px"></span>' +
        twist +
        '<input class="cell-input" data-field="name" value="' +
        escapeAttr(task.name) +
        '" placeholder="Task name" /></div></td>' +
        '<td class="col-start"><input class="cell-input" type="date" data-field="start" value="' +
        escapeAttr(task.start || "") +
        '"' +
        ro +
        title +
        " /></td>" +
        '<td class="col-dur"><input class="cell-input" type="number" min="0" step="1" data-field="duration" value="' +
        escapeAttr(task.duration == null ? "" : String(task.duration)) +
        '"' +
        ro +
        title +
        " /></td>" +
        '<td class="col-end"><input class="cell-input" type="date" data-field="end" value="' +
        escapeAttr(task.end || "") +
        '"' +
        ro +
        title +
        " /></td>" +
        '<td class="col-prog"><input class="cell-input" type="number" min="0" max="100" step="1" data-field="progress" value="' +
        escapeAttr(String(task.progress || 0)) +
        '"' +
        ro +
        title +
        " /></td>" +
        '<td class="col-deps"><input class="cell-input" data-field="dependencies" value="' +
        escapeAttr(S.formatDependencies(task.dependencies)) +
        '" placeholder="3FS+2, 5SS" title="Predecessors: 3, 4FS+2, 5SS-1, 6FF, 7SF" /></td>' +
        '<td class="col-who"><input class="cell-input" data-field="assignee" value="' +
        escapeAttr(task.assignee) +
        '" placeholder="Name" /></td>' +
        '<td class="col-alloc"><input class="cell-input" type="number" min="0" max="400" step="5" data-field="allocation" value="' +
        escapeAttr(String(task.allocation == null ? 100 : task.allocation)) +
        '" title="Percent of the assignee\'s time" /></td>' +
        '<td class="col-ms"><div class="ms-cell"><input class="cell-check" type="checkbox" data-field="milestone"' +
        (task.milestone ? " checked" : "") +
        (group ? " disabled" : "") +
        ' title="Milestone (zero duration)" /></div></td>' +
        '<td class="col-var"><div class="var-cell">' +
        varianceHtml(task) +
        "</div></td>" +
        '<td class="col-color"><input class="color-input" type="color" data-field="color" value="' +
        escapeAttr(task.color || PALETTE[0]) +
        '" title="Bar color" /></td>' +
        '<td class="col-actions"><div class="row-actions">' +
        '<button type="button" data-act="up" title="Move up">' +
        icon('<path d="M8 3v10M4 7l4-4 4 4"/>') +
        "</button>" +
        '<button type="button" data-act="down" title="Move down">' +
        icon('<path d="M8 13V3M4 9l4 4 4-4"/>') +
        "</button>" +
        '<button type="button" data-act="outdent" title="Outdent">' +
        icon('<path d="M14 8H6M8 5l-3 3 3 3M2 3v10"/>') +
        "</button>" +
        '<button type="button" data-act="indent" title="Indent (make subtask)">' +
        icon('<path d="M2 8h8M8 5l3 3-3 3M14 3v10"/>') +
        "</button>" +
        '<button type="button" class="danger" data-act="delete" title="Delete task">' +
        icon('<path d="M3 5h10M6 5V3.5h4V5M6.5 7v5M9.5 7v5M4.5 5l.5 8h6l.5-8"/>') +
        "</button>" +
        "</div></td></tr>";
    });
    body.innerHTML = html;
    $("task-count").textContent = state.tasks.length
      ? state.tasks.length + (state.tasks.length === 1 ? " task" : " tasks")
      : "";
    $("dur-heading").textContent = state.calendarMode === "workdays" ? "Workdays" : "Days";
  }

  function escapeAttr(s) {
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/"/g, "&quot;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  function syncComputedCells() {
    var active = document.activeElement;
    var needFull = false;
    state.tasks.forEach(function (task) {
      var tr = $("task-body").querySelector('tr[data-id="' + cssEscape(task.id) + '"]');
      if (!tr) {
        needFull = true;
        return;
      }
      tr.classList.toggle("is-group", !!task.isGroup);
      tr.classList.toggle("is-hidden", !isRowVisible(indexById(task.id)));
      tr.classList.toggle("is-selected", state.selectedId === task.id);
      tr.classList.toggle("is-hover", hoverId === task.id);
      tr.classList.toggle("is-critical", !!(task.critical && state.showCritical));
      ["start", "duration", "end", "progress", "allocation"].forEach(function (field) {
        var input = tr.querySelector('[data-field="' + field + '"]');
        if (!input || input === active) return;
        var val = task[field] == null ? "" : String(task[field]);
        if (field === "allocation" && val === "") val = "100";
        if (input.value !== val) input.value = val;
        if (field !== "allocation") input.readOnly = !!task.isGroup;
      });
      var deps = tr.querySelector('[data-field="dependencies"]');
      if (deps && deps !== active) {
        var formatted = S.formatDependencies(task.dependencies);
        if (deps.value !== formatted) deps.value = formatted;
      }
      var ms = tr.querySelector('[data-field="milestone"]');
      if (ms && ms !== active) {
        ms.checked = !!task.milestone;
        ms.disabled = !!task.isGroup;
      }
      var varCell = tr.querySelector(".var-cell");
      if (varCell) varCell.innerHTML = varianceHtml(task);
    });
    if (needFull) renderTable();
  }

  function cssEscape(s) {
    if (window.CSS && CSS.escape) return CSS.escape(s);
    return String(s).replace(/"/g, '\\"');
  }

  function focusCell(id, field) {
    requestAnimationFrame(function () {
      var input = $("task-body").querySelector(
        'tr[data-id="' + cssEscape(id) + '"] [data-field="' + field + '"]'
      );
      if (input) {
        input.focus();
        if (input.select) input.select();
      }
    });
  }

  /* ---- timeline --------------------------------------------------------- */

  function pxPerDay() {
    return PX[state.zoom] || PX.week;
  }

  function computeRange() {
    var start = S.todayISO();
    var end = start;
    state.tasks.forEach(function (t) {
      ["start", "end", "baselineStart", "baselineEnd"].forEach(function (k) {
        if (S.isISO(t[k]) && t[k] < start) start = t[k];
        if (S.isISO(t[k]) && t[k] > end) end = t[k];
      });
    });
    start = S.addDays(start, -7);
    end = S.addDays(end, state.zoom === "month" ? 21 : 14);
    var pane = $("timeline-body-scroll");
    var width = pane && pane.clientWidth ? pane.clientWidth : 800;
    var days = S.diffDays(start, end) + 1;
    var minDays = Math.ceil(width / pxPerDay()) + 2;
    if (days < minDays) end = S.addDays(start, minDays - 1);
    return { start: start, end: end, days: S.diffDays(start, end) + 1 };
  }

  var lastRange = null;

  function dayLeft(range, iso) {
    return S.diffDays(range.start, iso) * pxPerDay();
  }

  function renderTimelineHeader(range, width) {
    var header = $("timeline-header");
    header.style.width = width + "px";
    var hhtml = "";
    var px = pxPerDay();
    var monthStart = range.start;
    var curMonth = monthStart.slice(0, 7);
    S.eachDay(range.start, range.end, function (iso) {
      var key = iso.slice(0, 7);
      if (key !== curMonth) {
        hhtml += bandHtml(range, monthStart, S.addDays(iso, -1), monthLabel(monthStart));
        curMonth = key;
        monthStart = iso;
      }
    });
    hhtml += bandHtml(range, monthStart, range.end, monthLabel(monthStart));

    S.eachDay(range.start, range.end, function (iso, d) {
      var left = dayLeft(range, iso);
      var weekend = d.getDay() === 0 || d.getDay() === 6;
      var today = iso === S.todayISO();
      var dow = ["S", "M", "T", "W", "T", "F", "S"][d.getDay()];
      var label =
        state.zoom === "month"
          ? today || d.getDate() === 1 || d.getDate() % 5 === 0
            ? '<span class="dom">' + d.getDate() + "</span>"
            : ""
          : '<span class="dom">' +
            d.getDate() +
            "</span>" +
            (state.zoom === "day" ? '<span class="dow">' + dow + "</span>" : "");
      hhtml +=
        '<div class="tl-day' +
        (weekend ? " is-weekend" : "") +
        (today ? " is-today" : "") +
        '" style="left:' +
        left +
        "px;width:" +
        px +
        'px">' +
        label +
        "</div>";
    });

    var today = S.todayISO();
    if (today >= range.start && today <= range.end) {
      var tx = dayLeft(range, today) + px / 2;
      hhtml +=
        '<div class="today-line" style="left:' +
        tx +
        'px"></div><div class="today-label" style="left:' +
        tx +
        'px">Today</div>';
    }
    header.innerHTML = hhtml;
  }

  function gridHtml(range, width, height) {
    var px = pxPerDay();
    var chtml = '<div class="tl-grid" style="position:absolute;inset:0;height:' + height + "px;width:" + width + 'px">';
    S.eachDay(range.start, range.end, function (iso, d) {
      var left = dayLeft(range, iso);
      if (d.getDay() === 0 || d.getDay() === 6) {
        chtml += '<div class="tl-weekend" style="left:' + left + "px;width:" + px + 'px"></div>';
      }
      if (state.zoom === "day" || d.getDay() === 1 || d.getDate() === 1) {
        chtml += '<div class="tl-gridline" style="left:' + left + 'px"></div>';
      }
    });
    var today = S.todayISO();
    if (today >= range.start && today <= range.end) {
      chtml += '<div class="today-line" style="left:' + (dayLeft(range, today) + px / 2) + 'px"></div>';
    }
    chtml += "</div>";
    return chtml;
  }

  function renderTimeline() {
    var range = computeRange();
    lastRange = range;
    var px = pxPerDay();
    var width = range.days * px;
    renderTimelineHeader(range, width);

    $("resource-canvas").hidden = state.view !== "resources";
    document.querySelector(".workspace").classList.toggle("is-resources", state.view === "resources");
    document.querySelector(".workspace").classList.toggle("show-critical", !!state.showCritical);

    if (state.view === "resources") {
      renderResources(range, width);
      $("timeline-header-clip").scrollLeft = $("timeline-body-scroll").scrollLeft;
      return;
    }

    var entries = visibleEntries();
    var height = Math.max(entries.length * ROW_H, $("timeline-body-scroll").clientHeight || 200);
    var canvas = $("timeline-canvas");
    var chtml = gridHtml(range, width, height);

    if (!state.tasks.length) {
      chtml +=
        '<div class="empty-state"><strong>No tasks yet</strong><span>Add a task or load the sample project. Bars fill in as you enter start dates and durations.</span></div>';
    }

    entries.forEach(function (entry) {
      var task = entry.task;
      var cls =
        "tl-row" +
        (state.selectedId === task.id ? " is-selected" : "") +
        (hoverId === task.id ? " is-hover" : "");
      chtml += '<div class="' + cls + '" data-id="' + escapeAttr(task.id) + '">';
      if (S.isISO(task.baselineStart) && S.isISO(task.baselineEnd)) {
        var bl = dayLeft(range, task.baselineStart) + 1;
        var bw = Math.max(px - 2, Math.max(1, S.diffDays(task.baselineStart, task.baselineEnd) + 1) * px - 2);
        chtml +=
          '<div class="bar-baseline" style="left:' +
          bl +
          "px;width:" +
          (task.milestone ? 8 : bw) +
          'px"></div>';
      }
      if (S.isISO(task.start) && S.isISO(task.end) && task.end >= task.start) {
        if (task.milestone) {
          var mx = dayLeft(range, task.start) + px / 2;
          chtml +=
            '<div class="milestone-diamond' +
            (task.critical ? " is-critical" : "") +
            (drag && drag.id === task.id ? " is-dragging" : "") +
            '" data-id="' +
            escapeAttr(task.id) +
            '" title="' +
            escapeAttr((task.name || "Milestone #" + task.number) + " · " + task.start) +
            '" style="left:' +
            mx +
            "px;background:" +
            escapeAttr(task.color || PALETTE[0]) +
            '"></div>';
        } else {
          var left = dayLeft(range, task.start) + 1;
          var w = Math.max(px - 2, (S.diffDays(task.start, task.end) + 1) * px - 2);
          var lum = hexLuminance(task.color || PALETTE[0]);
          var labelColor = lum > 0.62 ? "#1c212b" : "#fff";
          chtml +=
            '<div class="bar' +
            (task.isGroup ? " is-group" : "") +
            (task.critical ? " is-critical" : "") +
            (drag && drag.id === task.id ? " is-dragging" : "") +
            '" data-id="' +
            escapeAttr(task.id) +
            '" style="left:' +
            left +
            "px;width:" +
            w +
            "px;background:" +
            escapeAttr(task.color || PALETTE[0]) +
            '">' +
            '<div class="bar-fill" style="width:' +
            S.clampProgress(task.progress) +
            '%"></div>' +
            (task.isGroup
              ? ""
              : '<div class="bar-label" style="color:' +
                labelColor +
                '">' +
                escapeAttr(task.name || "Task #" + task.number) +
                "</div>") +
            (task.isGroup
              ? ""
              : '<div class="bar-grip" data-handle="move" title="Drag to move"></div><div class="bar-handle start" data-handle="start"></div><div class="bar-handle end" data-handle="end"></div>') +
            "</div>";
        }
      }
      chtml += "</div>";
    });

    chtml +=
      '<svg class="arrow-layer" width="' +
      width +
      '" height="' +
      height +
      '" viewBox="0 0 ' +
      width +
      " " +
      height +
      '">';
    var visIndex = {};
    entries.forEach(function (e, i) {
      visIndex[e.task.number] = i;
    });
    var cal = calendar();
    entries.forEach(function (entry) {
      var task = entry.task;
      S.normalizeDeps(task.dependencies).forEach(function (dep) {
        var pred = S.findTaskByNumber(state.tasks, dep.pred);
        if (!pred || !S.isISO(pred.start) || !S.isISO(pred.end) || !S.isISO(task.start) || !S.isISO(task.end)) return;
        if (visIndex[pred.number] == null || visIndex[task.number] == null) return;
        var y1 = visIndex[pred.number] * ROW_H + ROW_H / 2;
        var y2 = visIndex[task.number] * ROW_H + ROW_H / 2;
        var type = dep.type || "FS";
        var predX = type === "SS" || type === "SF" ? pred.start : pred.end;
        var succX = type === "FF" || type === "SF" ? task.end : task.start;
        var x1, x2;
        if (type === "SS" || type === "SF") x1 = dayLeft(range, pred.start) + (pred.milestone ? px / 2 : 1);
        else x1 = dayLeft(range, pred.end) + (pred.milestone ? px / 2 : px - 1);
        if (type === "FF" || type === "SF") x2 = dayLeft(range, task.end) + (task.milestone ? px / 2 : px - 1);
        else x2 = dayLeft(range, task.start) + (task.milestone ? px / 2 : 1);
        void predX;
        void succX;
        var crit = S.isCriticalDep(pred, task, dep, cal);
        chtml += arrowPath(x1, y1, x2, y2, crit);
      });
    });
    chtml += "</svg>";

    canvas.style.width = width + "px";
    canvas.style.height = height + "px";
    canvas.innerHTML = chtml;
    $("timeline-header-clip").scrollLeft = $("timeline-body-scroll").scrollLeft;
  }

  function loadColor(pct) {
    // 0 = empty, 1-50 muted teal, 51-100 solid teal, >100 rose
    if (pct <= 0) return "transparent";
    if (pct > 100) {
      var over = Math.min(1, (pct - 100) / 100);
      return "rgba(225, 29, 72," + (0.45 + over * 0.4).toFixed(2) + ")";
    }
    var a = 0.18 + (pct / 100) * 0.55;
    return "rgba(15, 118, 110," + a.toFixed(2) + ")";
  }

  function renderResources(range, width) {
    var canvas = $("resource-canvas");
    var cal = calendar();
    var rows = S.resourceUsage(state.tasks, range.start, range.end, cal);
    var height = Math.max(rows.length * RES_ROW_H, $("timeline-body-scroll").clientHeight || 200);
    var px = pxPerDay();
    var html = gridHtml(range, width, height);
    if (!rows.length) {
      html +=
        '<div class="res-empty">Assign people to tasks to see their workload. Days over 100% are flagged as overallocated.</div>';
    }
    rows.forEach(function (row, i) {
      html +=
        '<div class="res-row" data-assignee="' +
        escapeAttr(row.name) +
        '"><div class="res-label' +
        (row.overallocated ? " is-over" : "") +
        '"><strong>' +
        escapeAttr(row.name) +
        "</strong><span>" +
        (row.overallocated
          ? row.overDays.length + " day" + (row.overDays.length === 1 ? "" : "s") + " over 100%"
          : "Peak " + row.peak + "%") +
        "</span></div>";
      Object.keys(row.days).forEach(function (iso) {
        if (iso < range.start || iso > range.end) return;
        var left = dayLeft(range, iso);
        var pct = row.days[iso];
        html +=
          '<div class="res-cell' +
          (pct > 100 ? " is-over" : "") +
          '" title="' +
          escapeAttr(row.name + " · " + iso + " · " + pct + "%") +
          '" style="left:' +
          left +
          "px;width:" +
          Math.max(2, px - 1) +
          "px;background:" +
          loadColor(pct) +
          '"></div>';
      });
      html += "</div>";
    });
    canvas.style.width = width + "px";
    canvas.style.height = height + "px";
    canvas.innerHTML = html;
  }

  function bandHtml(range, start, end, label) {
    var left = dayLeft(range, start);
    var width = (S.diffDays(start, end) + 1) * pxPerDay();
    return (
      '<div class="tl-band" style="left:' +
      left +
      "px;width:" +
      width +
      'px">' +
      escapeAttr(label) +
      "</div>"
    );
  }

  function monthLabel(iso) {
    var d = S.parseISO(iso);
    return d.toLocaleString(undefined, { month: "short", year: "numeric" });
  }

  function arrowPath(x1, y1, x2, y2, critical) {
    var dx = Math.max(16, Math.abs(x2 - x1) / 2);
    var c1x = x1 + dx;
    var c2x = x2 - dx;
    var tip = 5;
    var cls = critical && state.showCritical ? ' class="is-critical"' : "";
    var path = "M" + x1 + " " + y1 + " C " + c1x + " " + y1 + ", " + c2x + " " + y2 + ", " + x2 + " " + y2;
    return (
      "<path" +
      cls +
      ' d="' +
      path +
      '" />' +
      "<polygon" +
      cls +
      ' points="' +
      x2 +
      "," +
      y2 +
      " " +
      (x2 - tip) +
      "," +
      (y2 - tip) +
      " " +
      (x2 - tip) +
      "," +
      (y2 + tip) +
      '" />'
    );
  }

  function hexLuminance(hex) {
    var h = (hex || "").replace("#", "");
    if (h.length !== 6) return 0.3;
    var r = parseInt(h.slice(0, 2), 16) / 255;
    var g = parseInt(h.slice(2, 4), 16) / 255;
    var b = parseInt(h.slice(4, 6), 16) / 255;
    var lin = function (c) {
      return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
    };
    return 0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b);
  }

  function updateStatus() {
    var starts = state.tasks.map(function (t) { return t.start; }).filter(S.isISO);
    var ends = state.tasks.map(function (t) { return t.end; }).filter(S.isISO);
    var el = $("status-range");
    if (!starts.length) el.textContent = "";
    else {
      var a = S.minISO.apply(null, starts);
      var b = S.maxISO.apply(null, ends);
      el.textContent = formatLong(a) + " – " + formatLong(b);
    }
    var cal = $("status-calendar");
    var bits = [state.calendarMode === "workdays" ? "Workdays (skip weekends & holidays)" : "Calendar days"];
    if (state.calendarMode === "workdays" && state.holidays.length) {
      bits.push(state.holidays.length + " holiday" + (state.holidays.length === 1 ? "" : "s"));
    }
    if (state.baselineSetAt) bits.push("Baseline set");
    var critN = state.tasks.filter(function (t) { return t.critical && !t.isGroup; }).length;
    if (state.showCritical) bits.push(critN + " critical");
    cal.textContent = bits.join(" · ");
  }

  function formatLong(iso) {
    var d = S.parseISO(iso);
    return d ? d.toLocaleDateString(undefined, { month: "short", day: "numeric", year: "numeric" }) : iso;
  }

  function isMobileLayout() {
    return window.innerWidth <= MOBILE_BP;
  }

  function applyMobilePane() {
    var ws = $("workspace") || document.querySelector(".workspace");
    if (!ws) return;
    var pane = state.mobilePane || "timeline";
    ws.classList.toggle("is-mobile-table", pane === "table");
    ws.classList.toggle("is-mobile-timeline", pane === "timeline");
    ws.classList.toggle("is-mobile-resources", pane === "resources");
    document.querySelectorAll("[data-pane]").forEach(function (btn) {
      btn.setAttribute("aria-pressed", btn.getAttribute("data-pane") === pane ? "true" : "false");
    });
    if (pane === "resources" && state.view !== "resources") {
      state.view = "resources";
    } else if (pane === "timeline" && state.view !== "timeline") {
      state.view = "timeline";
    }
  }

  function setMobilePane(pane) {
    if (pane !== "table" && pane !== "timeline" && pane !== "resources") return;
    state.mobilePane = pane;
    if (pane === "resources") state.view = "resources";
    if (pane === "timeline") state.view = "timeline";
    applyMobilePane();
    renderZoom();
    renderTimeline();
    persist();
  }

  function renderZoom() {
    document.querySelectorAll("[data-zoom]").forEach(function (btn) {
      var on = btn.getAttribute("data-zoom") === state.zoom;
      btn.setAttribute("aria-pressed", on ? "true" : "false");
      if (btn.getAttribute("role") === "menuitemradio") btn.setAttribute("aria-checked", on ? "true" : "false");
      btn.classList.toggle("is-active", on);
    });
    document.querySelectorAll("[data-cal]").forEach(function (btn) {
      var on = btn.getAttribute("data-cal") === state.calendarMode;
      btn.setAttribute("aria-pressed", on ? "true" : "false");
      if (btn.getAttribute("role") === "menuitemradio") btn.setAttribute("aria-checked", on ? "true" : "false");
      btn.classList.toggle("is-active", on);
    });
    document.querySelectorAll("[data-view]").forEach(function (btn) {
      var on = btn.getAttribute("data-view") === state.view;
      btn.setAttribute("aria-pressed", on ? "true" : "false");
      if (btn.getAttribute("role") === "menuitemradio") btn.setAttribute("aria-checked", on ? "true" : "false");
      btn.classList.toggle("is-active", on);
    });
    var crit = $("chk-critical");
    if (crit) {
      crit.setAttribute("aria-checked", state.showCritical ? "true" : "false");
      crit.classList.toggle("is-active", !!state.showCritical);
    }
    document.querySelectorAll("[data-critical-item]").forEach(function (el) {
      el.setAttribute("aria-checked", state.showCritical ? "true" : "false");
      el.classList.toggle("is-active", !!state.showCritical);
    });
    applyMobilePane();
  }

  function renderAll() {
    $("project-title").value = state.projectName;
    applySplit();
    applyMobilePane();
    renderZoom();
    renderTable();
    renderTimeline();
    updateStatus();
    updateUndoButtons();
    renderHolidayList();
    updateInstallButtons();
  }

  /* ---- interactions ----------------------------------------------------- */

  function onTableInput(e) {
    var input = e.target.closest("[data-field]");
    if (!input) return;
    var tr = input.closest("tr");
    if (!tr) return;
    maybeCheckpoint();
    var val = input.type === "checkbox" ? input.checked : input.value;
    onFieldInput(tr.getAttribute("data-id"), input.getAttribute("data-field"), val);
  }

  function maybeCheckpoint() {
    if (!typingCheckpoint) {
      pushHistory();
      typingCheckpoint = true;
    }
  }

  function onTableFocusOut() {
    typingCheckpoint = false;
  }

  function onTableClick(e) {
    var tr = e.target.closest("tr");
    if (tr && tr.dataset.id) {
      state.selectedId = tr.dataset.id;
      syncComputedCells();
      renderTimeline();
    }
    var btn = e.target.closest("[data-act]");
    if (!btn || btn.classList.contains("drag-handle")) return;
    var row = btn.closest("tr");
    if (!row) return;
    var id = row.dataset.id;
    var idx = indexById(id);
    if (idx < 0) return;
    var act = btn.getAttribute("data-act");
    if (act === "delete") deleteTask(id);
    else if (act === "up") moveBlockUp(idx);
    else if (act === "down") moveBlockDown(idx);
    else if (act === "indent") indentBlock(idx);
    else if (act === "outdent") outdentBlock(idx);
    else if (act === "collapse") toggleCollapse(idx);
  }

  function onTablePointerOver(e) {
    var tr = e.target.closest("tr");
    var id = tr && tr.dataset.id;
    if (hoverId === id) return;
    hoverId = id || null;
    syncRowHover();
  }

  function syncRowHover() {
    $("task-body").querySelectorAll("tr").forEach(function (tr) {
      tr.classList.toggle("is-hover", tr.dataset.id === hoverId);
    });
    $("timeline-canvas").querySelectorAll(".tl-row").forEach(function (row) {
      row.classList.toggle("is-hover", row.dataset.id === hoverId);
    });
  }

  function onGrabDragStart(e) {
    var handle = e.target.closest("[data-act=grab]");
    if (!handle) return;
    var tr = handle.closest("tr");
    var id = tr.dataset.id;
    var idx = indexById(id);
    rowDrag = { id: id, index: idx, block: S.getBlockIndexes(state.tasks, idx) };
    e.dataTransfer.effectAllowed = "move";
    e.dataTransfer.setData("text/plain", id);
    tr.style.opacity = "0.5";
  }

  function onGrabDragEnd() {
    $("task-body").querySelectorAll("tr").forEach(function (tr) {
      tr.style.opacity = "";
      tr.classList.remove("drop-line");
      tr.classList.remove("drop-after");
    });
    rowDrag = null;
  }

  function dropInsertIndex(tr, ev) {
    var id = tr.dataset.id;
    var idx = indexById(id);
    if (idx < 0) return -1;
    var rect = tr.getBoundingClientRect();
    var after = ev.clientY > rect.top + rect.height / 2;
    if (after) {
      var block = S.getBlockIndexes(state.tasks, idx);
      return block[block.length - 1] + 1;
    }
    return idx;
  }

  function onTableDragOver(e) {
    if (!rowDrag) return;
    var tr = e.target.closest("tr");
    if (!tr) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = "move";
    $("task-body").querySelectorAll("tr").forEach(function (r) {
      r.classList.remove("drop-line");
    });
    var insertAt = dropInsertIndex(tr, e);
    var block = rowDrag.block;
    if (insertAt > rowDrag.index && insertAt <= block[block.length - 1] + 1) return;
    var rect = tr.getBoundingClientRect();
    var after = ev.clientY > rect.top + rect.height / 2;
    if (after) {
      var lastId = state.tasks[S.getBlockIndexes(state.tasks, indexById(tr.dataset.id)).pop()].id;
      var lastTr = $("task-body").querySelector('tr[data-id="' + cssEscape(lastId) + '"]');
      if (lastTr) lastTr.classList.add("drop-line");
    } else {
      tr.classList.add("drop-line");
    }
  }

  function onTableDrop(e) {
    var tr = e.target.closest("tr");
    if (!tr || !rowDrag) return;
    e.preventDefault();
    var insertAt = dropInsertIndex(tr, e);
    var from = indexById(rowDrag.id);
    if (from < 0) return;
    pushHistory();
    if (moveBlock(from, insertAt)) {
      normalizeIndents();
      runSchedule();
      renderAll();
      persist();
    } else {
      historyStack.pop();
      updateUndoButtons();
    }
  }

  function startBarDrag(task, mode, clientX) {
    pushHistory();
    drag = {
      id: task.id,
      mode: mode,
      startX: clientX,
      origStart: task.start,
      origEnd: task.end,
      origDuration: task.duration,
      snapshot: state.tasks.map(function (t) {
        return { id: t.id, start: t.start, end: t.end, duration: t.duration };
      }),
    };
    window.addEventListener("pointermove", onBarPointerMove);
    window.addEventListener("pointerup", onBarPointerUp);
    window.addEventListener("pointercancel", onBarPointerUp);
  }

  function clearPressTimer() {
    if (pressTimer) {
      clearTimeout(pressTimer);
      pressTimer = null;
    }
  }

  function onCanvasPointerDown(e) {
    var handle = e.target.closest(".bar-handle, .bar-grip");
    var bar = e.target.closest(".bar, .milestone-diamond");
    if (!bar) {
      var row = e.target.closest(".tl-row");
      if (row) {
        state.selectedId = row.dataset.id;
        syncComputedCells();
        renderTimeline();
      }
      return;
    }
    var task = getTask(bar.dataset.id);
    if (!task) return;
    state.selectedId = task.id;
    if (task.isGroup && handle && handle.classList.contains("bar-handle")) return;
    if (task.milestone && handle && handle.classList.contains("bar-handle")) return;

    var mode = "move";
    if (handle) {
      var kind = handle.getAttribute("data-handle");
      if (kind === "start") mode = "resize-start";
      else if (kind === "end") mode = "resize-end";
      else mode = "move";
    }

    var isTouch = e.pointerType === "touch";
    var needsPress = isTouch && !handle && !task.isGroup;
    if (!needsPress) {
      e.preventDefault();
      startBarDrag(task, mode, e.clientX);
      return;
    }

    // Long-press to drag, so a swipe on a bar can still scroll the timeline.
    var startX = e.clientX;
    var startY = e.clientY;
    var pointerId = e.pointerId;
    clearPressTimer();
    function onMove(ev) {
      if (ev.pointerId !== pointerId) return;
      if (Math.hypot(ev.clientX - startX, ev.clientY - startY) > 8) {
        cancelPress();
      }
    }
    function cancelPress() {
      clearPressTimer();
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", cancelPress);
      window.removeEventListener("pointercancel", cancelPress);
    }
    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", cancelPress);
    window.addEventListener("pointercancel", cancelPress);
    pressTimer = setTimeout(function () {
      cancelPress();
      startBarDrag(task, "move", startX);
    }, 380);
  }

  function onBarPointerMove(e) {
    if (!drag || !lastRange) return;
    var delta = Math.round((e.clientX - drag.startX) / pxPerDay());
    drag.snapshot.forEach(function (s) {
      var t = getTask(s.id);
      if (!t) return;
      t.start = s.start;
      t.end = s.end;
      t.duration = s.duration;
    });
    var task = getTask(drag.id);
    var idx = indexById(drag.id);
    var cal = calendar();
    if (drag.mode === "move") {
      S.shiftBlock(state.tasks, idx, delta);
      if (!task.isGroup && S.isISO(task.start)) {
        task.end = S.endFromStartDuration(task.start, S.taskDuration(task), cal);
      }
    } else if (drag.mode === "resize-end") {
      var newDur = Math.max(1, (drag.origDuration || 1) + delta);
      if (state.calendarMode === "workdays") {
        var tentativeEnd = S.addDays(drag.origEnd, delta);
        newDur = Math.max(1, S.inclusiveDuration(task.start, tentativeEnd, cal) || 1);
      }
      task.duration = newDur;
      task.milestone = false;
      task.end = S.endFromStartDuration(task.start, newDur, cal);
    } else if (drag.mode === "resize-start") {
      var newStart = S.addDays(drag.origStart, delta);
      var nd = S.inclusiveDuration(newStart, drag.origEnd, cal);
      if (!nd || nd < 1) {
        task.start = drag.origEnd;
        task.duration = 1;
        task.end = drag.origEnd;
      } else {
        task.start = newStart;
        task.duration = nd;
        task.end = drag.origEnd;
      }
      task.milestone = false;
    }
    runSchedule();
    syncComputedCells();
    renderTimeline();
  }

  function onBarPointerUp() {
    window.removeEventListener("pointermove", onBarPointerMove);
    window.removeEventListener("pointerup", onBarPointerUp);
    window.removeEventListener("pointercancel", onBarPointerUp);
    if (!drag) return;
    var changed =
      JSON.stringify(drag.snapshot) !==
      JSON.stringify(
        state.tasks.map(function (t) {
          return { id: t.id, start: t.start, end: t.end, duration: t.duration };
        })
      );
    drag = null;
    if (!changed) {
      historyStack.pop();
      updateUndoButtons();
    }
    persist();
    renderTimeline();
  }

  var syncingScroll = false;
  function onTimelineScroll() {
    $("timeline-header-clip").scrollLeft = $("timeline-body-scroll").scrollLeft;
    if (!syncingScroll && state.view === "timeline") {
      syncingScroll = true;
      $("table-scroll").scrollTop = $("timeline-body-scroll").scrollTop;
      syncingScroll = false;
    }
  }
  function onTableScroll() {
    if (syncingScroll || state.view !== "timeline") return;
    syncingScroll = true;
    $("timeline-body-scroll").scrollTop = $("table-scroll").scrollTop;
    syncingScroll = false;
  }

  function applySplit() {
    if (isMobileLayout()) {
      $("task-pane").style.width = "";
      return;
    }
    $("task-pane").style.width = state.split + "px";
  }

  function initSplitter() {
    var el = $("splitter");
    var dragging = false;
    var startX = 0;
    var startW = 0;
    el.addEventListener("pointerdown", function (e) {
      dragging = true;
      startX = e.clientX;
      startW = $("task-pane").getBoundingClientRect().width;
      el.classList.add("is-dragging");
      el.setPointerCapture(e.pointerId);
      e.preventDefault();
    });
    el.addEventListener("pointermove", function (e) {
      if (!dragging) return;
      var w = Math.min(Math.max(280, startW + (e.clientX - startX)), Math.floor(window.innerWidth * 0.8));
      state.split = w;
      applySplit();
      renderTimeline();
    });
    function end() {
      if (!dragging) return;
      dragging = false;
      el.classList.remove("is-dragging");
      persist();
    }
    el.addEventListener("pointerup", end);
    el.addEventListener("pointercancel", end);
    el.addEventListener("keydown", function (e) {
      var step = e.shiftKey ? 40 : 16;
      if (e.key === "ArrowLeft") {
        state.split = Math.max(280, state.split - step);
        applySplit();
        renderTimeline();
        persist();
      } else if (e.key === "ArrowRight") {
        state.split = Math.min(Math.floor(window.innerWidth * 0.8), state.split + step);
        applySplit();
        renderTimeline();
        persist();
      }
    });
  }

  /* ---- export / import -------------------------------------------------- */

  function exportJSON() {
    var data = persistPayload();
    data.exportedAt = new Date().toISOString();
    data.tasks = state.tasks.map(function (t) {
      return {
        number: t.number,
        name: t.name,
        start: t.start,
        end: t.end,
        duration: t.duration,
        progress: t.progress,
        dependencies: S.normalizeDeps(t.dependencies),
        assignee: t.assignee,
        allocation: t.allocation,
        milestone: !!t.milestone,
        baselineStart: t.baselineStart || null,
        baselineEnd: t.baselineEnd || null,
        variance: t.variance,
        critical: !!t.critical,
        color: t.color,
        indent: t.indent,
        collapsed: !!t.collapsed,
      };
    });
    downloadFile(slug(state.projectName) + ".json", "application/json", JSON.stringify(data, null, 2));
  }

  function currentProject() {
    return {
      projectName: state.projectName,
      calendarMode: state.calendarMode,
      holidays: state.holidays,
      zoom: state.zoom,
      today: S.todayISO(),
      tasks: state.tasks,
    };
  }

  function exportExcelBuffer() {
    if (!XLSX || !Excel) throw new Error("Excel library failed to load.");
    return Excel.writeBuffer(currentProject(), XLSX, S);
  }

  function exportExcel() {
    try {
      var buf = exportExcelBuffer();
      var blob = new Blob([buf], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      downloadFile(slug(state.projectName) + ".xlsx", blob.type, blob);
      toast("Downloaded Excel workbook");
    } catch (err) {
      toast("Excel export failed: " + err.message);
    }
  }

  function slug(s) {
    return (
      String(s || "gantt-project")
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, "-")
        .replace(/^-|-$/g, "") || "gantt-project"
    );
  }

  function downloadFile(name, type, contents) {
    var blob = contents instanceof Blob ? contents : new Blob([contents], { type: type });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(function () {
      URL.revokeObjectURL(url);
    }, 1000);
  }

  function importFile(file) {
    var name = (file.name || "").toLowerCase();
    var reader = new FileReader();
    if (name.endsWith(".csv") || file.type === "text/csv") {
      reader.onload = function () {
        try {
          var data = Excel.parseCSV(String(reader.result), S);
          pushHistory();
          if (!data.projectName) data.projectName = state.projectName;
          applyProject(data);
          toast("Imported CSV");
        } catch (err) {
          toast("Import failed: " + err.message);
        }
      };
      reader.readAsText(file);
      return;
    }
    if (name.endsWith(".xlsx") || name.endsWith(".xls") || /sheet|excel/.test(file.type || "")) {
      reader.onload = function () {
        try {
          var data = Excel.parseWorkbook(new Uint8Array(reader.result), XLSX, S);
          pushHistory();
          if (!data.projectName) data.projectName = slug(file.name.replace(/\.(xlsx|xls)$/i, "")) || state.projectName;
          applyProject(data);
          toast("Imported Excel workbook");
        } catch (err) {
          toast("Import failed: " + err.message);
        }
      };
      reader.readAsArrayBuffer(file);
      return;
    }
    reader.onload = function () {
      try {
        var data = JSON.parse(String(reader.result));
        pushHistory();
        applyProject(data);
        toast("Imported project");
      } catch (err) {
        toast("Import failed: " + err.message);
      }
    };
    reader.readAsText(file);
  }

  /* PNG export (custom canvas; includes baseline + milestones) */
  function exportPNG() {
    var range = lastRange || computeRange();
    var px = Math.max(10, Math.min(pxPerDay(), 22));
    var days = range.days;
    var col = {
      id: 40,
      name: 200,
      start: 90,
      dur: 44,
      end: 90,
      prog: 40,
      deps: 72,
      who: 90,
    };
    var tableW = col.id + col.name + col.start + col.dur + col.end + col.prog + col.deps + col.who;
    var pad = 28;
    var headerH = 64;
    var tlHead = 44;
    var rowH = 28;
    var rows = state.tasks;
    var bodyH = Math.max(rows.length * rowH, 80);
    var tlW = days * px;
    var width = pad * 2 + tableW + 8 + tlW;
    var height = pad * 2 + headerH + tlHead + bodyH + 16;
    var canvas = document.createElement("canvas");
    var scale = 2;
    canvas.width = width * scale;
    canvas.height = height * scale;
    var ctx = canvas.getContext("2d");
    ctx.scale(scale, scale);

    var dark = document.documentElement.getAttribute("data-theme") === "dark";
    var bg = dark ? "#10141b" : "#ffffff";
    var elev = dark ? "#181d27" : "#f8f9fb";
    var text = dark ? "#e8eaef" : "#1c212b";
    var soft = dark ? "#9aa3b5" : "#5b6372";
    var faint = dark ? "#6d7689" : "#8b93a2";
    var line = dark ? "#2a3140" : "#e1e4ea";
    var weekend = dark ? "rgba(255,255,255,0.04)" : "rgba(28,33,43,0.05)";
    var todayC = "#e11d48";

    ctx.fillStyle = bg;
    ctx.fillRect(0, 0, width, height);
    ctx.fillStyle = text;
    ctx.font = "650 20px Instrument Sans, sans-serif";
    ctx.fillText(state.projectName || "Untitled project", pad, pad + 22);
    ctx.font = "12px Instrument Sans, sans-serif";
    ctx.fillStyle = soft;
    var sub = "Gantt  ·  Prerelease";
    var starts = rows.map(function (t) { return t.start; }).filter(S.isISO);
    var ends = rows.map(function (t) { return t.end; }).filter(S.isISO);
    if (starts.length) sub += "  ·  " + formatLong(S.minISO.apply(null, starts)) + " – " + formatLong(S.maxISO.apply(null, ends));
    if (state.calendarMode === "workdays") sub += "  ·  Workdays";
    ctx.fillText(sub, pad, pad + 42);

    var ox = pad;
    var oy = pad + headerH;
    ctx.fillStyle = elev;
    ctx.fillRect(ox, oy, tableW + 8 + tlW, tlHead + bodyH);
    ctx.strokeStyle = line;
    ctx.strokeRect(ox + 0.5, oy + 0.5, tableW + 8 + tlW - 1, tlHead + bodyH - 1);

    function colX(name) {
      var order = ["id", "name", "start", "dur", "end", "prog", "deps", "who"];
      var x = ox;
      for (var i = 0; i < order.length; i++) {
        if (order[i] === name) return x;
        x += col[order[i]];
      }
      return x;
    }

    ctx.font = "600 10px Instrument Sans, sans-serif";
    ctx.fillStyle = faint;
    [
      ["id", "ID"],
      ["name", "TASK"],
      ["start", "START"],
      ["dur", state.calendarMode === "workdays" ? "WD" : "DAYS"],
      ["end", "END"],
      ["prog", "%"],
      ["deps", "DEPS"],
      ["who", "ASSIGNEE"],
    ].forEach(function (h) {
      ctx.fillText(h[1], colX(h[0]) + 8, oy + 18);
    });

    var tlX = ox + tableW + 8;
    var monthStart = range.start;
    var curMonth = monthStart.slice(0, 7);
    ctx.font = "600 10px Instrument Sans, sans-serif";
    ctx.fillStyle = soft;
    function flushBand(from) {
      var x = tlX + S.diffDays(range.start, from) * px;
      ctx.fillText(monthLabel(from), x + 6, oy + 16);
      ctx.strokeStyle = line;
      ctx.beginPath();
      ctx.moveTo(x + 0.5, oy);
      ctx.lineTo(x + 0.5, oy + tlHead + bodyH);
      ctx.stroke();
    }
    S.eachDay(range.start, range.end, function (iso) {
      var key = iso.slice(0, 7);
      if (key !== curMonth) {
        flushBand(monthStart);
        curMonth = key;
        monthStart = iso;
      }
    });
    flushBand(monthStart);

    var today = S.todayISO();
    S.eachDay(range.start, range.end, function (iso, d) {
      var x = tlX + S.diffDays(range.start, iso) * px;
      if (d.getDay() === 0 || d.getDay() === 6) {
        ctx.fillStyle = weekend;
        ctx.fillRect(x, oy + tlHead, px, bodyH);
      }
      if (state.zoom !== "month" || d.getDate() === 1 || d.getDate() % 5 === 0 || iso === today) {
        ctx.fillStyle = iso === today ? todayC : faint;
        ctx.font = "10px IBM Plex Mono, monospace";
        ctx.fillText(String(d.getDate()), x + 2, oy + 34);
      }
    });
    if (today >= range.start && today <= range.end) {
      var tx = tlX + S.diffDays(range.start, today) * px + px / 2;
      ctx.strokeStyle = todayC;
      ctx.lineWidth = 1.5;
      ctx.beginPath();
      ctx.moveTo(tx, oy);
      ctx.lineTo(tx, oy + tlHead + bodyH);
      ctx.stroke();
      ctx.lineWidth = 1;
    }

    rows.forEach(function (task, i) {
      var y = oy + tlHead + i * rowH;
      ctx.strokeStyle = line;
      ctx.beginPath();
      ctx.moveTo(ox, y + 0.5);
      ctx.lineTo(ox + tableW + 8 + tlW, y + 0.5);
      ctx.stroke();
      ctx.fillStyle = faint;
      ctx.font = "11px IBM Plex Mono, monospace";
      ctx.fillText("#" + task.number, colX("id") + 8, y + 18);
      ctx.fillStyle = text;
      ctx.font = (task.isGroup ? "650 " : "500 ") + "12px Instrument Sans, sans-serif";
      var name = (task.indent ? "  ".repeat(task.indent) : "") + (task.name || "Untitled") + (task.milestone ? "  ◆" : "");
      ctx.save();
      ctx.beginPath();
      ctx.rect(colX("name") + 6, y, col.name - 10, rowH);
      ctx.clip();
      ctx.fillText(name, colX("name") + 8, y + 18);
      ctx.restore();
      ctx.font = "11px IBM Plex Mono, monospace";
      ctx.fillStyle = soft;
      ctx.fillText(task.start || "", colX("start") + 8, y + 18);
      ctx.fillText(String(task.duration == null ? "" : task.duration), colX("dur") + 8, y + 18);
      ctx.fillText(task.end || "", colX("end") + 8, y + 18);
      ctx.fillText(String(task.progress || 0), colX("prog") + 8, y + 18);
      ctx.fillText(S.formatDependencies(task.dependencies), colX("deps") + 8, y + 18);
      ctx.font = "12px Instrument Sans, sans-serif";
      ctx.fillText(task.assignee || "", colX("who") + 8, y + 18);

      if (S.isISO(task.baselineStart) && S.isISO(task.baselineEnd)) {
        var bx = tlX + S.diffDays(range.start, task.baselineStart) * px + 1;
        var bw = Math.max(3, (S.diffDays(task.baselineStart, task.baselineEnd) + 1) * px - 2);
        ctx.fillStyle = faint;
        ctx.globalAlpha = 0.35;
        ctx.fillRect(bx, y + 20, task.milestone ? 6 : bw, 4);
        ctx.globalAlpha = 1;
      }
      if (S.isISO(task.start) && S.isISO(task.end)) {
        var x = tlX + S.diffDays(range.start, task.start) * px + 1;
        var w = Math.max(px - 2, (S.diffDays(task.start, task.end) + 1) * px - 2);
        if (task.milestone) {
          var mx = tlX + S.diffDays(range.start, task.start) * px + px / 2;
          var my = y + rowH / 2;
          ctx.save();
          ctx.translate(mx, my);
          ctx.rotate(Math.PI / 4);
          ctx.fillStyle = task.color || PALETTE[0];
          ctx.fillRect(-5, -5, 10, 10);
          if (task.critical && state.showCritical) {
            ctx.strokeStyle = todayC;
            ctx.lineWidth = 2;
            ctx.strokeRect(-5, -5, 10, 10);
          }
          ctx.restore();
        } else {
          var bh = task.isGroup ? 10 : 14;
          var by = y + (rowH - bh) / 2 - 2;
          roundRect(ctx, x, by, w, bh, 4);
          ctx.fillStyle = task.color || PALETTE[0];
          ctx.fill();
          if (task.critical && state.showCritical) {
            ctx.strokeStyle = todayC;
            ctx.lineWidth = 1.5;
            ctx.stroke();
            ctx.lineWidth = 1;
          }
          if (task.progress > 0) {
            ctx.save();
            ctx.beginPath();
            roundRect(ctx, x, by, w, bh, 4);
            ctx.clip();
            ctx.fillStyle = "rgba(255,255,255,0.28)";
            ctx.fillRect(x, by, w * (S.clampProgress(task.progress) / 100), bh);
            ctx.restore();
          }
        }
      }
    });

    ctx.strokeStyle = faint;
    ctx.fillStyle = faint;
    ctx.lineWidth = 1.2;
    rows.forEach(function (task, i) {
      S.normalizeDeps(task.dependencies).forEach(function (dep) {
        var pred = S.findTaskByNumber(rows, dep.pred);
        if (!pred || !S.isISO(pred.end) || !S.isISO(task.start)) return;
        var pi = rows.indexOf(pred);
        var y1 = oy + tlHead + pi * rowH + rowH / 2;
        var y2 = oy + tlHead + i * rowH + rowH / 2;
        var type = dep.type || "FS";
        var x1 =
          type === "SS" || type === "SF"
            ? tlX + S.diffDays(range.start, pred.start) * px + 1
            : tlX + S.diffDays(range.start, pred.end) * px + px - 1;
        var x2 =
          type === "FF" || type === "SF"
            ? tlX + S.diffDays(range.start, task.end) * px + px - 1
            : tlX + S.diffDays(range.start, task.start) * px + 1;
        var dx = Math.max(12, Math.abs(x2 - x1) / 2);
        if (task.critical && pred.critical && state.showCritical) {
          ctx.strokeStyle = todayC;
        } else ctx.strokeStyle = faint;
        ctx.beginPath();
        ctx.moveTo(x1, y1);
        ctx.bezierCurveTo(x1 + dx, y1, x2 - dx, y2, x2, y2);
        ctx.stroke();
      });
    });

    canvas.toBlob(function (blob) {
      if (!blob) {
        toast("Could not build the PNG");
        return;
      }
      downloadFile(slug(state.projectName) + ".png", "image/png", blob);
      toast("Downloaded PNG");
    }, "image/png");
  }

  function roundRect(ctx, x, y, w, h, r) {
    var rr = Math.min(r, w / 2, h / 2);
    ctx.beginPath();
    ctx.moveTo(x + rr, y);
    ctx.arcTo(x + w, y, x + w, y + h, rr);
    ctx.arcTo(x + w, y + h, x, y + h, rr);
    ctx.arcTo(x, y + h, x, y, rr);
    ctx.arcTo(x, y, x + w, y, rr);
    ctx.closePath();
  }

  /* ---- holidays + theme ------------------------------------------------- */

  function renderHolidayList() {
    var list = $("holiday-list");
    if (!list) return;
    if (!state.holidays.length) {
      list.innerHTML = '<li class="empty">No holidays yet.</li>';
      return;
    }
    list.innerHTML = state.holidays
      .map(function (iso) {
        return (
          "<li><span>" +
          escapeAttr(iso) +
          ' · ' +
          escapeAttr(formatLong(iso)) +
          '</span><button type="button" class="btn icon danger" data-remove-holiday="' +
          escapeAttr(iso) +
          '" aria-label="Remove ' +
          escapeAttr(iso) +
          '">&times;</button></li>'
        );
      })
      .join("");
  }

  function addHoliday(iso) {
    if (!S.isISO(iso)) return;
    if (state.holidays.indexOf(iso) >= 0) return;
    pushHistory();
    state.holidays = S.uniqueSortedDates(state.holidays.concat([iso]));
    runSchedule();
    renderAll();
    persist();
  }

  function removeHoliday(iso) {
    pushHistory();
    state.holidays = state.holidays.filter(function (d) { return d !== iso; });
    runSchedule();
    renderAll();
    persist();
  }

  function currentTheme() {
    return document.documentElement.getAttribute("data-theme") || "light";
  }

  function applyTheme(theme) {
    document.documentElement.setAttribute("data-theme", theme);
    var meta = document.querySelector('meta[name="theme-color"]');
    if (meta) meta.setAttribute("content", theme === "dark" ? "#10141b" : "#0f766e");
    try {
      localStorage.setItem(THEME_KEY, theme);
    } catch (e) {}
  }

  function initTheme() {
    var saved = null;
    try {
      saved = localStorage.getItem(THEME_KEY);
    } catch (e) {}
    if (saved === "light" || saved === "dark") applyTheme(saved);
    else applyTheme(window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light");
  }

  /* ---- menus, overflow sheet, PWA -------------------------------------- */

  function closeMenus() {
    document.querySelectorAll(".menu").forEach(function (menu) {
      var pop = menu.querySelector(".menu-pop");
      var trigger = menu.querySelector(".menu-trigger");
      if (pop) pop.hidden = true;
      if (trigger) trigger.setAttribute("aria-expanded", "false");
    });
  }

  function openMenu(name) {
    closeMenus();
    var menu = document.querySelector('.menu[data-menu="' + name + '"]');
    if (!menu) return null;
    var pop = menu.querySelector(".menu-pop");
    var trigger = menu.querySelector(".menu-trigger");
    if (!pop) return null;
    pop.hidden = false;
    if (trigger) trigger.setAttribute("aria-expanded", "true");
    var first = pop.querySelector('button:not([hidden]):not([disabled])');
    if (first) first.focus();
    return pop;
  }

  function isMenuOpen() {
    return !!document.querySelector(".menu-pop:not([hidden])");
  }

  function closeSheet() {
    var sheet = $("overflow-sheet");
    if (!sheet) return;
    sheet.hidden = true;
    var btn = $("btn-overflow");
    if (btn) {
      btn.setAttribute("aria-expanded", "false");
      btn.focus();
    }
  }

  function openSheet() {
    closeMenus();
    var sheet = $("overflow-sheet");
    if (!sheet) return;
    sheet.hidden = false;
    var btn = $("btn-overflow");
    if (btn) btn.setAttribute("aria-expanded", "true");
    var panel = sheet.querySelector(".sheet-panel");
    if (panel) panel.focus();
  }

  function isSheetOpen() {
    var sheet = $("overflow-sheet");
    return !!(sheet && !sheet.hidden);
  }

  function indentSelected() {
    if (!state.selectedId) {
      toast("Select a task first");
      return;
    }
    var idx = indexById(state.selectedId);
    if (idx >= 0) indentBlock(idx);
  }

  function outdentSelected() {
    if (!state.selectedId) {
      toast("Select a task first");
      return;
    }
    var idx = indexById(state.selectedId);
    if (idx >= 0) outdentBlock(idx);
  }

  function toggleCritical(force) {
    state.showCritical = force == null ? !state.showCritical : !!force;
    renderZoom();
    renderTimeline();
    updateStatus();
    persist();
  }

  function openHolidays() {
    renderHolidayList();
    $("holiday-dialog").showModal();
  }

  function isStandalone() {
    return (
      window.matchMedia("(display-mode: standalone)").matches ||
      window.navigator.standalone === true
    );
  }

  function isIos() {
    return /iPad|iPhone|iPod/.test(navigator.userAgent) ||
      (navigator.platform === "MacIntel" && navigator.maxTouchPoints > 1);
  }

  function updateInstallButtons() {
    var show = !isStandalone() && (!!pendingInstall || isIos());
    var btn = $("btn-install");
    if (btn) btn.hidden = !show;
    document.querySelectorAll("[data-install-item]").forEach(function (el) {
      el.hidden = !show;
    });
  }

  function installApp() {
    if (pendingInstall) {
      var ev = pendingInstall;
      pendingInstall = null;
      ev.prompt();
      ev.userChoice.then(function () {
        updateInstallButtons();
      });
      return;
    }
    if (isIos() && !isStandalone()) {
      $("install-dialog").showModal();
      return;
    }
    toast("This browser does not offer an install prompt. Use your browser's Install or Add to Home Screen option.");
  }

  function runAction(name) {
    if (name === "new") newProject();
    else if (name === "sample") loadSample();
    else if (name === "import") $("import-file").click();
    else if (name === "export-json") exportJSON();
    else if (name === "export-xlsx") exportExcel();
    else if (name === "export-png") exportPNG();
    else if (name === "install") installApp();
    else if (name === "undo") undo();
    else if (name === "redo") redo();
    else if (name === "add") addTask();
    else if (name === "indent") indentSelected();
    else if (name === "outdent") outdentSelected();
    else if (name === "theme") applyTheme(currentTheme() === "dark" ? "light" : "dark");
    else if (name === "critical") toggleCritical();
    else if (name === "holidays") openHolidays();
    else if (name === "baseline") setBaseline();
    else if (name === "clear-baseline") clearBaseline();
  }

  function initMenus() {
    document.querySelectorAll(".menu-trigger").forEach(function (trigger) {
      trigger.addEventListener("click", function (e) {
        e.preventDefault();
        e.stopPropagation();
        var menu = trigger.closest(".menu");
        var name = menu && menu.getAttribute("data-menu");
        var pop = menu && menu.querySelector(".menu-pop");
        if (pop && !pop.hidden) closeMenus();
        else if (name) openMenu(name);
      });
      trigger.addEventListener("keydown", function (e) {
        if (e.key === "ArrowDown" || e.key === "Enter" || e.key === " ") {
          e.preventDefault();
          var menu = trigger.closest(".menu");
          var name = menu && menu.getAttribute("data-menu");
          if (name) openMenu(name);
        }
      });
    });

    document.querySelectorAll(".menu-pop").forEach(function (pop) {
      pop.addEventListener("keydown", function (e) {
        var items = Array.prototype.slice.call(pop.querySelectorAll("button:not([hidden]):not([disabled])"));
        var idx = items.indexOf(document.activeElement);
        if (e.key === "Escape") {
          e.preventDefault();
          closeMenus();
          var trigger = pop.parentElement.querySelector(".menu-trigger");
          if (trigger) trigger.focus();
        } else if (e.key === "ArrowDown") {
          e.preventDefault();
          items[(idx + 1) % items.length].focus();
        } else if (e.key === "ArrowUp") {
          e.preventDefault();
          items[(idx - 1 + items.length) % items.length].focus();
        } else if (e.key === "Home") {
          e.preventDefault();
          if (items[0]) items[0].focus();
        } else if (e.key === "End") {
          e.preventDefault();
          if (items.length) items[items.length - 1].focus();
        } else if (e.key === "Tab") {
          closeMenus();
        }
      });
      pop.addEventListener("click", function () {
        closeMenus();
      });
    });

    document.addEventListener("pointerdown", function (e) {
      if (!e.target.closest(".menu")) closeMenus();
      if (isSheetOpen() && e.target.closest("[data-close-sheet]")) closeSheet();
    });
  }

  function initSheet() {
    $("btn-overflow").addEventListener("click", function () {
      if (isSheetOpen()) closeSheet();
      else openSheet();
    });
    $("btn-sheet-close").addEventListener("click", closeSheet);
    $("overflow-sheet").addEventListener("click", function (e) {
      var item = e.target.closest(".sheet-item, [data-action], [data-zoom], [data-cal]");
      if (!item || item.closest(".menu-pop")) return;
      // Let bind() handlers run; close after a tick so the click is processed.
      setTimeout(closeSheet, 0);
    });
  }

  function registerServiceWorker() {
    if (!("serviceWorker" in navigator)) return;
    var swUrl = new URL("sw.js", window.location.href);
    navigator.serviceWorker
      .register(swUrl.href)
      .then(function (reg) {
        function onInstalled() {
          if (navigator.serviceWorker.controller) promptSwUpdate(reg);
          else if (reg.waiting) reg.waiting.postMessage({ type: "SKIP_WAITING" });
        }
        if (reg.waiting) onInstalled();
        reg.addEventListener("updatefound", function () {
          var worker = reg.installing;
          if (!worker) return;
          worker.addEventListener("statechange", function () {
            if (worker.state === "installed") onInstalled();
          });
        });
      })
      .catch(function () {});
    navigator.serviceWorker.addEventListener("controllerchange", function () {
      if (!swRefreshing) return;
      window.location.reload();
    });
  }

  function promptSwUpdate(reg) {
    if (!reg.waiting) return;
    toast("Update available.", {
      actionLabel: "Reload",
      onAction: function () {
        swRefreshing = true;
        if (reg.waiting) reg.waiting.postMessage({ type: "SKIP_WAITING" });
      },
    });
  }

  /* ---- boot ------------------------------------------------------------- */

  function bind() {
    $("btn-add").addEventListener("click", function () { addTask(); });
    $("btn-add-header").addEventListener("click", function () { addTask(); });
    $("btn-new").addEventListener("click", newProject);
    $("btn-sample").addEventListener("click", loadSample);
    $("btn-export-json").addEventListener("click", exportJSON);
    $("btn-export-xlsx").addEventListener("click", exportExcel);
    $("btn-export-png").addEventListener("click", exportPNG);
    $("btn-import").addEventListener("click", function () { $("import-file").click(); });
    $("import-file").addEventListener("change", function (e) {
      var file = e.target.files && e.target.files[0];
      e.target.value = "";
      if (file) importFile(file);
    });
    $("btn-undo").addEventListener("click", undo);
    $("btn-redo").addEventListener("click", redo);
    $("btn-theme").addEventListener("click", function () {
      applyTheme(currentTheme() === "dark" ? "light" : "dark");
    });
    $("btn-baseline").addEventListener("click", setBaseline);
    $("btn-clear-baseline").addEventListener("click", clearBaseline);
    $("chk-critical").addEventListener("click", function (e) {
      e.preventDefault();
      toggleCritical();
    });
    $("btn-holidays").addEventListener("click", openHolidays);
    $("btn-install").addEventListener("click", installApp);
    $("holiday-add").addEventListener("click", function () {
      addHoliday($("holiday-date").value);
      $("holiday-date").value = "";
    });
    $("holiday-list").addEventListener("click", function (e) {
      var btn = e.target.closest("[data-remove-holiday]");
      if (btn) removeHoliday(btn.getAttribute("data-remove-holiday"));
    });
    $("project-title").addEventListener("focus", maybeCheckpoint);
    $("project-title").addEventListener("input", function (e) {
      state.projectName = e.target.value;
      persist();
    });
    $("project-title").addEventListener("blur", function () {
      typingCheckpoint = false;
    });

    document.addEventListener("click", function (e) {
      var action = e.target.closest("[data-action]");
      if (!action) return;
      if (action.id && (action.id === "btn-new" || action.id === "btn-sample" || action.id === "btn-import" || action.id.indexOf("btn-export") === 0 || action.id === "btn-install" || action.id === "btn-holidays" || action.id === "btn-baseline" || action.id === "btn-clear-baseline" || action.id === "chk-critical" || action.id === "btn-theme")) {
        return;
      }
      runAction(action.getAttribute("data-action"));
    });

    document.querySelectorAll("[data-zoom]").forEach(function (btn) {
      btn.addEventListener("click", function () {
        state.zoom = btn.getAttribute("data-zoom");
        closeMenus();
        renderZoom();
        renderTimeline();
        persist();
      });
    });
    document.querySelectorAll("[data-cal]").forEach(function (btn) {
      btn.addEventListener("click", function () {
        var mode = btn.getAttribute("data-cal");
        if (mode === state.calendarMode) return;
        pushHistory();
        state.calendarMode = mode;
        closeMenus();
        runSchedule();
        renderAll();
        persist();
      });
    });
    document.querySelectorAll("[data-view]").forEach(function (btn) {
      btn.addEventListener("click", function () {
        var view = btn.getAttribute("data-view");
        state.view = view === "resources" ? "resources" : "timeline";
        state.mobilePane = state.view;
        closeMenus();
        renderZoom();
        renderTimeline();
        persist();
      });
    });
    document.querySelectorAll("[data-pane]").forEach(function (btn) {
      btn.addEventListener("click", function () {
        setMobilePane(btn.getAttribute("data-pane"));
      });
    });

    var body = $("task-body");
    body.addEventListener("input", onTableInput);
    body.addEventListener("change", onTableInput);
    body.addEventListener("focusout", onTableFocusOut);
    body.addEventListener("click", onTableClick);
    body.addEventListener("pointerover", onTablePointerOver);
    body.addEventListener("dragstart", onGrabDragStart);
    body.addEventListener("dragend", onGrabDragEnd);
    body.addEventListener("dragover", onTableDragOver);
    body.addEventListener("drop", onTableDrop);

    $("timeline-canvas").addEventListener("pointerover", function (e) {
      var row = e.target.closest(".tl-row");
      var id = row && row.dataset.id;
      if (hoverId === id) return;
      hoverId = id || null;
      syncRowHover();
    });
    $("timeline-canvas").addEventListener("pointerdown", onCanvasPointerDown);
    $("timeline-body-scroll").addEventListener("scroll", onTimelineScroll);
    $("table-scroll").addEventListener("scroll", onTableScroll);

    document.addEventListener("keydown", function (e) {
      var meta = e.metaKey || e.ctrlKey;
      if (e.key === "Escape") {
        if (isSheetOpen()) {
          e.preventDefault();
          closeSheet();
          return;
        }
        if (isMenuOpen()) {
          e.preventDefault();
          closeMenus();
          return;
        }
      }
      if (meta && e.key.toLowerCase() === "z") {
        e.preventDefault();
        if (e.shiftKey) redo();
        else undo();
      } else if (meta && e.key.toLowerCase() === "y") {
        e.preventDefault();
        redo();
      }
    });

    window.addEventListener("resize", function () {
      applySplit();
      applyMobilePane();
      renderTimeline();
    });

    window.addEventListener("beforeinstallprompt", function (e) {
      e.preventDefault();
      pendingInstall = e;
      updateInstallButtons();
    });
    window.addEventListener("appinstalled", function () {
      pendingInstall = null;
      updateInstallButtons();
      toast("Gantt was installed");
    });

    initMenus();
    initSheet();
    initSplitter();
  }

  function boot() {
    initTheme();
    bind();
    registerServiceWorker();
    updateInstallButtons();
    if (!loadPersisted()) {
      renderAll();
    } else {
      applySplit();
      applyMobilePane();
    }
  }

  window.GanttApp = {
    getState: function () { return clone(state); },
    reschedule: runSchedule,
    addTask: function (partial) { return addTask(partial || {}, { silent: false }); },
    addTaskSilent: function (partial) {
      var t = addTask(partial || {}, { silent: true });
      renderAll();
      persist();
      return t;
    },
    setField: function (id, field, value) { onFieldInput(id, field, value); },
    loadSample: function () { applyProject(sampleProject()); },
    newProject: function () {
      state.projectName = "Untitled project";
      state.tasks = [];
      state.nextNumber = 1;
      state.selectedId = null;
      state.calendarMode = "calendar";
      state.holidays = [];
      state.showCritical = false;
      state.view = "timeline";
      state.mobilePane = "timeline";
      state.baselineSetAt = null;
      renderAll();
      persist();
    },
    render: renderAll,
    computeRange: computeRange,
    setCalendarMode: function (mode) {
      state.calendarMode = mode === "workdays" ? "workdays" : "calendar";
      runSchedule();
      renderAll();
      persist();
    },
    setHolidays: function (list) {
      state.holidays = S.uniqueSortedDates(list || []);
      runSchedule();
      renderAll();
      persist();
    },
    setBaseline: setBaseline,
    clearBaseline: clearBaseline,
    setShowCritical: function (on) {
      state.showCritical = !!on;
      renderAll();
      persist();
    },
    setView: function (view) {
      state.view = view === "resources" ? "resources" : "timeline";
      state.mobilePane = state.view;
      renderAll();
      persist();
    },
    setMobilePane: setMobilePane,
    openMenu: openMenu,
    closeMenus: closeMenus,
    openSheet: openSheet,
    closeSheet: closeSheet,
    isMenuOpen: isMenuOpen,
    isSheetOpen: isSheetOpen,
    isMobileLayout: isMobileLayout,

    exportExcelBuffer: exportExcelBuffer,
    importProject: function (data) { applyProject(data); },
    sampleProject: sampleProject,
  };

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", boot);
  else boot();
})();
