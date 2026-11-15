/* Gantt chart SPA — table, timeline, persistence, import/export */
(function () {
  "use strict";

  var S = window.GanttSchedule;
  var STORAGE_KEY = "gantt-chart-app:v1";
  var THEME_KEY = "gantt-chart-app:theme";
  var ROW_H = 40;
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
    projectName: "Untitled project",
    tasks: [],
    nextNumber: 1,
    zoom: "week",
    split: 560,
    selectedId: null,
  };

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

  function snapshot() {
    return {
      projectName: state.projectName,
      tasks: clone(state.tasks),
      nextNumber: state.nextNumber,
      zoom: state.zoom,
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
    S.reschedule(state.tasks);
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
    $("btn-undo").disabled = !historyStack.length;
    $("btn-redo").disabled = !redoStack.length;
  }

  function toast(msg) {
    var el = $("toast");
    el.textContent = msg;
    el.classList.add("show");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () {
      el.classList.remove("show");
    }, 2400);
  }

  function persist() {
    if (suppressSave) return;
    try {
      localStorage.setItem(
        STORAGE_KEY,
        JSON.stringify({
          projectName: state.projectName,
          tasks: state.tasks,
          nextNumber: state.nextNumber,
          zoom: state.zoom,
          split: state.split,
        })
      );
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
    S.reschedule(state.tasks);
    if (opts.resetHistory) {
      historyStack = [];
      redoStack = [];
    }
    renderAll();
    persist();
  }

  function normalizeIncoming(raw, num) {
    return {
      id: raw.id || uid(),
      number: num,
      name: raw.name != null ? String(raw.name) : "",
      start: S.isISO(raw.start) ? raw.start : null,
      end: S.isISO(raw.end) ? raw.end : null,
      duration: raw.duration == null || raw.duration === "" ? null : Number(raw.duration),
      progress: S.clampProgress(raw.progress),
      dependencies: S.parseDependencyInput(
        Array.isArray(raw.dependencies) ? raw.dependencies.join(",") : raw.dependencies || ""
      ),
      assignee: raw.assignee != null ? String(raw.assignee) : "",
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
    var task = {
      id: uid(),
      number: state.nextNumber++,
      name: partial.name || "",
      start: partial.start || S.todayISO(),
      duration: partial.duration != null ? partial.duration : 1,
      end: partial.end || null,
      progress: partial.progress || 0,
      dependencies: partial.dependencies || [],
      assignee: partial.assignee || "",
      color: partial.color || nextColor(),
      indent: partial.indent || 0,
      collapsed: false,
    };
    return task;
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
      tasks: [
        { number: 1, name: "Discovery", indent: 0, color: "#0f766e", assignee: "" },
        { number: 2, name: "Kickoff meeting", start: "2026-11-02", duration: 1, progress: 100, indent: 1, color: "#0f766e", assignee: "Alex" },
        { number: 3, name: "Stakeholder interviews", start: "2026-11-03", duration: 4, progress: 100, indent: 1, color: "#0f766e", assignee: "Alex", dependencies: [2] },
        { number: 4, name: "Requirements document", start: "2026-11-09", duration: 5, progress: 100, indent: 1, color: "#0f766e", assignee: "Jordan", dependencies: [3] },
        { number: 5, name: "Design", indent: 0, color: "#7c3aed", dependencies: [4] },
        { number: 6, name: "Wireframes", duration: 4, progress: 40, indent: 1, color: "#7c3aed", assignee: "Sam", dependencies: [4] },
        { number: 7, name: "Visual design", duration: 6, progress: 0, indent: 1, color: "#7c3aed", assignee: "Sam", dependencies: [6] },
        { number: 8, name: "Design review", duration: 2, progress: 0, indent: 1, color: "#7c3aed", assignee: "Alex", dependencies: [7] },
        { number: 9, name: "Build", indent: 0, color: "#0369a1", dependencies: [8] },
        { number: 10, name: "Frontend implementation", duration: 9, indent: 1, color: "#0369a1", assignee: "Riley", dependencies: [8] },
        { number: 11, name: "Backend APIs", duration: 8, indent: 1, color: "#0369a1", assignee: "Quinn", dependencies: [8] },
        { number: 12, name: "Integrations", duration: 4, indent: 1, color: "#0369a1", assignee: "Quinn", dependencies: [11] },
        { number: 13, name: "Launch", indent: 0, color: "#c2410c", dependencies: [9] },
        { number: 14, name: "QA & bug bash", duration: 3, indent: 1, color: "#c2410c", assignee: "Jordan", dependencies: [9] },
        { number: 15, name: "Staging deploy", duration: 1, indent: 1, color: "#c2410c", assignee: "Riley", dependencies: [14] },
        { number: 16, name: "Go live", duration: 1, indent: 1, color: "#be185d", assignee: "Alex", dependencies: [15] },
      ],
    };
  }

  /* ---- schedule + field edits ------------------------------------------ */

  function runSchedule() {
    var result = S.reschedule(state.tasks);
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
    else if (field === "progress") {
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
      task.duration = Math.max(1, Math.round(n));
    } else if (field === "end") {
      if (value === "") {
        return false;
      } else if (S.isISO(value)) {
        if (S.isISO(task.start) && value < task.start) {
          task.duration = S.inclusiveDuration(task.start, value) ? 1 : 1;
          task.end = task.start;
          task.duration = 1;
        } else if (S.isISO(task.start)) {
          task.end = value;
          task.duration = S.inclusiveDuration(task.start, task.end);
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
    if (task.isGroup && (field === "start" || field === "end" || field === "duration" || field === "progress")) {
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
    var desc = S.getDescendantIndexes(state.tasks, idx);
    state.tasks.splice(idx, 1);
    desc.forEach(function () {
      /* descendants stay; their indent is reduced if they would float */
    });
    /* After removing the parent only, former children keep indent. Reduce orphans that would have a gap. */
    normalizeIndents();
    state.tasks.forEach(function (t) {
      t.dependencies = (t.dependencies || []).filter(function (d) {
        return d !== num;
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
    renderAll();
    persist();
  }

  function loadSample() {
    if (state.tasks.length && !confirm("Replace the current project with the sample?")) return;
    pushHistory();
    applyProject(sampleProject());
    toast("Loaded sample project");
  }

  /* ---- render table ----------------------------------------------------- */

  function icon(path) {
    return (
      '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6">' +
      path +
      "</svg>"
    );
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
        '<td class="col-dur"><input class="cell-input" type="number" min="1" step="1" data-field="duration" value="' +
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
        '" placeholder="e.g. 2, 4" title="Predecessor task IDs, comma-separated" /></td>' +
        '<td class="col-who"><input class="cell-input" data-field="assignee" value="' +
        escapeAttr(task.assignee) +
        '" placeholder="Name" /></td>' +
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
      ["start", "duration", "end", "progress"].forEach(function (field) {
        var input = tr.querySelector('[data-field="' + field + '"]');
        if (!input || input === active) return;
        var val = task[field] == null ? "" : String(task[field]);
        if (input.value !== val) input.value = val;
        var ro = !!task.isGroup;
        input.readOnly = ro;
      });
      var deps = tr.querySelector('[data-field="dependencies"]');
      if (deps && deps !== active) {
        var formatted = S.formatDependencies(task.dependencies);
        if (deps.value !== formatted) deps.value = formatted;
      }
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
      if (S.isISO(t.start) && t.start < start) start = t.start;
      if (S.isISO(t.end) && t.end > end) end = t.end;
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

  function renderTimeline() {
    var range = computeRange();
    lastRange = range;
    var px = pxPerDay();
    var width = range.days * px;
    var entries = visibleEntries();
    var height = Math.max(entries.length * ROW_H, $("timeline-body-scroll").clientHeight || 200);

    var header = $("timeline-header");
    header.style.width = width + "px";
    var hhtml = "";
    var monthStart = range.start;
    var curMonth = monthStart.slice(0, 7);
    S.eachDay(range.start, range.end, function (iso, d) {
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
      var showDay = state.zoom === "day" || state.zoom === "week" || d.getDate() === 1 || d.getDate() % (state.zoom === "month" ? 5 : 1) === 0;
      if (state.zoom === "month" && !showDay && !today && !weekend) {
        /* still draw a quiet tick */
      }
      var dow = ["S", "M", "T", "W", "T", "F", "S"][d.getDay()];
      var label =
        state.zoom === "month"
          ? today || d.getDate() === 1 || d.getDate() % 5 === 0
            ? '<span class="dom">' + d.getDate() + "</span>"
            : ""
          : '<span class="dom">' + d.getDate() + "</span>" + (state.zoom === "day" ? '<span class="dow">' + dow + "</span>" : "");
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

    var canvas = $("timeline-canvas");
    var chtml = "";
    /* weekend + grid */
    chtml += '<div class="tl-grid" style="position:absolute;inset:0;height:' + height + "px;width:" + width + 'px">';
    S.eachDay(range.start, range.end, function (iso, d) {
      var left = dayLeft(range, iso);
      if (d.getDay() === 0 || d.getDay() === 6) {
        chtml += '<div class="tl-weekend" style="left:' + left + "px;width:" + px + 'px"></div>';
      }
      if (state.zoom === "day" || d.getDay() === 1 || d.getDate() === 1) {
        chtml += '<div class="tl-gridline" style="left:' + left + 'px"></div>';
      }
    });
    if (today >= range.start && today <= range.end) {
      chtml += '<div class="today-line" style="left:' + (dayLeft(range, today) + px / 2) + 'px"></div>';
    }
    chtml += "</div>";

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
      if (S.isISO(task.start) && S.isISO(task.end) && task.end >= task.start) {
        var left = dayLeft(range, task.start) + 1;
        var w = Math.max(px - 2, task.duration * px - 2);
        var lum = hexLuminance(task.color || PALETTE[0]);
        var labelColor = lum > 0.62 ? "#1c212b" : "#fff";
        chtml +=
          '<div class="bar' +
          (task.isGroup ? " is-group" : "") +
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
          (task.isGroup ? "" : '<div class="bar-handle start" data-handle="start"></div><div class="bar-handle end" data-handle="end"></div>') +
          "</div>";
      }
      chtml += "</div>";
    });

    /* arrows */
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
    entries.forEach(function (entry) {
      var task = entry.task;
      (task.dependencies || []).forEach(function (depNum) {
        var pred = S.findTaskByNumber(state.tasks, depNum);
        if (!pred || !S.isISO(pred.end) || !S.isISO(task.start)) return;
        if (visIndex[pred.number] == null || visIndex[task.number] == null) return;
        var y1 = visIndex[pred.number] * ROW_H + ROW_H / 2;
        var y2 = visIndex[task.number] * ROW_H + ROW_H / 2;
        var x1 = dayLeft(range, pred.end) + px - 1;
        var x2 = dayLeft(range, task.start) + 1;
        chtml += arrowPath(x1, y1, x2, y2);
      });
    });
    chtml += "</svg>";

    canvas.style.width = width + "px";
    canvas.style.height = height + "px";
    canvas.innerHTML = chtml;

    $("timeline-header-clip").scrollLeft = $("timeline-body-scroll").scrollLeft;
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

  function arrowPath(x1, y1, x2, y2) {
    var dx = Math.max(16, Math.abs(x2 - x1) / 2);
    var c1x = x1 + dx;
    var c2x = x2 - dx;
    var tip = 5;
    var path = "M" + x1 + " " + y1 + " C " + c1x + " " + y1 + ", " + c2x + " " + y2 + ", " + x2 + " " + y2;
    /* arrowhead */
    return (
      '<path d="' +
      path +
      '" />' +
      '<polygon points="' +
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
    if (!starts.length) {
      el.textContent = "";
      return;
    }
    var a = S.minISO.apply(null, starts);
    var b = S.maxISO.apply(null, ends);
    el.textContent = formatLong(a) + " – " + formatLong(b);
  }

  function formatLong(iso) {
    var d = S.parseISO(iso);
    return d ? d.toLocaleDateString(undefined, { month: "short", day: "numeric", year: "numeric" }) : iso;
  }

  function renderZoom() {
    document.querySelectorAll(".seg [data-zoom]").forEach(function (btn) {
      btn.setAttribute("aria-pressed", btn.getAttribute("data-zoom") === state.zoom ? "true" : "false");
    });
  }

  function renderAll() {
    $("project-title").value = state.projectName;
    applySplit();
    renderZoom();
    renderTable();
    renderTimeline();
    updateStatus();
    updateUndoButtons();
  }

  /* ---- interactions ----------------------------------------------------- */

  function onTableInput(e) {
    var input = e.target.closest("[data-field]");
    if (!input) return;
    var tr = input.closest("tr");
    if (!tr) return;
    maybeCheckpoint();
    onFieldInput(tr.getAttribute("data-id"), input.getAttribute("data-field"), input.value);
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

  /* row drag */
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
    var after = e.clientY > rect.top + rect.height / 2;
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

  /* bar drag */
  function onCanvasPointerDown(e) {
    var handle = e.target.closest(".bar-handle");
    var bar = e.target.closest(".bar");
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
    if (task.isGroup && handle) return;
    e.preventDefault();
    var mode = handle ? (handle.getAttribute("data-handle") === "start" ? "resize-start" : "resize-end") : "move";
    pushHistory();
    drag = {
      id: task.id,
      mode: mode,
      startX: e.clientX,
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
    if (drag.mode === "move") {
      S.shiftBlock(state.tasks, idx, delta);
    } else if (drag.mode === "resize-end") {
      var newDur = Math.max(1, drag.origDuration + delta);
      task.duration = newDur;
      task.end = S.endFromStartDuration(task.start, newDur);
    } else if (drag.mode === "resize-start") {
      var newStart = S.addDays(drag.origStart, delta);
      var nd = S.inclusiveDuration(newStart, drag.origEnd);
      if (!nd || nd < 1) {
        task.start = drag.origEnd;
        task.duration = 1;
        task.end = drag.origEnd;
      } else {
        task.start = newStart;
        task.duration = nd;
        task.end = drag.origEnd;
      }
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
    var changed = JSON.stringify(drag.snapshot) !== JSON.stringify(state.tasks.map(function (t) {
      return { id: t.id, start: t.start, end: t.end, duration: t.duration };
    }));
    drag = null;
    if (!changed) {
      historyStack.pop();
      updateUndoButtons();
    }
    persist();
    renderTimeline();
  }

  /* scroll sync */
  function onTimelineScroll() {
    $("timeline-header-clip").scrollLeft = $("timeline-body-scroll").scrollLeft;
    if (!syncingScroll) {
      syncingScroll = true;
      $("table-scroll").scrollTop = $("timeline-body-scroll").scrollTop;
      syncingScroll = false;
    }
  }
  var syncingScroll = false;
  function onTableScroll() {
    if (syncingScroll) return;
    syncingScroll = true;
    $("timeline-body-scroll").scrollTop = $("table-scroll").scrollTop;
    syncingScroll = false;
  }

  /* splitter */
  function applySplit() {
    if (window.innerWidth <= 860) {
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
    var data = {
      version: 1,
      projectName: state.projectName,
      zoom: state.zoom,
      nextNumber: state.nextNumber,
      exportedAt: new Date().toISOString(),
      tasks: state.tasks.map(function (t) {
        return {
          number: t.number,
          name: t.name,
          start: t.start,
          end: t.end,
          duration: t.duration,
          progress: t.progress,
          dependencies: t.dependencies,
          assignee: t.assignee,
          color: t.color,
          indent: t.indent,
          collapsed: !!t.collapsed,
        };
      }),
    };
    downloadFile(slug(state.projectName) + ".json", "application/json", JSON.stringify(data, null, 2));
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

  function importJSONFile(file) {
    var reader = new FileReader();
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
      deps: 64,
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
    var heads = [
      ["id", "ID"],
      ["name", "TASK"],
      ["start", "START"],
      ["dur", "DAYS"],
      ["end", "END"],
      ["prog", "%"],
      ["deps", "DEPS"],
      ["who", "ASSIGNEE"],
    ];
    heads.forEach(function (h) {
      ctx.fillText(h[1], colX(h[0]) + 8, oy + 18);
    });

    /* timeline month bands + days */
    var tlX = ox + tableW + 8;
    var monthStart = range.start;
    var curMonth = monthStart.slice(0, 7);
    ctx.font = "600 10px Instrument Sans, sans-serif";
    ctx.fillStyle = soft;
    function flushBand(from, to) {
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
        flushBand(monthStart, S.addDays(iso, -1));
        curMonth = key;
        monthStart = iso;
      }
    });
    flushBand(monthStart, range.end);

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
      ctx.fillStyle = text;
      ctx.font = (task.isGroup ? "650 " : "") + "11px Instrument Sans, sans-serif";
      ctx.fillStyle = faint;
      ctx.font = "11px IBM Plex Mono, monospace";
      ctx.fillText("#" + task.number, colX("id") + 8, y + 18);
      ctx.fillStyle = text;
      ctx.font = (task.isGroup ? "650 " : "500 ") + "12px Instrument Sans, sans-serif";
      var name = (task.indent ? "  ".repeat(task.indent) : "") + (task.name || "Untitled");
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

      if (S.isISO(task.start) && S.isISO(task.end)) {
        var x = tlX + S.diffDays(range.start, task.start) * px + 1;
        var w = Math.max(px - 2, (task.duration || 1) * px - 2);
        var bh = task.isGroup ? 10 : 14;
        var by = y + (rowH - bh) / 2;
        roundRect(ctx, x, by, w, bh, 4);
        ctx.fillStyle = task.color || PALETTE[0];
        ctx.fill();
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
    });

    /* arrows on export */
    ctx.strokeStyle = faint;
    ctx.fillStyle = faint;
    ctx.lineWidth = 1.2;
    rows.forEach(function (task, i) {
      (task.dependencies || []).forEach(function (depNum) {
        var pred = S.findTaskByNumber(rows, depNum);
        if (!pred || !S.isISO(pred.end) || !S.isISO(task.start)) return;
        var pi = rows.indexOf(pred);
        var y1 = oy + tlHead + pi * rowH + rowH / 2;
        var y2 = oy + tlHead + i * rowH + rowH / 2;
        var x1 = tlX + S.diffDays(range.start, pred.end) * px + px - 1;
        var x2 = tlX + S.diffDays(range.start, task.start) * px + 1;
        var dx = Math.max(12, Math.abs(x2 - x1) / 2);
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

  /* ---- theme ------------------------------------------------------------ */

  function currentTheme() {
    return document.documentElement.getAttribute("data-theme") || "light";
  }

  function applyTheme(theme) {
    document.documentElement.setAttribute("data-theme", theme);
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

  /* ---- boot ------------------------------------------------------------- */

  function bind() {
    $("btn-add").addEventListener("click", function () {
      addTask();
    });
    $("btn-new").addEventListener("click", newProject);
    $("btn-sample").addEventListener("click", loadSample);
    $("btn-export-json").addEventListener("click", exportJSON);
    $("btn-export-png").addEventListener("click", exportPNG);
    $("btn-import").addEventListener("click", function () {
      $("import-file").click();
    });
    $("import-file").addEventListener("change", function (e) {
      var file = e.target.files && e.target.files[0];
      e.target.value = "";
      if (file) importJSONFile(file);
    });
    $("btn-undo").addEventListener("click", undo);
    $("btn-redo").addEventListener("click", redo);
    $("btn-theme").addEventListener("click", function () {
      applyTheme(currentTheme() === "dark" ? "light" : "dark");
    });
    $("project-title").addEventListener("focus", maybeCheckpoint);
    $("project-title").addEventListener("input", function (e) {
      state.projectName = e.target.value;
      persist();
    });
    $("project-title").addEventListener("blur", function () {
      typingCheckpoint = false;
    });

    document.querySelectorAll(".seg [data-zoom]").forEach(function (btn) {
      btn.addEventListener("click", function () {
        state.zoom = btn.getAttribute("data-zoom");
        renderZoom();
        renderTimeline();
        persist();
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
      renderTimeline();
    });

    initSplitter();
  }

  function boot() {
    initTheme();
    bind();
    if (!loadPersisted()) {
      renderAll();
    } else {
      applySplit();
    }
  }

  window.GanttApp = {
    getState: function () {
      return clone(state);
    },
    reschedule: runSchedule,
    addTask: function (partial) {
      return addTask(partial || {}, { silent: false });
    },
    addTaskSilent: function (partial) {
      var t = addTask(partial || {}, { silent: true });
      renderAll();
      persist();
      return t;
    },
    setField: function (id, field, value) {
      onFieldInput(id, field, value);
    },
    loadSample: function () {
      applyProject(sampleProject());
    },
    newProject: function () {
      state.projectName = "Untitled project";
      state.tasks = [];
      state.nextNumber = 1;
      state.selectedId = null;
      renderAll();
      persist();
    },
    render: renderAll,
    computeRange: computeRange,
  };

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", boot);
  else boot();
})();
