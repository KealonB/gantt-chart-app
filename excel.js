/**
 * Excel / CSV import and export for the Gantt app.
 * Pass a SheetJS-compatible library (xlsx-js-style) as `XLSX`.
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory();
  } else {
    root.GanttExcel = factory();
  }
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  var TASK_HEADERS = [
    "ID",
    "WBS",
    "Name",
    "Start",
    "End",
    "Duration",
    "Progress",
    "Dependencies",
    "Assignee",
    "Allocation %",
    "Milestone",
    "Baseline Start",
    "Baseline End",
    "Variance",
    "Critical",
    "Color",
    "Indent",
  ];

  var HEADER_ALIASES = {
    id: "ID",
    number: "ID",
    taskid: "ID",
    "task id": "ID",
    wbs: "WBS",
    outline: "WBS",
    "outline number": "WBS",
    indent: "Indent",
    level: "Indent",
    name: "Name",
    task: "Name",
    title: "Name",
    "task name": "Name",
    start: "Start",
    "start date": "Start",
    begin: "Start",
    end: "End",
    finish: "End",
    "end date": "End",
    "finish date": "End",
    duration: "Duration",
    days: "Duration",
    dur: "Duration",
    progress: "Progress",
    "%": "Progress",
    "% complete": "Progress",
    complete: "Progress",
    dependencies: "Dependencies",
    depends: "Dependencies",
    deps: "Dependencies",
    predecessors: "Dependencies",
    predecessor: "Dependencies",
    assignee: "Assignee",
    resource: "Assignee",
    owner: "Assignee",
    who: "Assignee",
    allocation: "Allocation %",
    "allocation %": "Allocation %",
    alloc: "Allocation %",
    units: "Allocation %",
    milestone: "Milestone",
    mile: "Milestone",
    "baseline start": "Baseline Start",
    "baseline start date": "Baseline Start",
    "baseline end": "Baseline End",
    "baseline finish": "Baseline End",
    "baseline end date": "Baseline End",
    variance: "Variance",
    slip: "Variance",
    critical: "Critical",
    "critical path": "Critical",
    color: "Color",
    colour: "Color",
  };

  function headerFill(rgb) {
    return {
      fill: { patternType: "solid", fgColor: { rgb: rgb } },
      font: { bold: true, color: { rgb: "FFFFFF" }, name: "Calibri", sz: 11 },
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
      border: thinBorder("94A3B8"),
    };
  }

  function thinBorder(rgb) {
    var side = { style: "thin", color: { rgb: rgb || "E2E8F0" } };
    return { top: side, bottom: side, left: side, right: side };
  }

  function dateCell(iso, S) {
    if (!S.isISO(iso)) return { t: "s", v: "", s: { border: thinBorder(), alignment: { horizontal: "center" } } };
    var d = S.parseISO(iso);
    // Noon local avoids timezone rollover when Excel serializes the date.
    var dt = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 12, 0, 0);
    return {
      t: "d",
      v: dt,
      z: "yyyy-mm-dd",
      s: { border: thinBorder(), alignment: { horizontal: "center" }, numFmt: "yyyy-mm-dd" },
    };
  }

  function textCell(v, align) {
    return {
      t: "s",
      v: v == null ? "" : String(v),
      s: { border: thinBorder(), alignment: { horizontal: align || "left", vertical: "center" } },
    };
  }

  function numCell(v, align) {
    if (v == null || v === "") return textCell("", align || "center");
    return {
      t: "n",
      v: Number(v),
      s: { border: thinBorder(), alignment: { horizontal: align || "center", vertical: "center" } },
    };
  }

  function boolCell(v) {
    return {
      t: "s",
      v: v ? "Yes" : "No",
      s: { border: thinBorder(), alignment: { horizontal: "center" } },
    };
  }

  function colLetter(i) {
    var s = "";
    var n = i + 1;
    while (n > 0) {
      var rem = (n - 1) % 26;
      s = String.fromCharCode(65 + rem) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  function cellRef(r, c) {
    return colLetter(c) + String(r + 1);
  }

  function encodeRgb(hex) {
    var h = String(hex || "").replace("#", "").toUpperCase();
    if (h.length === 3) h = h[0] + h[0] + h[1] + h[1] + h[2] + h[2];
    if (h.length !== 6) return "0F766E";
    return h;
  }

  function luminance(hex) {
    var h = encodeRgb(hex);
    function lin(c) {
      c = parseInt(c, 16) / 255;
      return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
    }
    return 0.2126 * lin(h.slice(0, 2)) + 0.7152 * lin(h.slice(2, 4)) + 0.0722 * lin(h.slice(4, 6));
  }

  function buildWorkbook(project, XLSX, S) {
    S = S || (typeof GanttSchedule !== "undefined" ? GanttSchedule : require("./schedule.js"));
    var cal = S.makeCalendar(project.calendarMode || "calendar", project.holidays || []);
    var tasks = (project.tasks || []).slice();
    S.reschedule(tasks, { mode: cal.mode, holidays: cal.holidays, today: project.today });

    var wb = XLSX.utils.book_new();
    wb.Props = {
      Title: project.projectName || "Gantt project",
      CreatedDate: new Date(),
    };

    var taskSheet = buildTasksSheet(tasks, XLSX, S, cal);
    XLSX.utils.book_append_sheet(wb, taskSheet, "Tasks");

    var tlSheet = buildTimelineSheet(tasks, XLSX, S, cal, project);
    XLSX.utils.book_append_sheet(wb, tlSheet, "Timeline");

    var cfgSheet = buildConfigSheet(project, XLSX, S);
    XLSX.utils.book_append_sheet(wb, cfgSheet, "Config");

    return wb;
  }

  function buildTasksSheet(tasks, XLSX, S, cal) {
    var ws = {};
    TASK_HEADERS.forEach(function (h, c) {
      ws[cellRef(0, c)] = { t: "s", v: h, s: headerFill("0F766E") };
    });
    tasks.forEach(function (t, i) {
      var r = i + 1;
      var wbs = S.outlineNumber(tasks, i);
      var values = [
        numCell(t.number),
        textCell(wbs, "left"),
        textCell(t.name || ""),
        dateCell(t.start, S),
        dateCell(t.end, S),
        numCell(t.duration),
        numCell(t.progress),
        textCell(S.formatDependencies(t.dependencies)),
        textCell(t.assignee || ""),
        numCell(t.allocation == null ? 100 : t.allocation),
        boolCell(!!t.milestone),
        dateCell(t.baselineStart, S),
        dateCell(t.baselineEnd, S),
        t.variance == null ? textCell("") : numCell(t.variance),
        boolCell(!!t.critical),
        textCell(t.color || ""),
        numCell(t.indent || 0),
      ];
      values.forEach(function (cell, c) {
        if (t.critical && c === 14) {
          cell.s = cell.s || {};
          cell.s.fill = { patternType: "solid", fgColor: { rgb: "FEE2E2" } };
          cell.s.font = { bold: true, color: { rgb: "9F1239" } };
        }
        if (c === 15 && t.color) {
          var rgb = encodeRgb(t.color);
          cell.s.fill = { patternType: "solid", fgColor: { rgb: rgb } };
          cell.s.font = { color: { rgb: luminance(t.color) > 0.62 ? "1C212B" : "FFFFFF" } };
        }
        ws[cellRef(r, c)] = cell;
      });
    });
    ws["!ref"] = "A1:" + cellRef(Math.max(tasks.length, 1), TASK_HEADERS.length - 1);
    ws["!cols"] = [
      { wch: 6 },
      { wch: 10 },
      { wch: 32 },
      { wch: 14 },
      { wch: 14 },
      { wch: 10 },
      { wch: 10 },
      { wch: 18 },
      { wch: 16 },
      { wch: 12 },
      { wch: 11 },
      { wch: 16 },
      { wch: 16 },
      { wch: 10 },
      { wch: 10 },
      { wch: 12 },
      { wch: 8 },
    ];
    ws["!rows"] = [{ hpt: 22 }].concat(tasks.map(function () { return { hpt: 18 }; }));
    ws["!freeze"] = { xSplit: 3, ySplit: 1, topLeftCell: "D2", activePane: "bottomRight", state: "frozen" };
    return ws;
  }

  function buildTimelineSheet(tasks, XLSX, S, cal, project) {
    var start = S.todayISO();
    var end = start;
    tasks.forEach(function (t) {
      if (S.isISO(t.start) && t.start < start) start = t.start;
      if (S.isISO(t.end) && t.end > end) end = t.end;
      if (S.isISO(t.baselineStart) && t.baselineStart < start) start = t.baselineStart;
      if (S.isISO(t.baselineEnd) && t.baselineEnd > end) end = t.baselineEnd;
    });
    if (!tasks.length) {
      end = S.addDays(start, 13);
    }
    var span = S.diffDays(start, end) + 1;
    var useWeeks = span > 120;
    var columns = [];
    if (useWeeks) {
      var cur = start;
      // Align to Monday
      var d0 = S.parseISO(cur);
      var back = (d0.getDay() + 6) % 7;
      cur = S.addDays(cur, -back);
      while (cur <= end) {
        columns.push({ start: cur, end: S.addDays(cur, 6), label: cur });
        cur = S.addDays(cur, 7);
      }
    } else {
      S.eachDay(start, end, function (iso) {
        columns.push({ start: iso, end: iso, label: iso });
      });
    }

    var ws = {};
    ws[cellRef(0, 0)] = { t: "s", v: "ID", s: headerFill("0F766E") };
    ws[cellRef(0, 1)] = { t: "s", v: "Name", s: headerFill("0F766E") };
    columns.forEach(function (col, i) {
      var label = useWeeks ? col.start : col.start.slice(5);
      var weekend = !useWeeks && S.isWeekend(col.start);
      ws[cellRef(0, i + 2)] = {
        t: "s",
        v: label,
        s: {
          fill: { patternType: "solid", fgColor: { rgb: weekend ? "E2E8F0" : "0F766E" } },
          font: { bold: true, color: { rgb: weekend ? "475569" : "FFFFFF" }, name: "Calibri", sz: 8 },
          alignment: { horizontal: "center", vertical: "center", textRotation: useWeeks ? 0 : 90 },
          border: thinBorder("94A3B8"),
        },
      };
    });

    tasks.forEach(function (t, r) {
      var row = r + 1;
      ws[cellRef(row, 0)] = numCell(t.number);
      ws[cellRef(row, 1)] = textCell((t.indent ? "  ".repeat(t.indent) : "") + (t.name || ""));
      columns.forEach(function (col, i) {
        var covers = S.isISO(t.start) && S.isISO(t.end) && t.end >= col.start && t.start <= col.end;
        var baseline =
          S.isISO(t.baselineStart) &&
          S.isISO(t.baselineEnd) &&
          t.baselineEnd >= col.start &&
          t.baselineStart <= col.end;
        var weekend = !useWeeks && S.isWeekend(col.start);
        var cell = {
          t: "s",
          v: covers ? (t.milestone ? "◆" : t.critical ? "●" : "") : "",
          s: {
            alignment: { horizontal: "center", vertical: "center" },
            border: thinBorder(weekend ? "E2E8F0" : "F1F5F9"),
            fill: {
              patternType: "solid",
              fgColor: { rgb: weekend ? "F8FAFC" : "FFFFFF" },
            },
          },
        };
        if (covers) {
          var rgb = encodeRgb(t.color || "0F766E");
          cell.s.fill = { patternType: "solid", fgColor: { rgb: rgb } };
          cell.s.font = {
            color: { rgb: luminance(t.color || "0F766E") > 0.62 ? "1C212B" : "FFFFFF" },
            bold: true,
          };
        } else if (baseline) {
          cell.s.fill = { patternType: "solid", fgColor: { rgb: "CBD5E1" } };
        }
        ws[cellRef(row, i + 2)] = cell;
      });
    });

    var lastRow = Math.max(tasks.length, 1);
    var lastCol = columns.length + 1;
    ws["!ref"] = "A1:" + cellRef(lastRow, lastCol);
    ws["!cols"] = [{ wch: 6 }, { wch: 28 }].concat(
      columns.map(function () {
        return { wch: useWeeks ? 10 : 3.4 };
      })
    );
    ws["!rows"] = [{ hpt: useWeeks ? 22 : 36 }].concat(tasks.map(function () { return { hpt: 16 }; }));
    ws["!freeze"] = { xSplit: 2, ySplit: 1, topLeftCell: "C2", activePane: "bottomRight", state: "frozen" };
    return ws;
  }

  function buildConfigSheet(project, XLSX, S) {
    var rows = [
      ["Key", "Value"],
      ["Project name", project.projectName || ""],
      ["Calendar", project.calendarMode === "workdays" ? "workdays" : "calendar"],
      ["Holidays", (project.holidays || []).join(", ")],
      ["Zoom", project.zoom || "week"],
      ["Exported at", new Date().toISOString()],
    ];
    var ws = XLSX.utils.aoa_to_sheet(rows);
    ws["A1"].s = headerFill("0F766E");
    ws["B1"].s = headerFill("0F766E");
    ws["!cols"] = [{ wch: 18 }, { wch: 60 }];
    return ws;
  }

  function writeBuffer(project, XLSX, S) {
    var wb = buildWorkbook(project, XLSX, S);
    return XLSX.write(wb, { type: "array", bookType: "xlsx", cellStyles: true });
  }

  function writeNodeBuffer(project, XLSX, S) {
    var wb = buildWorkbook(project, XLSX, S);
    return XLSX.write(wb, { type: "buffer", bookType: "xlsx", cellStyles: true });
  }

  /* ---- import ----------------------------------------------------------- */

  function parseExcelDate(val, S) {
    if (val == null || val === "") return null;
    if (Object.prototype.toString.call(val) === "[object Date]" && !isNaN(val.getTime())) {
      return S.toISO(val);
    }
    if (typeof val === "string") {
      var s = val.trim();
      if (S.isISO(s)) return s;
      var m = /^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/.exec(s);
      if (m) {
        var iso = m[3] + "-" + String(m[1]).padStart(2, "0") + "-" + String(m[2]).padStart(2, "0");
        if (S.isISO(iso)) return iso;
        iso = m[3] + "-" + String(m[2]).padStart(2, "0") + "-" + String(m[1]).padStart(2, "0");
        if (S.isISO(iso)) return iso;
      }
      var d = new Date(s);
      if (!isNaN(d.getTime()) && /\d{4}/.test(s)) return S.toISO(d);
      return null;
    }
    if (typeof val === "number" && Number.isFinite(val)) {
      // Excel serial, 1899-12-30 epoch, treat as UTC date to avoid TZ shifts
      var epoch = Date.UTC(1899, 11, 30) + Math.round(val) * 86400000;
      var dt = new Date(epoch);
      return dt.getUTCFullYear() + "-" + String(dt.getUTCMonth() + 1).padStart(2, "0") + "-" + String(dt.getUTCDate()).padStart(2, "0");
    }
    return null;
  }

  function parseBool(val) {
    if (val == null || val === "") return false;
    if (typeof val === "boolean") return val;
    var s = String(val).trim().toLowerCase();
    return s === "y" || s === "yes" || s === "true" || s === "1" || s === "milestone";
  }

  function normalizeHeader(h) {
    return String(h || "")
      .trim()
      .toLowerCase()
      .replace(/[_]+/g, " ")
      .replace(/\s+/g, " ");
  }

  function mapHeaders(row) {
    var map = {};
    row.forEach(function (h, i) {
      var key = HEADER_ALIASES[normalizeHeader(h)];
      if (key && map[key] == null) map[key] = i;
    });
    return map;
  }

  function rowValue(row, map, key) {
    if (map[key] == null) return "";
    var v = row[map[key]];
    return v == null ? "" : v;
  }

  function tasksFromRows(header, rows, S) {
    var map = mapHeaders(header);
    if (map.Name == null && map.ID == null && map.Start == null) {
      throw new Error("Could not find a Name, ID, or Start column to import.");
    }
    var tasks = [];
    rows.forEach(function (row, i) {
      if (!row || !row.length) return;
      var empty = row.every(function (c) {
        return c == null || String(c).trim() === "";
      });
      if (empty) return;
      var indentVal = rowValue(row, map, "Indent");
      var wbs = String(rowValue(row, map, "WBS") || "");
      var indent = indentVal === "" ? (wbs ? Math.max(0, wbs.split(".").length - 1) : 0) : Number(indentVal);
      var idRaw = rowValue(row, map, "ID");
      var number = Number(idRaw);
      if (!Number.isFinite(number) || number <= 0) number = i + 1;
      tasks.push({
        number: Math.round(number),
        name: String(rowValue(row, map, "Name") || ""),
        start: parseExcelDate(rowValue(row, map, "Start"), S),
        end: parseExcelDate(rowValue(row, map, "End"), S),
        duration: rowValue(row, map, "Duration") === "" ? null : Number(rowValue(row, map, "Duration")),
        progress: rowValue(row, map, "Progress") === "" ? 0 : Number(rowValue(row, map, "Progress")),
        dependencies: S.parseDependencyInput(String(rowValue(row, map, "Dependencies") || "")),
        assignee: String(rowValue(row, map, "Assignee") || ""),
        allocation: rowValue(row, map, "Allocation %") === "" ? 100 : Number(rowValue(row, map, "Allocation %")),
        milestone: parseBool(rowValue(row, map, "Milestone")),
        baselineStart: parseExcelDate(rowValue(row, map, "Baseline Start"), S),
        baselineEnd: parseExcelDate(rowValue(row, map, "Baseline End"), S),
        color: String(rowValue(row, map, "Color") || ""),
        indent: Math.max(0, Math.round(Number(indent) || 0)),
      });
    });
    return tasks;
  }

  function parseConfigSheet(ws, XLSX) {
    var cfg = {};
    if (!ws) return cfg;
    var rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: "" });
    rows.forEach(function (row) {
      var key = String(row[0] || "").trim().toLowerCase();
      var val = row[1];
      if (key === "project name") cfg.projectName = String(val || "");
      if (key === "calendar") {
        var c = String(val || "").toLowerCase();
        cfg.calendarMode = c.indexOf("work") >= 0 ? "workdays" : "calendar";
      }
      if (key === "holidays") {
        cfg.holidays = String(val || "")
          .split(/[,;\s]+/)
          .map(function (s) { return s.trim(); })
          .filter(Boolean);
      }
      if (key === "zoom") cfg.zoom = String(val || "");
    });
    return cfg;
  }

  function parseWorkbook(input, XLSX, S) {
    S = S || (typeof GanttSchedule !== "undefined" ? GanttSchedule : require("./schedule.js"));
    var wb =
      input && input.SheetNames
        ? input
        : XLSX.read(input, { type: typeof input === "string" ? "binary" : "array", cellDates: true, raw: true });
    var sheetName =
      wb.SheetNames.indexOf("Tasks") >= 0
        ? "Tasks"
        : wb.SheetNames.find(function (n) {
            return /task/i.test(n);
          }) || wb.SheetNames[0];
    if (!sheetName) throw new Error("The workbook has no sheets.");
    var rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, raw: true, defval: "" });
    if (!rows.length) throw new Error("The Tasks sheet is empty.");
    var header = rows[0];
    var tasks = tasksFromRows(header, rows.slice(1), S);
    var cfg = parseConfigSheet(wb.Sheets.Config || wb.Sheets.Settings, XLSX);
    return {
      projectName: cfg.projectName || "",
      calendarMode: cfg.calendarMode || "calendar",
      holidays: cfg.holidays || [],
      zoom: cfg.zoom,
      tasks: tasks,
    };
  }

  function parseCSV(text, S) {
    S = S || (typeof GanttSchedule !== "undefined" ? GanttSchedule : require("./schedule.js"));
    var rows = parseCsvText(text);
    if (!rows.length) throw new Error("The CSV file is empty.");
    return { tasks: tasksFromRows(rows[0], rows.slice(1), S) };
  }

  function parseCsvText(text) {
    var rows = [];
    var row = [];
    var cell = "";
    var inQuotes = false;
    var s = String(text || "").replace(/^\uFEFF/, "");
    for (var i = 0; i < s.length; i++) {
      var ch = s[i];
      if (inQuotes) {
        if (ch === '"') {
          if (s[i + 1] === '"') {
            cell += '"';
            i++;
          } else inQuotes = false;
        } else cell += ch;
      } else if (ch === '"') {
        inQuotes = true;
      } else if (ch === ",") {
        row.push(cell);
        cell = "";
      } else if (ch === "\n") {
        row.push(cell);
        rows.push(row);
        row = [];
        cell = "";
      } else if (ch !== "\r") {
        cell += ch;
      }
    }
    if (cell !== "" || row.length) {
      row.push(cell);
      rows.push(row);
    }
    return rows;
  }

  return {
    TASK_HEADERS: TASK_HEADERS,
    buildWorkbook: buildWorkbook,
    writeBuffer: writeBuffer,
    writeNodeBuffer: writeNodeBuffer,
    parseWorkbook: parseWorkbook,
    parseCSV: parseCSV,
    parseCsvText: parseCsvText,
    parseExcelDate: parseExcelDate,
  };
});
