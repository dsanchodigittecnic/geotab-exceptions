(function () {
  "use strict";

  var DEFAULT_FROM_DATE = "2026-02-18T23:00:00.000Z";
  var DEFAULT_TO_DATE = "2026-02-19T22:59:59.000Z";
  var DEFAULT_DIAGNOSTIC_ID = "DiagnosticAux1Id";

  var RULES_BY_DIAGNOSTIC = {
    DiagnosticPowerTakeoffEngagedId: [
      "aMJv_sZ41UUSTPJqNHmcCAw",
      "akWqirOGg_0mvKWrj3iACZQ",
      "ayvMMtxMesEWI40RGFEYDqw",
      "aTMStgBucpUWDxX4ykxsblw"
    ],
    DiagnosticAux1Id: [
      "aMJv_sZ41UUSTPJqNHmcCAw",
      "aCDCWlou1aUqfbbiwsGKdSg",
      "az8GWV5Q7QEmL87I8A0Dmnw",
      "aTMStgBucpUWDxX4ykxsblw"
    ]
  };

  var ctx = {
    api: null,
    initialized: false
  };

  function $(id) {
    return document.getElementById(id);
  }

  function log(message) {
    var logEl = $("log");
    var now = new Date().toISOString().replace("T", " ").slice(0, 19);
    if (!logEl) {
      return;
    }
    logEl.textContent += "[" + now + "] " + message + "\n";
    logEl.scrollTop = logEl.scrollHeight;
  }

  function clearLog() {
    var logEl = $("log");
    if (!logEl) {
      return;
    }
    logEl.textContent = "";
  }

  function setFormEnabled(enabled) {
    var form = $("report-form");
    var nodes = form.querySelectorAll("input, select, button");
    for (var i = 0; i < nodes.length; i += 1) {
      nodes[i].disabled = !enabled;
    }
  }

  function toInputDatetime(iso) {
    return iso.slice(0, 16);
  }

  function inputDatetimeToUtcIso(value) {
    return new Date(value + "Z").toISOString();
  }

  function parseIso(x) {
    var dt = x instanceof Date ? x : new Date(x);
    if (Number.isNaN(dt.getTime())) {
      throw new Error("Fecha invalida: " + x);
    }
    return dt;
  }

  function overlaps(aStart, aEnd, bStart, bEnd) {
    return aStart < bEnd && aEnd > bStart;
  }

  function isOneValue(x) {
    if (x === 1 || x === 1.0 || x === true) {
      return true;
    }
    if (x && typeof x === "object" && (x.value === 1 || x.value === 1.0 || x.value === true)) {
      return true;
    }
    return false;
  }

  function seconds(start, end) {
    return Math.max(0, (end.getTime() - start.getTime()) / 1000);
  }

  function geotabGet(api, typeName, search) {
    return new Promise(function (resolve, reject) {
      api.call(
        "Get",
        {
          typeName: typeName,
          search: search
        },
        function (result) {
          resolve(result || []);
        },
        function (error) {
          reject(error || new Error("Error llamando a Geotab API"));
        }
      );
    });
  }

  async function resolveGroupAuto(api, groupInput) {
    try {
      var byId = await geotabGet(api, "Group", { id: groupInput });
      if (byId.length) {
        return { id: byId[0].id, name: byId[0].name || "", how: "id" };
      }
    } catch (err) {
      // Intento por nombre debajo.
    }

    var byName = await geotabGet(api, "Group", { name: groupInput });
    if (byName.length) {
      return { id: byName[0].id, name: byName[0].name || groupInput, how: "name" };
    }

    throw new Error("No se encontro el grupo (ni por id ni por nombre): " + groupInput);
  }

  async function getDevicesInGroup(api, groupId, includeSubgroups) {
    if (!includeSubgroups) {
      return geotabGet(api, "Device", { groups: [{ id: groupId }] });
    }

    var groupIds = [groupId];
    try {
      var children = await geotabGet(api, "Group", { parent: { id: groupId } });
      for (var i = 0; i < children.length; i += 1) {
        groupIds.push(children[i].id);
      }
    } catch (err) {
      log("No se pudieron consultar subgrupos, se continua con grupo principal.");
    }

    var groupObjs = groupIds.map(function (id) {
      return { id: id };
    });

    return geotabGet(api, "Device", { groups: groupObjs });
  }

  async function getRuleName(api, ruleObjOrId, cache) {
    var ruleId;

    if (ruleObjOrId && typeof ruleObjOrId === "object") {
      if (ruleObjOrId.name) {
        return ruleObjOrId.name;
      }
      ruleId = ruleObjOrId.id;
    } else {
      ruleId = ruleObjOrId;
    }

    if (cache[ruleId]) {
      return cache[ruleId];
    }

    var rules = await geotabGet(api, "Rule", { id: ruleId });
    if (rules.length) {
      cache[ruleId] = rules[0].name;
      return cache[ruleId];
    }

    cache[ruleId] = ruleId;
    return ruleId;
  }

  async function getExceptions(api, deviceId, fromDate, toDate, ruleIds) {
    var events = [];

    for (var i = 0; i < ruleIds.length; i += 1) {
      var ruleId = ruleIds[i];
      var data = await geotabGet(api, "ExceptionEvent", {
        deviceSearch: { id: deviceId },
        ruleSearch: { id: ruleId },
        fromDate: fromDate,
        toDate: toDate
      });

      if (data.length) {
        events = events.concat(data);
      }
    }

    return events;
  }

  async function groupExceptionsByRule(api, exceptions, ruleCache) {
    var counts = {};
    var eventsByRule = {};

    for (var i = 0; i < exceptions.length; i += 1) {
      var e = exceptions[i];
      var name = await getRuleName(api, e.rule, ruleCache);

      counts[name] = (counts[name] || 0) + 1;
      if (!eventsByRule[name]) {
        eventsByRule[name] = [];
      }
      eventsByRule[name].push(e);
    }

    return { counts: counts, eventsByRule: eventsByRule };
  }

  async function getMeasurements(api, deviceId, diagnosticId, fromDate, toDate) {
    var data = await geotabGet(api, "StatusData", {
      deviceSearch: { id: deviceId },
      diagnosticSearch: { id: diagnosticId },
      fromDate: fromDate,
      toDate: toDate
    });

    data.sort(function (a, b) {
      return parseIso(a.dateTime) - parseIso(b.dateTime);
    });

    return data;
  }

  function countSegmentsValue1(status) {
    var intervals = [];
    var inOne = false;
    var start = null;

    if (!status || status.length < 2) {
      return { count: 0, intervals: [] };
    }

    for (var i = 0; i < status.length - 1; i += 1) {
      var cur = status[i];
      var nxt = status[i + 1];

      var curDt = parseIso(cur.dateTime);
      var nxtDt = parseIso(nxt.dateTime);

      var curIs = isOneValue(cur.data);
      var nxtIs = isOneValue(nxt.data);

      if (curIs && !inOne) {
        start = curDt;
        inOne = true;
      }

      if (curIs && !nxtIs && inOne) {
        intervals.push([start, nxtDt]);
        inOne = false;
      }
    }

    return { count: intervals.length, intervals: intervals };
  }

  function segmentsOverlapByRule(eventsByRule, intervals) {
    var result = {};
    var total = intervals.length;
    var ruleNames = Object.keys(eventsByRule);

    for (var i = 0; i < ruleNames.length; i += 1) {
      var ruleName = ruleNames[i];
      var events = eventsByRule[ruleName];
      var count = 0;

      for (var s = 0; s < intervals.length; s += 1) {
        var segStart = intervals[s][0];
        var segEnd = intervals[s][1];

        for (var e = 0; e < events.length; e += 1) {
          var exStart = parseIso(events[e].activeFrom);
          var exEnd = parseIso(events[e].activeTo);
          if (overlaps(exStart, exEnd, segStart, segEnd)) {
            count += 1;
            break;
          }
        }
      }

      result[ruleName] = [total, count];
    }

    return result;
  }

  function autosizeColumns(rows) {
    if (!rows || !rows.length) {
      return [];
    }

    var widths = [];
    for (var c = 0; c < rows[0].length; c += 1) {
      var maxLen = 10;
      for (var r = 0; r < rows.length; r += 1) {
        var cell = rows[r][c];
        var text = cell === null || cell === undefined ? "" : String(cell);
        if (text.length + 2 > maxLen) {
          maxLen = text.length + 2;
        }
      }
      widths.push({ wch: Math.min(70, maxLen) });
    }

    return widths;
  }

  function toDisplayDate(dt) {
    return dt.toISOString().replace("T", " ").slice(0, 19);
  }

  function safeFileName(name) {
    return String(name || "reporte")
      .replace(/[^a-zA-Z0-9._-]/g, "_")
      .replace(/_+/g, "_");
  }

  function escapeHtml(value) {
    return String(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function renderResultsTable(rows) {
    var container = $("tableContainer");
    if (!container) {
      return;
    }

    if (!rows || rows.length < 2) {
      container.className = "table-empty";
      container.textContent = "No hay datos para mostrar.";
      return;
    }

    var header = rows[0];
    var body = rows.slice(1);

    var html = "<div class=\"table-wrap\"><table class=\"results-table\"><thead><tr>";
    for (var h = 0; h < header.length; h += 1) {
      html += "<th>" + escapeHtml(header[h]) + "</th>";
    }
    html += "</tr></thead><tbody>";

    for (var r = 0; r < body.length; r += 1) {
      html += "<tr>";
      for (var c = 0; c < body[r].length; c += 1) {
        var cell = body[r][c];
        if (typeof cell === "number" && c === 6) {
          cell = (cell * 100).toFixed(1) + "%";
        }
        html += "<td>" + escapeHtml(cell === null || cell === undefined ? "" : cell) + "</td>";
      }
      html += "</tr>";
    }

    html += "</tbody></table></div>";
    container.className = "";
    container.innerHTML = html;
  }

  function addSheet(workbook, sheetName, rows) {
    var ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"] = autosizeColumns(rows);
    ws["!autofilter"] = { ref: XLSX.utils.encode_range(XLSX.utils.decode_range(ws["!ref"])) };
    XLSX.utils.book_append_sheet(workbook, ws, sheetName);
  }

  function setPercentColumn(ws, colIndexOneBased, fromRowOneBased, toRowOneBased) {
    var col = colIndexOneBased - 1;
    for (var r = fromRowOneBased - 1; r < toRowOneBased; r += 1) {
      var cellRef = XLSX.utils.encode_cell({ c: col, r: r });
      if (ws[cellRef]) {
        ws[cellRef].z = "0.0%";
      }
    }
  }

  async function generateReportData(params) {
    var diagnosticId = params.diagnosticId;
    if (!RULES_BY_DIAGNOSTIC[diagnosticId]) {
      throw new Error("Diagnostico no soportado: " + diagnosticId);
    }

    var fromDate = params.fromDate;
    var toDate = params.toDate;
    var ruleIds = RULES_BY_DIAGNOSTIC[diagnosticId];
    var api = ctx.api;

    log("Resolviendo grupo...");
    var group = await resolveGroupAuto(api, params.groupInput);

    log("Consultando vehiculos del grupo...");
    var devices = await getDevicesInGroup(api, group.id, params.includeSubgroups);
    if (!devices.length) {
      throw new Error("No se encontraron vehiculos en el grupo: " + params.groupInput);
    }

    var ruleCache = {};
    var ruleNames = [];
    for (var i = 0; i < ruleIds.length; i += 1) {
      ruleNames.push(await getRuleName(api, ruleIds[i], ruleCache));
    }

    var rowsPorVehiculo = [];
    var rowsSegmentos = [];
    var rowsExcepciones = [];

    var totalSegmentsGroup = 0;
    var totalExceptionsByRuleGroup = {};
    var totalOverlapByRuleGroup = {};

    for (i = 0; i < ruleNames.length; i += 1) {
      totalExceptionsByRuleGroup[ruleNames[i]] = 0;
      totalOverlapByRuleGroup[ruleNames[i]] = 0;
    }

    log("Grupo: " + group.name + " | id: " + group.id + " | detectado como: " + group.how);
    log("Diagnostic: " + diagnosticId);
    log("Vehiculos: " + devices.length);
    log("Rango UTC: " + fromDate + " -> " + toDate);

    for (i = 0; i < devices.length; i += 1) {
      var d = devices[i];
      var deviceId = d.id;
      var deviceName = d.name || deviceId;

      var exceptions = await getExceptions(api, deviceId, fromDate, toDate, ruleIds);
      var measurements = await getMeasurements(api, deviceId, diagnosticId, fromDate, toDate);

      var segmentData = countSegmentsValue1(measurements);
      var segmentsCount = segmentData.count;
      var intervals = segmentData.intervals;

      var grouped = await groupExceptionsByRule(api, exceptions, ruleCache);
      var counts = grouped.counts;
      var eventsByRule = grouped.eventsByRule;
      var segmentStats = segmentsOverlapByRule(eventsByRule, intervals);

      totalSegmentsGroup += segmentsCount;

      for (var s = 0; s < intervals.length; s += 1) {
        var segStart = intervals[s][0];
        var segEnd = intervals[s][1];
        var segDurS = seconds(segStart, segEnd);
        rowsSegmentos.push([
          deviceName,
          deviceId,
          toDisplayDate(segStart),
          toDisplayDate(segEnd),
          segDurS,
          segDurS / 60.0
        ]);
      }

      for (var ex = 0; ex < exceptions.length; ex += 1) {
        var exItem = exceptions[ex];
        var rn = await getRuleName(api, exItem.rule, ruleCache);
        var exStart = parseIso(exItem.activeFrom);
        var exEnd = parseIso(exItem.activeTo);
        var exDurS = seconds(exStart, exEnd);

        rowsExcepciones.push([
          deviceName,
          deviceId,
          rn,
          toDisplayDate(exStart),
          toDisplayDate(exEnd),
          exDurS,
          exDurS / 60.0
        ]);
      }

      for (var r = 0; r < ruleNames.length; r += 1) {
        var ruleName = ruleNames[r];
        var excCount = Number(counts[ruleName] || 0);
        var overlapCount = segmentsCount ? Number((segmentStats[ruleName] || [segmentsCount, 0])[1]) : 0;
        var pct = segmentsCount ? overlapCount / segmentsCount : 0;

        rowsPorVehiculo.push([
          deviceName,
          deviceId,
          ruleName,
          segmentsCount,
          excCount,
          overlapCount,
          pct
        ]);

        totalExceptionsByRuleGroup[ruleName] += excCount;
        totalOverlapByRuleGroup[ruleName] += overlapCount;
      }

      log("Procesado: " + deviceName + " | segmentos: " + segmentsCount + " | excepciones: " + exceptions.length);
    }

    var rows0 = [
      ["Campo", "Valor"],
      ["Grupo (input)", params.groupInput],
      ["Detectado como", group.how],
      ["Grupo (name)", group.name],
      ["Grupo (id)", group.id],
      ["Include subgrupos", String(params.includeSubgroups)],
      ["Desde (UTC)", fromDate],
      ["Hasta (UTC)", toDate],
      ["Diagnostic ID", diagnosticId],
      ["Reglas (nombres)", ruleNames.join(", ")]
    ];
    addSheet(wb, "00_Parametros", rows0);

    var rows1 = [["Regla", "Excepciones (grupo)", "Segmentos valor=1 (grupo)", "Segmentos con solape", "% solape (solape/segmentos)"]];
    for (i = 0; i < ruleNames.length; i += 1) {
      var rnGroup = ruleNames[i];
      var excGroup = totalExceptionsByRuleGroup[rnGroup] || 0;
      var ovGroup = totalOverlapByRuleGroup[rnGroup] || 0;
      var pctGroup = totalSegmentsGroup ? ovGroup / totalSegmentsGroup : 0;
      rows1.push([rnGroup, excGroup, totalSegmentsGroup, ovGroup, pctGroup]);
    }
    addSheet(wb, "01_Resumen_grupo", rows1);

    var rows2 = [["Vehiculo", "Device ID", "Regla", "Segmentos valor=1", "Excepciones", "Segmentos con solape", "% solape"]].concat(rowsPorVehiculo);
    addSheet(wb, "02_Por_vehiculo", rows2);

    var rows3 = [["Vehiculo", "Device ID", "Inicio segmento", "Fin segmento", "Duracion (s)", "Duracion (min)"]].concat(rowsSegmentos);
    addSheet(wb, "03_Segmentos_1", rows3);

    var rows4 = [["Vehiculo", "Device ID", "Regla", "Inicio excepcion", "Fin excepcion", "Duracion (s)", "Duracion (min)"]].concat(rowsExcepciones);
    return {
      rows0: rows0,
      rows1: rows1,
      rows2: rows2,
      rows3: rows3,
      rows4: rows4,
      groupName: group.name,
      fromDate: fromDate,
      toDate: toDate,
      diagnosticId: diagnosticId
    };
  }

  function downloadExcel(params, reportData) {
    var wb = XLSX.utils.book_new();
    addSheet(wb, "00_Parametros", reportData.rows0);
    addSheet(wb, "01_Resumen_grupo", reportData.rows1);
    addSheet(wb, "02_Por_vehiculo", reportData.rows2);
    addSheet(wb, "03_Segmentos_1", reportData.rows3);
    addSheet(wb, "04_Excepciones", reportData.rows4);

    var ws1 = wb.Sheets["01_Resumen_grupo"];
    var ws2 = wb.Sheets["02_Por_vehiculo"];
    setPercentColumn(ws1, 5, 2, reportData.rows1.length);
    setPercentColumn(ws2, 7, 2, reportData.rows2.length);

    var safeGroup = safeFileName(reportData.groupName || params.groupInput);
    var outName = (params.outFile || "").trim();
    if (!outName) {
      outName = "geotab_" + safeGroup + "_" + reportData.fromDate.slice(0, 10) + "_" + reportData.toDate.slice(0, 10) + "_" + reportData.diagnosticId + ".xlsx";
    }
    if (!/\.xlsx$/i.test(outName)) {
      outName += ".xlsx";
    }

    XLSX.writeFile(wb, outName, { compression: true });
    log("Excel generado: " + outName);
  }

  function readParamsFromForm() {
    var fromValue = $("fromInput").value;
    var toValue = $("toInput").value;

    if (!fromValue || !toValue) {
      throw new Error("Debes completar las fechas Desde/Hasta.");
    }

    var fromDate = inputDatetimeToUtcIso(fromValue);
    var toDate = inputDatetimeToUtcIso(toValue);

    if (parseIso(fromDate) >= parseIso(toDate)) {
      throw new Error("El rango es invalido: Desde debe ser menor que Hasta.");
    }

    return {
      groupInput: $("groupInput").value.trim(),
      diagnosticId: $("diagnosticInput").value,
      fromDate: fromDate,
      toDate: toDate,
      includeSubgroups: $("includeSubgroupsInput").checked,
      outFile: $("outFileInput").value.trim()
    };
  }

  function setupUi() {
    if (ctx.initialized) {
      return;
    }

    $("fromInput").value = toInputDatetime(DEFAULT_FROM_DATE);
    $("toInput").value = toInputDatetime(DEFAULT_TO_DATE);
    $("diagnosticInput").value = DEFAULT_DIAGNOSTIC_ID;

    $("report-form").addEventListener("submit", async function (ev) {
      ev.preventDefault();

      clearLog();
      try {
        var params = readParamsFromForm();
        if (!params.groupInput) {
          throw new Error("Debes indicar un grupo (id o nombre).");
        }

        setFormEnabled(false);
        log("Iniciando generacion del reporte...");
        var reportData = await generateReportData(params);
        downloadExcel(params, reportData);
      } catch (err) {
        log("ERROR: " + (err && err.message ? err.message : String(err)));
      } finally {
        setFormEnabled(true);
      }
    });

    $("tableBtn").addEventListener("click", async function () {
      clearLog();
      try {
        var params = readParamsFromForm();
        if (!params.groupInput) {
          throw new Error("Debes indicar un grupo (id o nombre).");
        }

        setFormEnabled(false);
        log("Iniciando generacion de tabla...");
        var reportData = await generateReportData(params);
        renderResultsTable(reportData.rows2);
      } catch (err) {
        log("ERROR: " + (err && err.message ? err.message : String(err)));
      } finally {
        setFormEnabled(true);
      }
    });

    ctx.initialized = true;
    log("Add-in listo.");
  }

  function buildAddin() {
    return {
      initialize: function (api, state, callback) {
        ctx.api = api;
        setupUi();
        if (typeof callback === "function") {
          callback();
        }
      },
      focus: function () {
        log("Add-in en foco.");
      },
      blur: function () {
        // Sin estado temporal que limpiar.
      }
    };
  }

  if (window.geotab && window.geotab.addin) {
    window.geotab.addin.exceptionsReportAddin = buildAddin;
  } else {
    // Modo local para maquetar UI.
    document.addEventListener("DOMContentLoaded", function () {
      setupUi();
      log("No se detecto runtime de Geotab. Esta vista sirve solo para UI local.");
    });
  }
})();
