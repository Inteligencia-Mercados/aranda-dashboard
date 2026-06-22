/* =============================================================================
   Centro de Monitoreo Operativo · Aranda Service Desk · UAO
   app.js — Lógica de datos, cálculo de KPIs, gráficos y tablas.

   FLUJO DE DATOS (ver también el comentario en index.html):
   SharePoint Online → Power Automate → GitHub Pages (JSON) → este dashboard.
   Este dashboard SOLO lee los JSON listados en CONFIG.dataSources mediante
   fetch (con cache-busting), cada CONFIG.refreshIntervalMs, sin recarga de
   página. Para producción, basta reemplazar las rutas de CONFIG.dataSources
   por la URL pública de GitHub Pages donde Power Automate publica los JSON.
   ============================================================================= */

(function () {
  "use strict";

  /* ============================ CONFIGURACIÓN ============================ */

  const CONFIG = {
    dataSources: {
      "PQRS": "data/PQRS.json",
      "Servicios Financieros": "data/ServiciosFinancieros.json",
      "CAE": "data/CAE.json"
    },
    refreshIntervalMs: 5 * 60 * 1000 // 5 minutos
  };

  const AREAS = ["PQRS", "Servicios Financieros", "CAE"];
  const AREA_PREFIX = { "PQRS": "pqrs", "Servicios Financieros": "sf", "CAE": "cae" };
  const AREA_SHORT = { "PQRS": "PQRS", "Servicios Financieros": "SF", "CAE": "CAE" };
  const AREA_COLORS = { "PQRS": "#C0151A", "Servicios Financieros": "#8C0F13", "CAE": "#6B5E54" };

  const STATUS_COLORS = { "Normal": "#9C8C7E", "Riesgo": "#D9A441", "Critico": "#C0151A", "Vencido": "#4A0608" };
  const STATUS_LABELS = { "Normal": "Normal", "Riesgo": "Riesgo", "Critico": "Crítico", "Vencido": "Vencido" };

  const SERIES_PALETTE = ["#8C0F13", "#C0151A", "#D9A441", "#9C8C7E", "#4A0608", "#B5654A", "#6B5E54", "#D9B68B", "#7A1E22", "#C98A3E"];

  /* Campos de la barra de filtros globales */
  const GF_FIELDS = [
    { key: "estado",      label: "Estado",      field: "Estado",      icon: "bi-circle-half" },
    { key: "area",        label: "Área",         field: "_area",       icon: "bi-diagram-2" },
    { key: "categoria",   label: "Categoría",   field: "Categoría",   icon: "bi-tags" },
    { key: "servicio",    label: "Servicio",     field: "Servicio",    icon: "bi-diagram-3" },
    { key: "responsable", label: "Responsable",  field: "Responsable", icon: "bi-person" },
    { key: "grupo",       label: "Grupo",        field: "Grupo",       icon: "bi-building" },
    { key: "urgencia",    label: "Urgencia",     field: "Urgencia",    icon: "bi-flag" }
  ];

  function classify(progreso) {
    if (progreso > 100) return "Vencido";
    if (progreso >= 90) return "Critico";
    if (progreso >= 70) return "Riesgo";
    return "Normal";
  }

  function effectiveClass(r) {
    return r["Estado"] === "Solucionado" ? "Normal" : classify(r["Progreso"]);
  }

  const VIEW_TITLES = {
    resumen: "Resumen ejecutivo",
    responsables: "Gestión de Responsables",
    atencion: "Necesita atención hoy",
    solucionados: "Casos solucionados",
    cae: "CAE",
    pqrs: "PQRS",
    sf: "Servicios Financieros"
  };

  const DT_LANG_ES = {
    search: "Buscar:",
    lengthMenu: "Mostrar _MENU_ registros",
    info: "Mostrando _START_ a _END_ de _TOTAL_ registros",
    infoEmpty: "Mostrando 0 a 0 de 0 registros",
    infoFiltered: "(filtrado de _MAX_ registros totales)",
    zeroRecords: "No se encontraron registros coincidentes",
    emptyTable: "No hay datos disponibles",
    paginate: { first: "Primero", last: "Último", next: "Siguiente", previous: "Anterior" },
    processing: "Procesando…"
  };

  /* ============================== ESTADO ============================== */

  const STATE = {
    rawData: {},       // datos sin filtrar, tal como llegan del JSON
    data: {},          // datos después de aplicar GLOBAL_FILTER
    errors: {},
    prevErrors: {},
    stats: {},
    combinedRecords: [],
    statsCombined: {},
    lastUpdated: null,
    firstLoadDone: false
  };

  /* Estado del filtro global: arrays vacíos = sin filtro (todos) */
  const GLOBAL_FILTER = {
    estado: [],
    area: [],
    categoria: [],
    servicio: [],
    responsable: [],
    grupo: [],
    urgencia: [],
    fechaDesde: "",
    fechaHasta: ""
  };

  const chartRegistry = {};
  const dtRegistry = {};
  let _respDetalleActual = null; // nombre del responsable actualmente abierto en el panel de detalle
  const RESP_SECTION_FILTER = { responsable: [], grupo: [] };
  let TENDENCY_PERIOD = "semana"; // "semana" | "mes" | "año"

  /* ============================ UTILIDADES ============================ */

  function esc(str) {
    if (str === null || str === undefined) return "";
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function pct(n, total) {
    if (!total) return 0;
    return Math.round((n / total) * 1000) / 10;
  }

  function countBy(records, field) {
    const counts = {};
    records.forEach(function (r) {
      const key = r[field] === undefined || r[field] === null || r[field] === "" ? "Sin dato" : r[field];
      counts[key] = (counts[key] || 0) + 1;
    });
    return counts;
  }

  function topEntry(countsObj) {
    let bestKey = null, bestCount = -1;
    Object.keys(countsObj).forEach(function (k) {
      if (countsObj[k] > bestCount) { bestCount = countsObj[k]; bestKey = k; }
    });
    return bestKey === null ? { key: "—", count: 0 } : { key: bestKey, count: bestCount };
  }

  function sortedEntries(countsObj, limit) {
    const entries = Object.keys(countsObj).map(function (k) { return [k, countsObj[k]]; });
    entries.sort(function (a, b) { return b[1] - a[1]; });
    return limit ? entries.slice(0, limit) : entries;
  }

  function toChartDataDoughnut(counts, labelMap, colorMap) {
    const keys = Object.keys(counts);
    return {
      labels: keys.map(function (k) { return labelMap ? (labelMap[k] || k) : k; }),
      datasets: [{
        data: keys.map(function (k) { return counts[k]; }),
        backgroundColor: keys.map(function (k, i) { return colorMap ? (colorMap[k] || SERIES_PALETTE[i % SERIES_PALETTE.length]) : SERIES_PALETTE[i % SERIES_PALETTE.length]; }),
        borderWidth: 0
      }]
    };
  }

  function toChartDataBar(counts, color, limit) {
    const entries = sortedEntries(counts, limit || 10);
    return {
      labels: entries.map(function (e) { return e[0]; }),
      datasets: [{
        data: entries.map(function (e) { return e[1]; }),
        backgroundColor: color || "#8C0F13",
        borderRadius: 4,
        maxBarThickness: 26
      }]
    };
  }

  function toChartDataBarMulti(counts, limit) {
    const entries = sortedEntries(counts, limit || 10);
    return {
      labels: entries.map(function (e) { return e[0]; }),
      datasets: [{
        data: entries.map(function (e) { return e[1]; }),
        backgroundColor: entries.map(function (e, i) { return SERIES_PALETTE[i % SERIES_PALETTE.length]; }),
        borderRadius: 4,
        maxBarThickness: 26
      }]
    };
  }

  function weekLabel(d) {
    return d.toLocaleDateString("es-CO", { day: "2-digit", month: "short" });
  }

  function getWeeklyCounts(records, weeksBack) {
    weeksBack = weeksBack || 12;
    const now = new Date();
    const buckets = [];
    for (let i = weeksBack - 1; i >= 0; i--) {
      const end = new Date(now);
      end.setDate(now.getDate() - i * 7);
      const start = new Date(end);
      start.setDate(end.getDate() - 6);
      buckets.push({ start: start, end: end, label: weekLabel(start), count: 0 });
    }
    records.forEach(function (r) {
      const fr = r["Fecha de registro"];
      if (!fr) return;
      const d = new Date(fr + "T00:00:00");
      for (let i = 0; i < buckets.length; i++) {
        if (d >= buckets[i].start && d <= buckets[i].end) { buckets[i].count++; break; }
      }
    });
    return buckets;
  }

  function getDailyCounts(records, daysBack) {
    daysBack = daysBack || 30;
    const now = new Date();
    const buckets = [];
    for (let i = daysBack - 1; i >= 0; i--) {
      const d = new Date(now);
      d.setDate(now.getDate() - i);
      const y = d.getFullYear();
      const m = String(d.getMonth() + 1).padStart(2, "0");
      const day = String(d.getDate()).padStart(2, "0");
      const dateStr = y + "-" + m + "-" + day;
      const label = d.toLocaleDateString("es-CO", { day: "2-digit", month: "short" });
      buckets.push({ dateStr: dateStr, label: label, count: 0 });
    }
    records.forEach(function (r) {
      const fr = r["Fecha de registro"];
      if (!fr) return;
      for (let i = 0; i < buckets.length; i++) {
        if (fr === buckets[i].dateStr) { buckets[i].count++; break; }
      }
    });
    return buckets;
  }

  function getMonthlyCounts(records, monthsBack) {
    const now = new Date();
    const buckets = [];
    for (let i = monthsBack - 1; i >= 0; i--) {
      const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
      const end = new Date(now.getFullYear(), now.getMonth() - i + 1, 0);
      buckets.push({
        start: d, end: end,
        label: d.toLocaleDateString("es-CO", { month: "short", year: "2-digit" }),
        count: 0
      });
    }
    records.forEach(function (r) {
      const fr = r["Fecha de registro"];
      if (!fr) return;
      const d = new Date(fr + "T00:00:00");
      for (let i = 0; i < buckets.length; i++) {
        if (d >= buckets[i].start && d <= buckets[i].end) { buckets[i].count++; break; }
      }
    });
    return buckets;
  }

  function getYearlyCounts(records) {
    const map = {};
    records.forEach(function (r) {
      const fr = r["Fecha de registro"];
      if (!fr) return;
      map[fr.substring(0, 4)] = (map[fr.substring(0, 4)] || 0) + 1;
    });
    return Object.keys(map).sort().map(function (y) { return { label: y, count: map[y] }; });
  }

  function isoToday() {
    const d = new Date();
    return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0");
  }

  function getDailyCountsRange(records, fromISO, toISO) {
    const from = new Date(fromISO + "T00:00:00");
    const to   = new Date(toISO   + "T00:00:00");
    const buckets = [];
    const cur = new Date(from);
    while (cur <= to) {
      const y = cur.getFullYear(), m = String(cur.getMonth() + 1).padStart(2, "0"), d = String(cur.getDate()).padStart(2, "0");
      buckets.push({ dateStr: y + "-" + m + "-" + d, label: cur.toLocaleDateString("es-CO", { day: "2-digit", month: "short" }), count: 0 });
      cur.setDate(cur.getDate() + 1);
    }
    records.forEach(function (r) {
      const fr = r["Fecha de registro"];
      if (!fr) return;
      for (let i = 0; i < buckets.length; i++) { if (fr === buckets[i].dateStr) { buckets[i].count++; break; } }
    });
    return buckets;
  }

  function getWeeklyCountsRange(records, fromISO, toISO) {
    const from = new Date(fromISO + "T00:00:00");
    const to   = new Date(toISO   + "T00:00:00");
    const buckets = [];
    const cur = new Date(from);
    while (cur <= to) {
      const wEnd = new Date(cur);
      wEnd.setDate(cur.getDate() + 6);
      buckets.push({ start: new Date(cur), end: wEnd > to ? new Date(to) : new Date(wEnd), label: weekLabel(cur), count: 0 });
      cur.setDate(cur.getDate() + 7);
    }
    records.forEach(function (r) {
      const fr = r["Fecha de registro"];
      if (!fr) return;
      const d = new Date(fr + "T00:00:00");
      for (let i = 0; i < buckets.length; i++) { if (d >= buckets[i].start && d <= buckets[i].end) { buckets[i].count++; break; } }
    });
    return buckets;
  }

  function getMonthlyCountsRange(records, fromISO, toISO) {
    const from = new Date(fromISO + "T00:00:00");
    const to   = new Date(toISO   + "T00:00:00");
    const buckets = [];
    let yr = from.getFullYear(), mo = from.getMonth();
    const eYr = to.getFullYear(), eMo = to.getMonth();
    while (yr < eYr || (yr === eYr && mo <= eMo)) {
      const start = new Date(yr, mo, 1), end = new Date(yr, mo + 1, 0);
      buckets.push({ start: start, end: end, label: start.toLocaleDateString("es-CO", { month: "short", year: "2-digit" }), count: 0 });
      mo++; if (mo > 11) { mo = 0; yr++; }
    }
    records.forEach(function (r) {
      const fr = r["Fecha de registro"];
      if (!fr) return;
      const d = new Date(fr + "T00:00:00");
      for (let i = 0; i < buckets.length; i++) { if (d >= buckets[i].start && d <= buckets[i].end) { buckets[i].count++; break; } }
    });
    return buckets;
  }

  function getTendenciaCounts(records, period) {
    if (period === "año") return getYearlyCounts(records);
    const hoy = isoToday();
    const hasta = GLOBAL_FILTER.fechaHasta || hoy;
    let desde = GLOBAL_FILTER.fechaDesde;
    if (!desde) {
      const d = new Date();
      if (period === "día")  d.setDate(d.getDate() - 29);
      else if (period === "mes") { d.setMonth(d.getMonth() - 11); d.setDate(1); }
      else d.setDate(d.getDate() - 83); // 12 semanas
      desde = d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0");
    }
    if (period === "día")  return getDailyCountsRange(records, desde, hasta);
    if (period === "mes")  return getMonthlyCountsRange(records, desde, hasta);
    return getWeeklyCountsRange(records, desde, hasta);
  }

  function buildPeriodBtnsHTML() {
    return '<div class="tend-period-btns">' +
      [["día", "Días"], ["semana", "Semanas"], ["mes", "Meses"], ["año", "Años"]].map(function (p) {
        return '<button class="tend-btn' + (TENDENCY_PERIOD === p[0] ? ' tend-btn--active' : '') +
          '" data-period="' + p[0] + '">' + p[1] + '</button>';
      }).join("") +
    '</div>';
  }

  function wireTendencyBtns(container) {
    if (!container) return;
    container.querySelectorAll(".tend-btn").forEach(function (btn) {
      btn.addEventListener("click", function () {
        TENDENCY_PERIOD = this.getAttribute("data-period");
        renderAll();
      });
    });
  }

  function kpi(label, value, variant, icon, foot) {
    return (
      '<div class="kpi-card kpi-card--' + variant + '">' +
        '<div class="kpi-icon"><i class="bi ' + icon + '"></i></div>' +
        '<div class="kpi-body">' +
          '<div class="kpi-value">' + esc(value) + '</div>' +
          '<div class="kpi-label">' + esc(label) + '</div>' +
          (foot ? '<div class="kpi-foot">' + esc(foot) + '</div>' : '') +
        '</div>' +
      '</div>'
    );
  }

  function progressCellHTML(progreso, cls) {
    const pctWidth = Math.max(4, Math.min(100, progreso));
    return (
      '<div class="progress-track">' +
        '<div class="progress-fill progress-fill--' + cls.toLowerCase() + '" style="width:' + pctWidth + '%"></div>' +
      '</div>' +
      '<div class="progress-text">' + progreso.toFixed(1) + '%</div>'
    );
  }

  function areaChipHTML(area) {
    return '<span class="area-chip" style="--chip-color:' + (AREA_COLORS[area] || '#9C8C7E') + '">' + esc(AREA_SHORT[area] || area) + '</span>';
  }

  /* ====================== FILTROS GLOBALES ====================== */

  function applyGlobalFilters() {
    AREAS.forEach(function (area) {
      const raw = STATE.rawData[area] || [];
      STATE.data[area] = raw.filter(function (r) {
        if (GLOBAL_FILTER.estado.length    && GLOBAL_FILTER.estado.indexOf(r["Estado"])     === -1) return false;
        if (GLOBAL_FILTER.area.length       && GLOBAL_FILTER.area.indexOf(r["_area"])        === -1) return false;
        if (GLOBAL_FILTER.categoria.length  && GLOBAL_FILTER.categoria.indexOf(r["Categoría"]) === -1) return false;
        if (GLOBAL_FILTER.servicio.length     && GLOBAL_FILTER.servicio.indexOf(r["Servicio"])       === -1) return false;
        if (GLOBAL_FILTER.responsable.length  && GLOBAL_FILTER.responsable.indexOf(r["Responsable"]) === -1) return false;
        if (GLOBAL_FILTER.grupo.length        && GLOBAL_FILTER.grupo.indexOf(r["Grupo"])             === -1) return false;
        if (GLOBAL_FILTER.urgencia.length     && GLOBAL_FILTER.urgencia.indexOf(r["Urgencia"])       === -1) return false;
        if (GLOBAL_FILTER.fechaDesde && (r["Fecha de registro"] || "") < GLOBAL_FILTER.fechaDesde) return false;
        if (GLOBAL_FILTER.fechaHasta && (r["Fecha de registro"] || "") > GLOBAL_FILTER.fechaHasta) return false;
        return true;
      });
    });
  }

  function onFilterChange() {
    applyGlobalFilters();
    renderAll();
    updateFilterBadges();
  }

  /* Construye el HTML de un multi-select dropdown */
  function buildMsDropHTML(key, label, icon, options, filterObj) {
    const sel = (filterObj || GLOBAL_FILTER)[key] || [];
    const badgeVis = (sel.length > 0 && sel.length < options.length) ? "" : "display:none";
    const optsHtml = options.map(function (opt) {
      const checked = sel.indexOf(opt) !== -1 ? " checked" : "";
      return (
        '<label class="ms-opt">' +
          '<input type="checkbox" class="ms-cb" value="' + esc(opt) + '"' + checked + '>' +
          '<span>' + esc(opt) + '</span>' +
        '</label>'
      );
    }).join("");
    const allChecked = (sel.length === 0 || (options.length > 0 && sel.length === options.length)) ? " checked" : "";
    return (
      '<div class="ms-drop" data-key="' + key + '">' +
        '<button class="ms-toggle" type="button">' +
          '<i class="bi ' + icon + '"></i>' +
          '<span class="ms-label">' + label + '</span>' +
          '<span class="ms-badge" style="' + badgeVis + '">' + sel.length + '</span>' +
          '<i class="bi bi-chevron-down ms-chevron"></i>' +
        '</button>' +
        '<div class="ms-panel" hidden>' +
          '<input class="ms-search" type="text" placeholder="Buscar…" autocomplete="off">' +
          '<div class="ms-opts-wrap">' +
            '<label class="ms-opt ms-opt--all">' +
              '<input type="checkbox" class="ms-cb-all"' + allChecked + '>' +
              '<span>Todos</span>' +
            '</label>' +
            optsHtml +
          '</div>' +
        '</div>' +
      '</div>'
    );
  }

  function initEstadoFilter() {
    const estadoSet = new Set();
    AREAS.forEach(function (a) {
      (STATE.rawData[a] || []).forEach(function (r) {
        if (r["Estado"]) estadoSet.add(r["Estado"]);
      });
    });
    GLOBAL_FILTER.estado = Array.from(estadoSet).sort();
  }

  function populateGlobalFilterBar() {
    const bar = document.getElementById("globalFilterBar");
    if (!bar) return;

    /* Recolectar valores únicos desde rawData (datos sin filtrar) */
    const allRaw = [];
    AREAS.forEach(function (area) {
      (STATE.rawData[area] || []).forEach(function (r) { allRaw.push(r); });
    });
    function uniqueVals(field) {
      return Array.from(new Set(allRaw.map(function (r) { return r[field] || ""; }).filter(Boolean))).sort();
    }

    const dropsHtml = GF_FIELDS.map(function (f) {
      return buildMsDropHTML(f.key, f.label, f.icon, uniqueVals(f.field));
    }).join("");

    bar.innerHTML =
      '<div class="gfb-inner">' +
        '<span class="gfb-title"><i class="bi bi-funnel-fill"></i> Filtros</span>' +
        '<div class="gfb-drops" id="gfbDrops">' + dropsHtml + '</div>' +
        '<div class="gfb-dates">' +
          '<div class="filter-group">' +
            '<label for="gfFechaDesde">Desde</label>' +
            '<input type="date" id="gfFechaDesde" class="filter-select filter-select--sm"' +
              (GLOBAL_FILTER.fechaDesde ? ' value="' + GLOBAL_FILTER.fechaDesde + '"' : '') + '>' +
          '</div>' +
          '<div class="filter-group">' +
            '<label for="gfFechaHasta">Hasta</label>' +
            '<input type="date" id="gfFechaHasta" class="filter-select filter-select--sm"' +
              (GLOBAL_FILTER.fechaHasta ? ' value="' + GLOBAL_FILTER.fechaHasta + '"' : '') + '>' +
          '</div>' +
        '</div>' +
        '<button class="gfb-clear" id="gfbClear" title="Limpiar todos los filtros">' +
          '<i class="bi bi-x-circle"></i> Limpiar' +
        '</button>' +
      '</div>';

    wireGlobalFilterBar();
  }

  function wireGlobalFilterBar() {
    /* Multi-select dropdowns — sólo los del filtro global (dentro de #gfbDrops) */
    document.querySelectorAll("#gfbDrops .ms-drop").forEach(function (drop) {
      const key = drop.getAttribute("data-key");
      const toggle = drop.querySelector(".ms-toggle");
      const panel = drop.querySelector(".ms-panel");
      const searchEl = drop.querySelector(".ms-search");
      const allCb = drop.querySelector(".ms-cb-all");
      const badge = drop.querySelector(".ms-badge");

      toggle.addEventListener("click", function (e) {
        e.stopPropagation();
        const isOpen = !panel.hidden;
        closeAllDropdowns();
        if (!isOpen) {
          panel.hidden = false;
          drop.classList.add("is-open");
          if (searchEl) { searchEl.value = ""; filterDropOptions(drop, ""); searchEl.focus(); }
        }
      });

      // Keep panel open when clicking inside it
      panel.addEventListener("click", function (e) { e.stopPropagation(); });

      if (searchEl) {
        searchEl.addEventListener("input", function () { filterDropOptions(drop, this.value); });
        searchEl.addEventListener("click", function (e) { e.stopPropagation(); });
      }

      if (allCb) {
        allCb.addEventListener("change", function () {
          if (this.checked) {
            drop.querySelectorAll(".ms-cb").forEach(function (cb) { cb.checked = false; });
            GLOBAL_FILTER[key] = [];
            badge.style.display = "none";
            badge.textContent = "0";
            onFilterChange();
          } else {
            this.checked = true; // no se puede desmarcar "Todos" sin seleccionar algo
          }
        });
      }

      drop.querySelectorAll(".ms-cb").forEach(function (cb) {
        cb.addEventListener("change", function () {
          const vals = [];
          drop.querySelectorAll(".ms-cb:checked").forEach(function (c) { vals.push(c.value); });
          GLOBAL_FILTER[key] = vals;
          const totalOpts = drop.querySelectorAll(".ms-cb").length;
          if (allCb) allCb.checked = (vals.length === 0 || vals.length === totalOpts);
          badge.textContent = vals.length;
          badge.style.display = (vals.length > 0 && vals.length < totalOpts) ? "" : "none";
          onFilterChange();
        });
      });
    });

    /* Fechas */
    ["gfFechaDesde", "gfFechaHasta"].forEach(function (id) {
      const el = document.getElementById(id);
      if (el) {
        el.addEventListener("change", function () {
          GLOBAL_FILTER.fechaDesde = document.getElementById("gfFechaDesde") ? document.getElementById("gfFechaDesde").value : "";
          GLOBAL_FILTER.fechaHasta = document.getElementById("gfFechaHasta") ? document.getElementById("gfFechaHasta").value : "";
          onFilterChange();
        });
      }
    });

    /* Botón limpiar */
    const clearBtn = document.getElementById("gfbClear");
    if (clearBtn) {
      clearBtn.addEventListener("click", function () {
        GF_FIELDS.forEach(function (f) { GLOBAL_FILTER[f.key] = []; });
        GLOBAL_FILTER.fechaDesde = "";
        GLOBAL_FILTER.fechaHasta = "";
        initEstadoFilter(); // re-inicializa Estado con todos los valores seleccionados
        populateGlobalFilterBar(); // reconstruye con estado limpio
        onFilterChange();
      });
    }

    /* Cerrar al hacer clic fuera */
    if (!document._gfbOutside) {
      document._gfbOutside = true;
      document.addEventListener("click", closeAllDropdowns);
    }
  }

  function closeAllDropdowns() {
    document.querySelectorAll(".ms-panel").forEach(function (p) { p.hidden = true; });
    document.querySelectorAll(".ms-drop").forEach(function (d) { d.classList.remove("is-open"); });
  }

  function filterDropOptions(drop, q) {
    const lower = q.toLowerCase();
    drop.querySelectorAll(".ms-opt:not(.ms-opt--all)").forEach(function (opt) {
      opt.style.display = opt.textContent.trim().toLowerCase().indexOf(lower) !== -1 ? "" : "none";
    });
  }

  function updateFilterBadges() {
    const total = GF_FIELDS.reduce(function (acc, f) { return acc + (GLOBAL_FILTER[f.key] || []).length; }, 0) +
                  (GLOBAL_FILTER.fechaDesde ? 1 : 0) + (GLOBAL_FILTER.fechaHasta ? 1 : 0);
    const clearBtn = document.getElementById("gfbClear");
    if (clearBtn) clearBtn.classList.toggle("has-filters", total > 0);
  }

  /* ====================== CÁLCULO DE ESTADÍSTICAS ====================== */

  function computeStats(records) {
    const total = records.length;
    let vencidos = 0, criticos = 0, riesgo = 0, normal = 0, sumTiempo = 0;
    const vencidosPorResponsable = {};
    const vencidosPorCategoria = {};
    const vencidosPorDepartamento = {};

    records.forEach(function (r) {
      const cls = effectiveClass(r);
      if (cls === "Vencido") vencidos++;
      else if (cls === "Critico") criticos++;
      else if (cls === "Riesgo") riesgo++;
      else normal++;
      sumTiempo += (r["Tiempo transcurrido"] || 0);

      if (cls === "Vencido") {
        const resp = r["Responsable"] || "Sin asignar";
        const cat = r["Categoría"] || "Sin categoría";
        const dep = r["Departamento"] || "Sin departamento";
        vencidosPorResponsable[resp] = (vencidosPorResponsable[resp] || 0) + 1;
        vencidosPorCategoria[cat] = (vencidosPorCategoria[cat] || 0) + 1;
        vencidosPorDepartamento[dep] = (vencidosPorDepartamento[dep] || 0) + 1;
      }
    });

    const cumplidos = total - vencidos;
    return {
      total: total,
      vencidos: vencidos,
      criticos: criticos,
      riesgo: riesgo,
      normal: normal,
      slaPct: pct(cumplidos, total),
      avgTiempo: total ? Math.round((sumTiempo / total) * 10) / 10 : 0,
      vencidosPorResponsable: vencidosPorResponsable,
      vencidosPorCategoria: vencidosPorCategoria,
      vencidosPorDepartamento: vencidosPorDepartamento
    };
  }

  function computeAllStats() {
    let combined = [];
    AREAS.forEach(function (area) {
      const records = STATE.data[area] || [];
      STATE.stats[area] = computeStats(records);
      combined = combined.concat(records);
    });
    STATE.combinedRecords = combined;
    STATE.statsCombined = computeStats(combined);
  }

  /* ============================ CHART.JS ============================ */

  function setChartDefaults() {
    if (typeof Chart === "undefined") return;
    Chart.defaults.font.family = "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif";
    Chart.defaults.font.size = 11.5;
    Chart.defaults.color = "#4A3F38";
  }

  function renderChart(canvasId, type, data, options) {
    const el = document.getElementById(canvasId);
    if (!el) return null;
    if (chartRegistry[canvasId]) { chartRegistry[canvasId].destroy(); }
    const ctx = el.getContext("2d");
    el.classList.remove("chart-skeleton");
    chartRegistry[canvasId] = new Chart(ctx, { type: type, data: data, options: options || {} });
    return chartRegistry[canvasId];
  }

  function gridOpts() { return { color: "#EEEBE7", drawBorder: false }; }

  function horizontalBarOpts(extra) {
    const base = {
      indexAxis: "y",
      plugins: { legend: { display: false } },
      scales: {
        x: { grid: gridOpts(), beginAtZero: true, ticks: { precision: 0 } },
        y: { grid: { display: false } }
      },
      maintainAspectRatio: false
    };
    return Object.assign(base, extra || {});
  }

  function barOpts(extra) {
    const base = {
      plugins: { legend: { display: false } },
      scales: {
        x: { grid: { display: false } },
        y: { grid: gridOpts(), beginAtZero: true, ticks: { precision: 0 } }
      },
      maintainAspectRatio: false
    };
    return Object.assign(base, extra || {});
  }

  function stackedBarOpts() {
    return {
      plugins: { legend: { position: "bottom", labels: { boxWidth: 11, boxHeight: 11, padding: 14 } } },
      scales: {
        x: { stacked: true, grid: { display: false } },
        y: { stacked: true, grid: gridOpts(), beginAtZero: true, ticks: { precision: 0 } }
      },
      maintainAspectRatio: false
    };
  }

  function doughnutOpts() {
    return {
      cutout: "62%",
      plugins: { legend: { position: "bottom", labels: { boxWidth: 11, boxHeight: 11, padding: 14 } } },
      maintainAspectRatio: false
    };
  }

  function lineOpts(extra) {
    const base = {
      interaction: { mode: "index", intersect: false },
      plugins: {
        legend: { position: "bottom", labels: { boxWidth: 11, boxHeight: 11, padding: 14 } },
        tooltip: { backgroundColor: "rgba(74,6,8,0.9)", titleColor: "#fff", bodyColor: "#e8e0d8", padding: 10, cornerRadius: 6 }
      },
      scales: {
        x: { grid: { display: false }, ticks: { maxRotation: 45, autoSkip: true } },
        y: { grid: gridOpts(), beginAtZero: true, ticks: { precision: 0 } }
      },
      elements: { point: { radius: 2, hoverRadius: 5 }, line: { tension: 0.32 } },
      maintainAspectRatio: false
    };
    if (extra) {
      if (extra.plugins) base.plugins = Object.assign({}, base.plugins, extra.plugins);
      Object.keys(extra).forEach(function (k) { if (k !== "plugins") base[k] = extra[k]; });
    }
    return base;
  }

  /* ============================ DATATABLES ============================ */

  function initDataTable(selector, options) {
    if (dtRegistry[selector]) {
      try { dtRegistry[selector].destroy(); } catch (e) { /* noop */ }
    }
    dtRegistry[selector] = $(selector).DataTable(Object.assign({ language: DT_LANG_ES }, options || {}));
    return dtRegistry[selector];
  }

  /* ===================== CONSTRUCCIÓN DE FILAS ===================== */

  function buildCaseRow(r, opts) {
    opts = opts || {};
    const cls = effectiveClass(r);
    const rowClass = cls === "Vencido" ? "row--vencido" : (cls === "Critico" ? "row--critico" : "");
    let html = '<tr class="' + rowClass + '">';
    html += '<td>' + esc(r["No. Caso"]) + '</td>';
    html += '<td>' + areaChipHTML(r["_area"]) + '</td>';
    if (opts.includeFechaRegistro) {
      html += '<td>' + esc(r["Fecha de registro"]) + '</td>';
    }
    html += '<td>' + esc(r["Estado"]) + '</td>';
    html += '<td>' + esc(r["Categoría"]) + '</td>';
    html += '<td>' + esc(r["Responsable"]) + '</td>';
    html += '<td>' + esc(r["Fecha Estimada de Solución"]) + '</td>';
    html += '<td data-order="' + (r["Tiempo transcurrido"] || 0) + '">' + (r["Tiempo transcurrido"] || 0).toFixed(1) + ' días</td>';
    html += '<td data-order="' + r["Progreso"] + '">' + progressCellHTML(r["Progreso"], cls) + '</td>';
    html += '<td>' + esc(r["Valor adicional"] || "") + '</td>';
    html += '</tr>';
    return html;
  }

  function buildModuleVencidoRow(r) {
    const cls = effectiveClass(r);
    const rowClass = cls === "Vencido" ? "row--vencido" : (cls === "Critico" ? "row--critico" : "");
    let html = '<tr class="' + rowClass + '">';
    html += '<td>' + esc(r["No. Caso"]) + '</td>';
    html += '<td>' + esc(r["Estado"]) + '</td>';
    html += '<td>' + esc(r["Categoría"]) + '</td>';
    html += '<td>' + esc(r["Responsable"]) + '</td>';
    html += '<td>' + esc(r["Fecha Estimada de Solución"]) + '</td>';
    html += '<td data-order="' + (r["Tiempo transcurrido"] || 0) + '">' + (r["Tiempo transcurrido"] || 0).toFixed(1) + ' días</td>';
    html += '<td data-order="' + r["Progreso"] + '">' + progressCellHTML(r["Progreso"], cls) + '</td>';
    html += '</tr>';
    return html;
  }

  function buildModuleAllRow(r) {
    const cls = effectiveClass(r);
    const rowClass = cls === "Vencido" ? "row--vencido" : (cls === "Critico" ? "row--critico" : "");
    let html = '<tr class="' + rowClass + '">';
    html += '<td>' + esc(r["No. Caso"]) + '</td>';
    html += '<td>' + esc(r["Fecha de registro"]) + '</td>';
    html += '<td>' + esc(r["Estado"]) + '</td>';
    html += '<td>' + esc(r["Categoría"]) + '</td>';
    html += '<td>' + esc(r["Servicio"]) + '</td>';
    html += '<td>' + esc(r["Responsable"]) + '</td>';
    html += '<td>' + esc(r["Grupo"]) + '</td>';
    html += '<td>' + esc(r["Urgencia"]) + '</td>';
    html += '<td data-order="' + r["Progreso"] + '">' + progressCellHTML(r["Progreso"], cls) + '</td>';
    if (r["_area"] === "CAE") html += '<td>' + esc(r["Valor adicional"] || "") + '</td>';
    html += '</tr>';
    return html;
  }

  /* ========================= CARGA DE DATOS ========================= */

  function loadDataSource(areaKey) {
    const url = CONFIG.dataSources[areaKey] + "?_=" + Date.now();
    return fetch(url, { cache: "no-store" })
      .then(function (resp) {
        if (!resp.ok) throw new Error("HTTP " + resp.status);
        return resp.json();
      })
      .then(function (json) {
        const records = (json && Array.isArray(json.records)) ? json.records : [];
        records.forEach(function (r) { r["_area"] = areaKey; });
        STATE.rawData[areaKey] = records;
        STATE.data[areaKey] = records; // se sobreescribe por applyGlobalFilters en renderAll
        STATE.errors[areaKey] = null;
      })
      .catch(function (err) {
        STATE.errors[areaKey] = "Origen no disponible (" + err.message + ")";
        if (!STATE.rawData[areaKey]) STATE.rawData[areaKey] = [];
        if (!STATE.data[areaKey])   STATE.data[areaKey]   = [];
      });
  }

  function loadAllData(isManual) {
    setSyncStatus("syncing");
    return Promise.all(AREAS.map(loadDataSource)).then(function () {
      STATE.lastUpdated = new Date();

      let hasError = false, allError = true;
      AREAS.forEach(function (a) {
        if (STATE.errors[a]) hasError = true; else allError = false;
      });
      setSyncStatus(allError ? "error" : (hasError ? "partial" : "ok"));

      renderErrorBanners();

      if (isManual !== true && STATE.firstLoadDone) {
        AREAS.forEach(function (a) {
          if (STATE.errors[a] && STATE.errors[a] !== STATE.prevErrors[a]) {
            showToast("No se pudo actualizar " + a + ". Mostrando últimos datos disponibles.");
          }
        });
      }
      STATE.prevErrors = Object.assign({}, STATE.errors);

      if (!STATE.firstLoadDone) {
        initEstadoFilter(); // pre-selecciona todos los estados en el primer arranque
      }
      populateGlobalFilterBar(); // reconstruye opciones con nuevos datos
      renderAll();
      STATE.firstLoadDone = true;
    });
  }

  /* ============================ UI: ESTADO ============================ */

  function setSyncStatus(state) {
    const pill = document.getElementById("syncPill");
    const icon = document.getElementById("syncIcon");
    const text = document.getElementById("syncText");
    if (!pill) return;
    pill.classList.remove("is-syncing", "is-error", "is-partial");
    if (state === "syncing") {
      pill.classList.add("is-syncing");
      icon.className = "bi bi-arrow-repeat";
      text.textContent = "Sincronizando…";
    } else if (state === "ok") {
      icon.className = "bi bi-check-circle";
      text.textContent = "Datos al día";
    } else if (state === "partial") {
      pill.classList.add("is-partial");
      icon.className = "bi bi-exclamation-circle";
      text.textContent = "Sincronización parcial";
    } else if (state === "error") {
      pill.classList.add("is-error");
      icon.className = "bi bi-x-circle";
      text.textContent = "Sin conexión a los orígenes";
    }
    if (state === "ok" || state === "partial" || state === "error") {
      updateLastUpdatedUI();
    }
  }

  function updateLastUpdatedUI() {
    const el = document.getElementById("lastUpdateValue");
    if (!el || !STATE.lastUpdated) return;
    const d = STATE.lastUpdated;
    const hh = String(d.getHours()).padStart(2, "0");
    const mm = String(d.getMinutes()).padStart(2, "0");
    const ss = String(d.getSeconds()).padStart(2, "0");
    el.textContent = hh + ":" + mm + ":" + ss;
  }

  function renderErrorBanners() {
    const stack = document.getElementById("errorBannerStack");
    if (!stack) return;
    let html = "";
    AREAS.forEach(function (a) {
      if (STATE.errors[a]) {
        html += '<div class="error-banner"><i class="bi bi-exclamation-triangle-fill"></i> ' +
          '<strong>' + esc(a) + ':</strong> ' + esc(STATE.errors[a]) + ' — mostrando los últimos datos disponibles.</div>';
      }
    });
    stack.innerHTML = html;
  }

  function showToast(message) {
    const stack = document.getElementById("toastStack");
    if (!stack) return;
    const item = document.createElement("div");
    item.className = "toast-item";
    item.innerHTML = '<i class="bi bi-exclamation-circle"></i> ' + esc(message);
    stack.appendChild(item);
    setTimeout(function () {
      item.classList.add("is-leaving");
      setTimeout(function () { item.remove(); }, 400);
    }, 6000);
  }

  /* ============================ NAVEGACIÓN ============================ */

  function switchView(key) {
    document.querySelectorAll(".nav-link[data-view]").forEach(function (btn) {
      btn.classList.toggle("active", btn.getAttribute("data-view") === key);
    });
    document.querySelectorAll(".view[data-view]").forEach(function (sec) {
      sec.classList.toggle("active", sec.getAttribute("data-view") === key);
    });
    const title = VIEW_TITLES[key] || key;
    const titleEl = document.getElementById("viewTitle");
    const crumbEl = document.getElementById("breadcrumbCurrent");
    if (titleEl) titleEl.textContent = title;
    if (crumbEl) crumbEl.textContent = title;

    document.body.classList.remove("sidebar-is-open");
    const sidebar = document.getElementById("sidebar");
    const overlay = document.getElementById("sidebarOverlay");
    if (sidebar) sidebar.classList.remove("is-open");
    if (overlay) overlay.classList.remove("is-open");

    // Ocultar la barra de filtros globales en la vista de Responsables
    const gfb = document.getElementById("globalFilterBar");
    if (gfb) gfb.style.display = (key === "responsables") ? "none" : "";

    setTimeout(function () {
      document.querySelectorAll(".chart-scroll-inner[data-pts]").forEach(function (inner) {
        const n = parseInt(inner.dataset.pts) || 0;
        const pW = inner.parentElement ? inner.parentElement.clientWidth : 0;
        const calcW = Math.max(n * 38, 300);
        inner.style.width = (pW > 0 ? Math.max(calcW, pW) : calcW) + "px";
      });
      Object.keys(chartRegistry).forEach(function (id) {
        if (chartRegistry[id]) {
          try { chartRegistry[id].resize(); } catch (e) { /* noop */ }
        }
      });
    }, 60);
  }

  /* ===================== PLANTILLA DE MÓDULO (HTML) ===================== */

  function buildModuleHTML(areaKey, prefix) {
    return (
      '<div class="kpi-grid" id="kpi-' + prefix + '-grid"></div>' +
      '<div class="module-charts">' +
        moduleChartPanel(prefix, "estado",       "bi-pie-chart",         "Casos por estado") +
        moduleChartPanel(prefix, "categoria",    "bi-tags",              "Casos por categoría") +
        moduleChartPanel(prefix, "prioridad",    "bi-flag",              "Casos por prioridad") +
        moduleChartPanel(prefix, "responsable",  "bi-people",            "Casos por responsable") +
        moduleChartPanel(prefix, "grupo",        "bi-building",          "Casos por grupo") +
        moduleChartPanel(prefix, "vencidos-resp","bi-person-exclamation","Vencidos por responsable") +
        moduleChartPanel(prefix, "tendencia",    "bi-graph-up-arrow",    "Tendencia de creación (9 semanas)") +
        moduleChartPanel(prefix, "servicio",     "bi-diagram-3",         "Distribución por servicio") +
      '</div>' +
      '<div class="panel">' +
        '<div class="panel-header">' +
          '<h2><i class="bi bi-table"></i> Todos los casos · ' + esc(areaKey) + '</h2>' +
          '<span class="panel-sub">Resultados según los filtros globales aplicados</span>' +
        '</div>' +
        '<div class="table-responsive">' +
          '<table class="table data-table" id="table-' + prefix + '-all" style="width:100%">' +
            '<thead><tr>' +
              '<th>No. Caso</th><th>Fecha Reg.</th><th>Estado</th><th>Categoría</th>' +
              '<th>Servicio</th><th>Responsable</th><th>Grupo</th><th>Urgencia</th><th>Progreso</th>' +
              (areaKey === "CAE" ? '<th>Inf. Adicional</th>' : '') +
            '</tr></thead><tbody></tbody>' +
          '</table>' +
        '</div>' +
      '</div>'
    );
  }

  function moduleChartPanel(prefix, key, icon, title) {
    const periodSlot = key === "tendencia"
      ? '<div class="tend-period-btns" id="tend-' + prefix + '-period"></div>'
      : "";
    const chartInner = key === "tendencia"
      ? '<div class="chart-wrap chart-wrap--hscroll"><div class="chart-scroll-inner" id="tend-' + prefix + '-inner"><canvas id="chart-' + prefix + '-' + key + '" class="chart-skeleton"></canvas></div></div>'
      : '<div class="chart-wrap"><canvas id="chart-' + prefix + '-' + key + '" class="chart-skeleton"></canvas></div>';
    return (
      '<div class="panel">' +
        '<div class="panel-header"><h2><i class="bi ' + icon + '"></i> ' + title + '</h2>' + periodSlot + '</div>' +
        chartInner +
      '</div>'
    );
  }

  function injectModuleTemplates() {
    document.querySelectorAll(".view[data-area]").forEach(function (sec) {
      const area = sec.getAttribute("data-area");
      const prefix = sec.getAttribute("data-prefix");
      sec.innerHTML = buildModuleHTML(area, prefix);
    });
  }

  /* ============================ RENDERIZADO ============================ */

  function renderAll() {
    applyGlobalFilters();
    computeAllStats();
    updateLastUpdatedUI();
    updateSidebarBadges();
    updateFilterBadges();
    renderExecutive();
    renderAttention();
    AREAS.forEach(renderModule);
    renderSolucionados();
    renderResponsables();
  }

  function updateSidebarBadges() {
    AREAS.forEach(function (a) {
      const el = document.getElementById("navBadge" + AREA_SHORT[a].replace(/\s/g, ""));
      if (el) el.textContent = STATE.stats[a].vencidos;
    });
    const atencionEl = document.getElementById("navBadgeAtencion");
    if (atencionEl) atencionEl.textContent = STATE.statsCombined.vencidos + STATE.statsCombined.criticos;

    const solEl = document.getElementById("navBadgeSolucionados");
    if (solEl) {
      let total = 0;
      AREAS.forEach(function (a) { total += (STATE.rawData[a] || []).filter(function (r) { return r["Estado"] === "Solucionado"; }).length; });
      solEl.textContent = total;
    }
  }

  function renderExecutive() {
    const s = STATE.statsCombined;
    const grid = document.getElementById("kpiExecGrid");
    if (grid) {
      grid.innerHTML =
        kpi("Total de casos", s.total, "info", "bi-collection", "PQRS · Servicios Financieros · CAE") +
        kpi("Vencidos", s.vencidos, "vencido", "bi-x-octagon", pct(s.vencidos, s.total) + "% del total") +
        kpi("Críticos", s.criticos, "critico", "bi-exclamation-triangle", pct(s.criticos, s.total) + "% del total") +
        kpi("En riesgo", s.riesgo, "riesgo", "bi-shield-exclamation", pct(s.riesgo, s.total) + "% del total") +
        kpi("SLA cumplido", s.slaPct + "%", "sla", "bi-patch-check", "Progreso ≤ 100%") +
        kpi("Tiempo promedio", s.avgTiempo + " días", "normal", "bi-clock-history", "transcurrido por caso") +
        kpi("Responsables con vencidos", Object.keys(s.vencidosPorResponsable).length, "info", "bi-person-badge", "personas con casos vencidos");
    }
    renderExecCharts();
    renderComparativo();
  }

  function renderExecCharts() {
    // Period buttons
    const periodContainer = document.getElementById("execTendPeriod");
    if (periodContainer) {
      periodContainer.innerHTML = buildPeriodBtnsHTML();
      wireTendencyBtns(periodContainer);
    }

    // Tendency chart — use combined records for shared labels, per-area for series
    const combinedBuckets = getTendenciaCounts(STATE.combinedRecords, TENDENCY_PERIOD);
    const labels = combinedBuckets.map(function (b) { return b.label; });
    const datasets = AREAS.map(function (area) {
      const areaBuckets = getTendenciaCounts(STATE.data[area] || [], TENDENCY_PERIOD);
      const labelToCount = {};
      areaBuckets.forEach(function (b) { labelToCount[b.label] = b.count; });
      return {
        label: AREA_SHORT[area],
        data: labels.map(function (l) { return labelToCount[l] || 0; }),
        borderColor: AREA_COLORS[area],
        backgroundColor: AREA_COLORS[area],
        fill: false
      };
    });
    const execTendInner = document.getElementById("execTendInner");
    if (execTendInner) {
      execTendInner.dataset.pts = labels.length;
      const pW = execTendInner.parentElement ? execTendInner.parentElement.clientWidth : 0;
      const calcW = Math.max(labels.length * 38, 300);
      execTendInner.style.width = (pW > 0 ? Math.max(calcW, pW) : calcW) + "px";
    }
    renderChart("chartExecTendencia", "line", { labels: labels, datasets: datasets }, lineOpts());

    const s = STATE.statsCombined;
    const counts = { Normal: s.normal, Riesgo: s.riesgo, Critico: s.criticos, Vencido: s.vencidos };
    renderChart("chartExecClasificacion", "doughnut", toChartDataDoughnut(counts, STATUS_LABELS, STATUS_COLORS), doughnutOpts());
  }

  function renderAttention() {
    const s = STATE.statsCombined;
    const grid = document.getElementById("kpiAttentionGrid");
    if (grid) {
      grid.innerHTML =
        kpi("Vencidos", s.vencidos, "vencido", "bi-x-octagon", "requieren acción inmediata") +
        kpi("Críticos", s.criticos, "critico", "bi-exclamation-triangle", "por vencer en horas") +
        kpi("Total en atención", s.vencidos + s.criticos, "atencion", "bi-megaphone", "vencidos + críticos") +
        kpi("SLA incumplido", (100 - s.slaPct).toFixed(1) + "%", "vencido", "bi-graph-down", "del total de casos");
    }

    const respTop = topEntry(s.vencidosPorResponsable);
    const catTop = topEntry(s.vencidosPorCategoria);

    let areaTop = { key: "—", count: 0 };
    AREAS.forEach(function (a) {
      if (STATE.stats[a].vencidos > areaTop.count) areaTop = { key: a, count: STATE.stats[a].vencidos };
    });

    setSpotlight("spotlightResponsable", respTop);
    setSpotlight("spotlightCategoria", catTop);
    setSpotlight("spotlightArea", areaTop);

    const atencionCases = STATE.combinedRecords
      .filter(function (r) { const cls = effectiveClass(r); return cls === "Vencido" || cls === "Critico"; })
      .sort(function (a, b) { return b["Progreso"] - a["Progreso"]; });

    const selAt = "#tableAtencion";
    if (dtRegistry[selAt]) { try { dtRegistry[selAt].destroy(); } catch (e) {} delete dtRegistry[selAt]; }

    const tbodyAt = document.querySelector(selAt + " tbody");
    if (tbodyAt) {
      tbodyAt.innerHTML = atencionCases.map(function (r) { return buildCaseRow(r, { includeFechaRegistro: true }); }).join("");
    }
    dtRegistry[selAt] = $(selAt).DataTable(Object.assign({ language: DT_LANG_ES }, {
      paging: true, pageLength: 15, order: [],
      dom: "frtipB",
      buttons: [
        { extend: "excelHtml5", text: '<i class="bi bi-file-earmark-excel"></i> Excel', className: "dt-button" },
        { extend: "csvHtml5",   text: '<i class="bi bi-filetype-csv"></i> CSV',         className: "dt-button" }
      ]
    }));
  }

  function setSpotlight(id, entry) {
    const el = document.getElementById(id);
    if (!el) return;
    if (!entry || entry.count === 0) {
      el.innerHTML = '<span class="spotlight-empty">Sin casos vencidos</span>';
      return;
    }
    el.innerHTML = esc(entry.key) + '<span class="spotlight-count">' + entry.count + ' vencido' + (entry.count === 1 ? "" : "s") + '</span>';
  }

  function renderComparativo() {
    const tbody = document.querySelector("#tableComparativo tbody");
    if (tbody) {
      let html = "";
      AREAS.forEach(function (a) {
        const s = STATE.stats[a];
        html += "<tr>";
        html += "<td>" + areaChipHTML(a) + " " + esc(a) + "</td>";
        html += '<td data-order="' + s.total + '">' + s.total + "</td>";
        html += '<td data-order="' + s.vencidos + '">' + s.vencidos + "</td>";
        html += '<td data-order="' + s.criticos + '">' + s.criticos + "</td>";
        html += '<td data-order="' + s.riesgo + '">' + s.riesgo + "</td>";
        html += '<td data-order="' + s.slaPct + '">' + s.slaPct + "%</td>";
        html += '<td data-order="' + s.avgTiempo + '">' + s.avgTiempo + "</td>";
        html += "</tr>";
      });
      tbody.innerHTML = html;
    }
    initDataTable("#tableComparativo", { paging: false, searching: false, info: false, order: [] });

    const datasets = ["Normal", "Riesgo", "Critico", "Vencido"].map(function (cls) {
      return {
        label: STATUS_LABELS[cls],
        data: AREAS.map(function (a) {
          const s = STATE.stats[a];
          return cls === "Normal" ? s.normal : (cls === "Riesgo" ? s.riesgo : (cls === "Critico" ? s.criticos : s.vencidos));
        }),
        backgroundColor: STATUS_COLORS[cls]
      };
    });
    renderChart("chartComparativoStack", "bar", { labels: AREAS.map(function (a) { return AREA_SHORT[a]; }), datasets: datasets }, stackedBarOpts());
  }

  function renderRadar() {
    const sorted = STATE.combinedRecords
      .filter(function (r) { return r["Estado"] !== "Solucionado"; })
      .slice()
      .sort(function (a, b) {
        if (b["Progreso"] !== a["Progreso"]) return b["Progreso"] - a["Progreso"];
        const da = a["Fecha Estimada de Solución"] || "";
        const db = b["Fecha Estimada de Solución"] || "";
        return da < db ? -1 : (da > db ? 1 : 0);
      })
    const tbody = document.querySelector("#tableRadar tbody");
    if (tbody) {
      tbody.innerHTML = sorted.map(function (r) { return buildCaseRow(r, { includeFechaRegistro: false }); }).join("");
    }
    initDataTable("#tableRadar", { paging: true, pageLength: 10, order: [], lengthChange: false });
  }

  function renderModule(areaKey) {
    const prefix = AREA_PREFIX[areaKey];
    const records = STATE.data[areaKey] || [];
    const s = STATE.stats[areaKey];

    const grid = document.getElementById("kpi-" + prefix + "-grid");
    if (grid) {
      grid.innerHTML =
        kpi("Total de casos", s.total, "info", "bi-collection", areaKey) +
        kpi("Vencidos", s.vencidos, "vencido", "bi-x-octagon", pct(s.vencidos, s.total) + "% del total") +
        kpi("Críticos", s.criticos, "critico", "bi-exclamation-triangle", pct(s.criticos, s.total) + "% del total") +
        kpi("En riesgo", s.riesgo, "riesgo", "bi-shield-exclamation", pct(s.riesgo, s.total) + "% del total") +
        kpi("SLA cumplido", s.slaPct + "%", "sla", "bi-patch-check", "Progreso ≤ 100%") +
        kpi("Tiempo promedio", s.avgTiempo + " días", "normal", "bi-clock-history", "transcurrido por caso");
    }

    const statusCounts = { Normal: s.normal, Riesgo: s.riesgo, Critico: s.criticos, Vencido: s.vencidos };
    renderChart("chart-" + prefix + "-estado",       "doughnut", toChartDataDoughnut(statusCounts, STATUS_LABELS, STATUS_COLORS), doughnutOpts());
    renderChart("chart-" + prefix + "-categoria",    "bar", toChartDataBar(countBy(records, "Categoría"),   "#8C0F13", 10), horizontalBarOpts());
    renderChart("chart-" + prefix + "-prioridad",    "bar", toChartDataBarMulti(countBy(records, "Prioridad"),  8), barOpts());
    renderChart("chart-" + prefix + "-responsable",  "bar", toChartDataBar(countBy(records, "Responsable"),  "#6B5E54", 8), horizontalBarOpts());
    renderChart("chart-" + prefix + "-grupo",        "bar", toChartDataBarMulti(countBy(records, "Grupo"),      8), barOpts());
    renderChart("chart-" + prefix + "-vencidos-resp","bar", toChartDataBar(s.vencidosPorResponsable,         "#4A0608", 8), horizontalBarOpts());

    const tendContainer = document.getElementById("tend-" + prefix + "-period");
    if (tendContainer) {
      tendContainer.innerHTML = buildPeriodBtnsHTML();
      wireTendencyBtns(tendContainer);
    }
    const buckets = getTendenciaCounts(records, TENDENCY_PERIOD);
    const tendInner = document.getElementById("tend-" + prefix + "-inner");
    if (tendInner) {
      tendInner.dataset.pts = buckets.length;
      const pW = tendInner.parentElement ? tendInner.parentElement.clientWidth : 0;
      const calcW = Math.max(buckets.length * 38, 300);
      tendInner.style.width = (pW > 0 ? Math.max(calcW, pW) : calcW) + "px";
    }
    renderChart("chart-" + prefix + "-tendencia", "line", {
      labels: buckets.map(function (b) { return b.label; }),
      datasets: [{ label: areaKey, data: buckets.map(function (b) { return b.count; }), borderColor: AREA_COLORS[areaKey], backgroundColor: AREA_COLORS[areaKey] + "22", fill: true }]
    }, lineOpts({ plugins: { legend: { display: false } } }));

    renderChart("chart-" + prefix + "-servicio", "bar", toChartDataBarMulti(countBy(records, "Servicio"), 8), barOpts());

    // Tabla "Todos los casos": destroy antes de setear tbody para evitar caché de DataTable
    const selAll = "#table-" + prefix + "-all";
    if (dtRegistry[selAll]) { try { dtRegistry[selAll].destroy(); } catch (e) {} delete dtRegistry[selAll]; }
    const tbodyAll = document.querySelector(selAll + " tbody");
    if (tbodyAll) tbodyAll.innerHTML = records.map(buildModuleAllRow).join("");
    dtRegistry[selAll] = $(selAll).DataTable(Object.assign({ language: DT_LANG_ES }, {
      paging: true, pageLength: 15, order: [[8, "desc"]],
      dom: "frtipB",
      buttons: [
        { extend: "excelHtml5", text: '<i class="bi bi-file-earmark-excel"></i> Excel', className: "dt-button" },
        { extend: "csvHtml5",   text: '<i class="bi bi-filetype-csv"></i> CSV',         className: "dt-button" }
      ]
    }));
  }

  function renderProximos() {
    const filtered = STATE.combinedRecords
      .filter(function (r) { return r["Progreso"] >= 90 && r["Estado"] !== "Solucionado"; })
      .sort(function (a, b) { return b["Progreso"] - a["Progreso"]; });

    const tbody = document.querySelector("#tableProximos tbody");
    if (tbody) {
      tbody.innerHTML = filtered.map(function (r) { return buildCaseRow(r, { includeFechaRegistro: true }); }).join("");
    }

    initDataTable("#tableProximos", {
      paging: true,
      pageLength: 10,
      order: [],
      dom: "frtipB",
      buttons: [
        { extend: "excelHtml5", text: '<i class="bi bi-file-earmark-excel"></i> Excel', className: "dt-button" },
        { extend: "csvHtml5",   text: '<i class="bi bi-filetype-csv"></i> CSV',         className: "dt-button" }
      ]
    });
  }

  /* ====================== SOLUCIONADOS ====================== */

  function renderSolucionados() {
    const sol = [];
    AREAS.forEach(function (a) {
      (STATE.rawData[a] || []).forEach(function (r) {
        if (r["Estado"] !== "Solucionado") return;
        // Aplica todos los filtros globales EXCEPTO Estado (la pestaña Solucionados es independiente de ese filtro)
        if (GLOBAL_FILTER.area.length        && GLOBAL_FILTER.area.indexOf(r["_area"])            === -1) return;
        if (GLOBAL_FILTER.categoria.length   && GLOBAL_FILTER.categoria.indexOf(r["Categoría"])   === -1) return;
        if (GLOBAL_FILTER.servicio.length    && GLOBAL_FILTER.servicio.indexOf(r["Servicio"])     === -1) return;
        if (GLOBAL_FILTER.responsable.length && GLOBAL_FILTER.responsable.indexOf(r["Responsable"]) === -1) return;
        if (GLOBAL_FILTER.grupo.length       && GLOBAL_FILTER.grupo.indexOf(r["Grupo"])           === -1) return;
        if (GLOBAL_FILTER.urgencia.length    && GLOBAL_FILTER.urgencia.indexOf(r["Urgencia"])     === -1) return;
        if (GLOBAL_FILTER.fechaDesde && (r["Fecha de registro"] || "") < GLOBAL_FILTER.fechaDesde) return;
        if (GLOBAL_FILTER.fechaHasta && (r["Fecha de registro"] || "") > GLOBAL_FILTER.fechaHasta) return;
        sol.push(r);
      });
    });

    // Total de casos (mismos filtros activos excepto Estado) para calcular % de solución
    let totalCasos = 0;
    AREAS.forEach(function (a) {
      (STATE.rawData[a] || []).forEach(function (r) {
        if (GLOBAL_FILTER.area.length        && GLOBAL_FILTER.area.indexOf(r["_area"])            === -1) return;
        if (GLOBAL_FILTER.categoria.length   && GLOBAL_FILTER.categoria.indexOf(r["Categoría"])   === -1) return;
        if (GLOBAL_FILTER.servicio.length    && GLOBAL_FILTER.servicio.indexOf(r["Servicio"])     === -1) return;
        if (GLOBAL_FILTER.responsable.length && GLOBAL_FILTER.responsable.indexOf(r["Responsable"]) === -1) return;
        if (GLOBAL_FILTER.grupo.length       && GLOBAL_FILTER.grupo.indexOf(r["Grupo"])           === -1) return;
        if (GLOBAL_FILTER.urgencia.length    && GLOBAL_FILTER.urgencia.indexOf(r["Urgencia"])     === -1) return;
        if (GLOBAL_FILTER.fechaDesde && (r["Fecha de registro"] || "") < GLOBAL_FILTER.fechaDesde) return;
        if (GLOBAL_FILTER.fechaHasta && (r["Fecha de registro"] || "") > GLOBAL_FILTER.fechaHasta) return;
        totalCasos++;
      });
    });
    const pctSolucion = pct(sol.length, totalCasos);

    // Clasificar cada solucionado por su Progreso real
    let vencidos = 0, criticos = 0, riesgo = 0, aTiempo = 0;
    sol.forEach(function (r) {
      const cls = classify(r["Progreso"]);
      if (cls === "Vencido")       vencidos++;
      else if (cls === "Critico")  criticos++;
      else if (cls === "Riesgo")   riesgo++;
      else                         aTiempo++;
    });

    const grid = document.getElementById("kpiSolucionadosGrid");
    if (grid) {
      grid.innerHTML =
        kpi("% de solución", pctSolucion + "%", "sla", "bi-graph-up", sol.length + " de " + totalCasos + " casos totales") +
        kpi("Total solucionados", sol.length, "sla", "bi-check2-circle", "todos los procesos") +
        kpi("A tiempo", aTiempo, "normal", "bi-patch-check", "Progreso ≤ 70% al resolver") +
        kpi("Resueltos en riesgo", riesgo, "riesgo", "bi-shield-exclamation", "Progreso 70–90% al resolver") +
        kpi("Resueltos críticos", criticos, "critico", "bi-exclamation-triangle", "Progreso 90–100% al resolver") +
        kpi("Resueltos vencidos", vencidos, "vencido", "bi-x-octagon", "Progreso >100% al resolver");
    }

    const sorted = sol.slice().sort(function (a, b) { return b["Progreso"] - a["Progreso"]; });
    const selSol = "#tableSolucionados";
    if (dtRegistry[selSol]) { try { dtRegistry[selSol].destroy(); } catch (e) {} delete dtRegistry[selSol]; }
    const tbody = document.querySelector(selSol + " tbody");
    if (tbody) {
      tbody.innerHTML = sorted.map(function (r) {
        const cls = classify(r["Progreso"]);
        const rowClass = cls === "Vencido" ? "row--vencido" : (cls === "Critico" ? "row--critico" : "");
        let html = '<tr class="' + rowClass + '">';
        html += '<td>' + esc(r["No. Caso"]) + '</td>';
        html += '<td>' + areaChipHTML(r["_area"]) + '</td>';
        html += '<td>' + esc(r["Fecha de registro"]) + '</td>';
        html += '<td>' + esc(r["Estado"]) + '</td>';
        html += '<td>' + esc(r["Categoría"]) + '</td>';
        html += '<td>' + esc(r["Responsable"]) + '</td>';
        html += '<td>' + esc(r["Fecha Estimada de Solución"]) + '</td>';
        html += '<td data-order="' + (r["Tiempo transcurrido"] || 0) + '">' + (r["Tiempo transcurrido"] || 0).toFixed(1) + ' días</td>';
        html += '<td data-order="' + r["Progreso"] + '">' + progressCellHTML(r["Progreso"], cls) + '</td>';
        html += '<td>' + esc(r["Valor adicional"] || "") + '</td>';
        html += '</tr>';
        return html;
      }).join("");
    }
    dtRegistry[selSol] = $(selSol).DataTable(Object.assign({ language: DT_LANG_ES }, {
      paging: true, pageLength: 15, order: [[8, "desc"]],
      dom: "frtipB",
      buttons: [
        { extend: "excelHtml5", text: '<i class="bi bi-file-earmark-excel"></i> Excel', className: "dt-button" },
        { extend: "csvHtml5",   text: '<i class="bi bi-filetype-csv"></i> CSV',         className: "dt-button" }
      ]
    }));
  }

  /* ====================== GESTIÓN DE RESPONSABLES ====================== */

  function computeResponsableStats() {
    const allRaw = [];
    AREAS.forEach(function (a) {
      (STATE.rawData[a] || []).forEach(function (r) {
        if (RESP_SECTION_FILTER.responsable.length && RESP_SECTION_FILTER.responsable.indexOf(r["Responsable"]) === -1) return;
        if (RESP_SECTION_FILTER.grupo.length       && RESP_SECTION_FILTER.grupo.indexOf(r["Grupo"])             === -1) return;
        allRaw.push(r);
      });
    });

    const fechas = allRaw.map(function (r) { return r["Fecha de registro"]; }).filter(Boolean).sort();
    const globalDays = fechas.length > 1
      ? Math.max(1, (new Date(fechas[fechas.length - 1] + "T00:00:00") - new Date(fechas[0] + "T00:00:00")) / 86400000) + 1
      : 1;

    const byResp = {};
    allRaw.forEach(function (r) {
      const resp = r["Responsable"] || "Sin asignar";
      if (!byResp[resp]) {
        byResp[resp] = {
          nombre: resp,
          totalCasos: 0,
          abiertos: 0,
          solucionados: 0,
          solucionadosSLA: 0,
          vencidosActivos: 0,
          criticosActivos: 0,
          riesgoActivos: 0,
          normalActivos: 0,
          tiemposSolucionados: [],
          tiemposAbiertos: [],
          categorias: {},
          areas: {}
        };
      }
      const d = byResp[resp];
      d.totalCasos++;
      d.areas[r["_area"]] = (d.areas[r["_area"]] || 0) + 1;
      const cat = r["Categoría"] || "Sin categoría";
      d.categorias[cat] = (d.categorias[cat] || 0) + 1;

      if (r["Estado"] === "Solucionado") {
        d.solucionados++;
        if (r["Progreso"] <= 100) d.solucionadosSLA++;
        if (r["Tiempo transcurrido"] != null) d.tiemposSolucionados.push(r["Tiempo transcurrido"]);
      } else {
        d.abiertos++;
        if (r["Tiempo transcurrido"] != null) d.tiemposAbiertos.push(r["Tiempo transcurrido"]);
        const cls = classify(r["Progreso"]);
        if (cls === "Vencido")      d.vencidosActivos++;
        else if (cls === "Critico") d.criticosActivos++;
        else if (cls === "Riesgo")  d.riesgoActivos++;
        else                        d.normalActivos++;
      }
    });

    Object.keys(byResp).forEach(function (k) {
      const d = byResp[k];
      d.tasaResolucion = d.totalCasos > 0 ? +(d.solucionados / d.totalCasos * 100).toFixed(1) : 0;
      d.slaCumplido    = d.solucionados > 0 ? +(d.solucionadosSLA / d.solucionados * 100).toFixed(1) : null;
      d.avgTiempoSolucionados = d.tiemposSolucionados.length > 0
        ? +(d.tiemposSolucionados.reduce(function (s, v) { return s + v; }, 0) / d.tiemposSolucionados.length).toFixed(1)
        : null;
      d.avgTiempoAbiertos = d.tiemposAbiertos.length > 0
        ? +(d.tiemposAbiertos.reduce(function (s, v) { return s + v; }, 0) / d.tiemposAbiertos.length).toFixed(1)
        : null;
      d.casosXDia  = +(d.totalCasos  / globalDays).toFixed(3);
      d.areasList  = Object.keys(d.areas).join(", ");
    });

    return { byResp: byResp, globalDays: Math.round(globalDays) };
  }

  function populateRespSectionFilters() {
    const bar = document.getElementById("respSectionFilter");
    if (!bar) return;
    const respSet = new Set();
    const grupoSet = new Set();
    AREAS.forEach(function (a) {
      (STATE.rawData[a] || []).forEach(function (r) {
        if (r["Responsable"]) respSet.add(r["Responsable"]);
        if (r["Grupo"])       grupoSet.add(r["Grupo"]);
      });
    });
    const respNames = Array.from(respSet).sort();
    const grupos    = Array.from(grupoSet).sort();

    bar.innerHTML =
      '<div class="gfb-inner">' +
        '<span class="gfb-title"><i class="bi bi-funnel-fill"></i> Filtros de sección</span>' +
        '<div class="gfb-drops" id="respSectionDrops">' +
          buildMsDropHTML("responsable", "Responsable", "bi-person",   respNames, RESP_SECTION_FILTER) +
          buildMsDropHTML("grupo",       "Grupo",        "bi-building", grupos,    RESP_SECTION_FILTER) +
        '</div>' +
        '<button class="gfb-clear" id="respFilterClear" title="Limpiar filtros de sección">' +
          '<i class="bi bi-x-circle"></i> Limpiar' +
        '</button>' +
      '</div>';

    wireRespFilterBar();
  }

  function wireRespFilterBar() {
    const bar = document.getElementById("respSectionDrops");
    if (!bar) return;

    bar.querySelectorAll(".ms-drop").forEach(function (drop) {
      const key    = drop.getAttribute("data-key");
      const toggle = drop.querySelector(".ms-toggle");
      const panel  = drop.querySelector(".ms-panel");
      const srch   = drop.querySelector(".ms-search");
      const allCb  = drop.querySelector(".ms-cb-all");
      const badge  = drop.querySelector(".ms-badge");

      toggle.addEventListener("click", function (e) {
        e.stopPropagation();
        const isOpen = !panel.hidden;
        closeAllDropdowns();
        if (!isOpen) {
          panel.hidden = false;
          drop.classList.add("is-open");
          if (srch) { srch.value = ""; filterDropOptions(drop, ""); srch.focus(); }
        }
      });
      panel.addEventListener("click", function (e) { e.stopPropagation(); });
      if (srch) {
        srch.addEventListener("input", function () { filterDropOptions(drop, this.value); });
        srch.addEventListener("click", function (e) { e.stopPropagation(); });
      }
      if (allCb) {
        allCb.addEventListener("change", function () {
          if (this.checked) {
            drop.querySelectorAll(".ms-cb").forEach(function (cb) { cb.checked = false; });
            RESP_SECTION_FILTER[key] = [];
            badge.style.display = "none";
            badge.textContent = "0";
            renderResponsablesContent();
          } else {
            this.checked = true;
          }
        });
      }
      drop.querySelectorAll(".ms-cb").forEach(function (cb) {
        cb.addEventListener("change", function () {
          const vals = [];
          drop.querySelectorAll(".ms-cb:checked").forEach(function (c) { vals.push(c.value); });
          RESP_SECTION_FILTER[key] = vals;
          if (allCb) allCb.checked = vals.length === 0;
          badge.textContent = vals.length;
          badge.style.display = vals.length > 0 ? "" : "none";
          renderResponsablesContent();
        });
      });
    });

    const clearBtn = document.getElementById("respFilterClear");
    if (clearBtn) {
      clearBtn.addEventListener("click", function () {
        RESP_SECTION_FILTER.responsable = [];
        RESP_SECTION_FILTER.grupo = [];
        populateRespSectionFilters(); // rebuild with empty state
        renderResponsables();
      });
    }
  }

  function renderResponsables() {
    populateRespSectionFilters();
    renderResponsablesContent();
  }

  function renderResponsablesContent() {
    const stats = computeResponsableStats();
    const byResp = stats.byResp;
    const globalDays = stats.globalDays;
    const responsables = Object.values(byResp).sort(function (a, b) { return b.totalCasos - a.totalCasos; });

    /* ---- KPI generales ---- */
    const topCarga = responsables.reduce(function (m, r) { return r.abiertos > m.abiertos ? r : m; }, { nombre: "—", abiertos: 0 });
    const topVenc  = responsables.reduce(function (m, r) { return r.vencidosActivos > m.vencidosActivos ? r : m; }, { nombre: "—", vencidosActivos: 0 });
    const conTiempo = responsables.filter(function (r) { return r.avgTiempoSolucionados !== null; });
    const avgGlobal = conTiempo.length > 0
      ? (conTiempo.reduce(function (s, r) { return s + r.avgTiempoSolucionados; }, 0) / conTiempo.length).toFixed(1)
      : null;

    const kpiGrid = document.getElementById("kpiResponsablesGrid");
    if (kpiGrid) {
      const n2w = function (n) { return n.split(" ").slice(0, 2).join(" "); };
      kpiGrid.innerHTML =
        kpi("Responsables en vista", responsables.length, "info",    "bi-people",             "según los filtros de sección") +
        kpi("Mayor carga activa",    n2w(topCarga.nombre), "info",    "bi-person-badge",       topCarga.abiertos + " casos abiertos") +
        kpi("Más vencidos activos",  n2w(topVenc.nombre),  "vencido", "bi-person-exclamation", topVenc.vencidosActivos + " casos vencidos") +
        kpi("Período analizado",     globalDays + " días", "sla",     "bi-calendar-range",     "rango de fechas en los datos") +
        kpi("Días prom. resolución", avgGlobal !== null ? avgGlobal + " días" : "—", "normal", "bi-clock-history", "promedio del grupo seleccionado");
    }

    /* ---- Tabla resumen (8 columnas) ---- */
    const selSum = "#tableRespResumen";
    if (dtRegistry[selSum]) { try { dtRegistry[selSum].destroy(); } catch (e) { /* noop */ } delete dtRegistry[selSum]; }
    const tbodySum = document.querySelector(selSum + " tbody");
    if (tbodySum) {
      tbodySum.innerHTML = responsables.map(function (r) {
        const rowCls   = r.vencidosActivos > 0 ? "row--vencido" : (r.criticosActivos > 0 ? "row--critico" : "");
        const slaStr   = r.slaCumplido !== null ? r.slaCumplido + "%" : "—";
        const tiempoStr = r.avgTiempoSolucionados !== null ? r.avgTiempoSolucionados + " d" : "—";
        const vBadge   = r.vencidosActivos > 0 ? '<span class="resp-badge resp-badge--vencido">' + r.vencidosActivos + '</span>' : "0";
        const cBadge   = r.criticosActivos > 0 ? '<span class="resp-badge resp-badge--critico">' + r.criticosActivos + '</span>' : "0";
        return (
          '<tr class="resp-row ' + rowCls + '" data-resp="' + esc(r.nombre) + '" title="Clic para ver el detalle de ' + esc(r.nombre) + '">' +
          '<td><strong>' + esc(r.nombre) + '</strong></td>' +
          '<td data-order="' + r.totalCasos + '">' + r.totalCasos + '</td>' +
          '<td data-order="' + r.abiertos + '">' + r.abiertos + '</td>' +
          '<td data-order="' + r.vencidosActivos + '">' + vBadge + '</td>' +
          '<td data-order="' + r.criticosActivos + '">' + cBadge + '</td>' +
          '<td data-order="' + r.tasaResolucion + '">' + r.tasaResolucion + '%</td>' +
          '<td data-order="' + (r.slaCumplido !== null ? r.slaCumplido : -1) + '">' + slaStr + '</td>' +
          '<td data-order="' + (r.avgTiempoSolucionados !== null ? r.avgTiempoSolucionados : 99999) + '">' + tiempoStr + '</td>' +
          '</tr>'
        );
      }).join("");
    }

    dtRegistry[selSum] = $(selSum).DataTable(Object.assign({ language: DT_LANG_ES }, {
      paging: true, pageLength: 5, lengthChange: false, order: [[2, "desc"]],
      autoWidth: false,
      dom: "frtipB",
      buttons: [
        { extend: "excelHtml5", text: '<i class="bi bi-file-earmark-excel"></i> Excel', className: "dt-button" },
        { extend: "csvHtml5",   text: '<i class="bi bi-filetype-csv"></i> CSV',         className: "dt-button" }
      ]
    }));

    // Delegación de clic en fila (funciona con paginación de DataTables)
    $(selSum + " tbody").off("click.resp").on("click.resp", "tr.resp-row", function () {
      const nombre = $(this).attr("data-resp");
      if (nombre && byResp[nombre]) {
        _respDetalleActual = nombre;
        renderResponsableDetalle(byResp[nombre]);
      }
    });

    // Cerrar detalle
    const closeBtn = document.getElementById("btnCerrarRespDetalle");
    if (closeBtn && !closeBtn._wired) {
      closeBtn._wired = true;
      closeBtn.addEventListener("click", function () {
        const panel = document.getElementById("panelRespDetalle");
        if (panel) panel.style.display = "none";
        _respDetalleActual = null;
      });
    }

    // Re-abrir detalle si había uno seleccionado
    if (_respDetalleActual && byResp[_respDetalleActual]) {
      renderResponsableDetalle(byResp[_respDetalleActual]);
    }
  }

  function renderResponsableDetalle(respData) {
    if (!respData) return;
    const panel = document.getElementById("panelRespDetalle");
    if (panel) panel.style.display = "";

    const el = document.getElementById("respDetalleNombre");
    if (el) el.innerHTML = '<i class="bi bi-person-circle"></i> ' + esc(respData.nombre);
    const areaEl = document.getElementById("respDetalleArea");
    if (areaEl) areaEl.textContent = respData.areasList;

    /* KPI del responsable */
    const kpiGrid = document.getElementById("kpiRespDetalle");
    if (kpiGrid) {
      const tiempoSolStr = respData.avgTiempoSolucionados !== null ? respData.avgTiempoSolucionados + " días" : "—";
      const tiempoAbStr  = respData.avgTiempoAbiertos !== null ? respData.avgTiempoAbiertos + " días" : "—";
      const slaStr       = respData.slaCumplido !== null ? respData.slaCumplido + "%" : "—";
      kpiGrid.innerHTML =
        kpi("Total asignados",          respData.totalCasos,         "info",    "bi-collection",         "todos los procesos") +
        kpi("Abiertos",                  respData.abiertos,           "info",    "bi-folder2-open",       "pendientes de resolución") +
        kpi("Solucionados",              respData.solucionados,       "sla",     "bi-check2-circle",      respData.tasaResolucion + "% de tasa de resolución") +
        kpi("Vencidos activos",          respData.vencidosActivos,    "vencido", "bi-x-octagon",          "requieren acción inmediata") +
        kpi("Críticos activos",          respData.criticosActivos,    "critico", "bi-exclamation-triangle","SLA ≥ 90%") +
        kpi("SLA cumplido (soluc.)",     slaStr,                      "sla",     "bi-patch-check",        "% de resueltos dentro de SLA") +
        kpi("Días prom. de resolución",  tiempoSolStr,                "normal",  "bi-clock-history",      "tiempo promedio al cerrar") +
        kpi("Antigüedad prom. (abiertos)", tiempoAbStr,               "riesgo",  "bi-hourglass-split",    "backlog acumulado sin resolver");
    }

    /* Gráfico: distribución de abiertos por clasificación */
    const clasifData = {
      Normal: respData.normalActivos,
      Riesgo: respData.riesgoActivos,
      Critico: respData.criticosActivos,
      Vencido: respData.vencidosActivos
    };
    renderChart("chartRespClasif", "doughnut",
      toChartDataDoughnut(clasifData, STATUS_LABELS, STATUS_COLORS), doughnutOpts());

    /* Gráfico: categorías más frecuentes (todos sus casos) */
    renderChart("chartRespCategorias", "bar",
      toChartDataBar(respData.categorias, "#8C0F13", 8), horizontalBarOpts());

    /* Tabla de casos abiertos */
    const allRaw = [];
    AREAS.forEach(function (a) { (STATE.rawData[a] || []).forEach(function (r) { allRaw.push(r); }); });
    const casosAbiertos = allRaw
      .filter(function (r) { return r["Responsable"] === respData.nombre && r["Estado"] !== "Solucionado"; })
      .sort(function (a, b) { return b["Progreso"] - a["Progreso"]; });

    const cntEl = document.getElementById("respDetalleCasosCount");
    if (cntEl) cntEl.textContent = casosAbiertos.length + " caso" + (casosAbiertos.length !== 1 ? "s" : "") + " abierto" + (casosAbiertos.length !== 1 ? "s" : "");

    const selCasos = "#tableRespCasos";
    if (dtRegistry[selCasos]) { try { dtRegistry[selCasos].destroy(); } catch (e) { /* noop */ } delete dtRegistry[selCasos]; }
    const tbodyCasos = document.querySelector(selCasos + " tbody");
    if (tbodyCasos) {
      tbodyCasos.innerHTML = casosAbiertos.map(function (r) {
        const cls = effectiveClass(r);
        const rowCls = cls === "Vencido" ? "row--vencido" : (cls === "Critico" ? "row--critico" : "");
        return (
          '<tr class="' + rowCls + '">' +
          '<td>' + esc(r["No. Caso"]) + '</td>' +
          '<td>' + areaChipHTML(r["_area"]) + '</td>' +
          '<td>' + esc(r["Fecha de registro"]) + '</td>' +
          '<td>' + esc(r["Estado"]) + '</td>' +
          '<td>' + esc(r["Categoría"]) + '</td>' +
          '<td>' + esc(r["Servicio"]) + '</td>' +
          '<td data-order="' + (r["Tiempo transcurrido"] || 0) + '">' + (r["Tiempo transcurrido"] || 0).toFixed(1) + ' días</td>' +
          '<td data-order="' + r["Progreso"] + '">' + progressCellHTML(r["Progreso"], cls) + '</td>' +
          '<td>' + esc(r["Valor adicional"] || "") + '</td>' +
          '</tr>'
        );
      }).join("");
    }
    dtRegistry[selCasos] = $(selCasos).DataTable(Object.assign({ language: DT_LANG_ES }, {
      paging: true, pageLength: 10, order: [[7, "desc"]],
      dom: "frtipB",
      buttons: [
        { extend: "excelHtml5", text: '<i class="bi bi-file-earmark-excel"></i> Excel', className: "dt-button" },
        { extend: "csvHtml5",   text: '<i class="bi bi-filetype-csv"></i> CSV',         className: "dt-button" }
      ]
    }));

    if (panel) panel.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  /* ====================== BÚSQUEDA GLOBAL (TOPBAR) ====================== */

  let globalSearchValue = "";

  function searchAllTables(val) {
    Object.keys(dtRegistry).forEach(function (sel) {
      try { if (dtRegistry[sel]) dtRegistry[sel].search(val).draw(); } catch (e) {}
    });
  }

  function wireGlobalSearch() {
    const input = document.getElementById("globalSearch");
    if (!input) return;
    let debounceTimer = null;
    input.addEventListener("input", function () {
      const val = this.value;
      clearTimeout(debounceTimer);
      debounceTimer = setTimeout(function () {
        globalSearchValue = val;
        searchAllTables(val);
      }, 250);
    });
    input.addEventListener("keydown", function (e) {
      if (e.key === "Escape") {
        input.value = "";
        globalSearchValue = "";
        searchAllTables("");
      }
    });
  }

  /* ============================ EVENTOS UI ============================ */

  function wireNav() {
    document.querySelectorAll(".nav-link[data-view]").forEach(function (btn) {
      btn.addEventListener("click", function () {
        switchView(btn.getAttribute("data-view"));
        closeAllDropdowns();
      });
    });
  }

  function wireSidebarMobile() {
    const sidebar = document.getElementById("sidebar");
    const overlay = document.getElementById("sidebarOverlay");
    const openBtn = document.getElementById("sidebarOpen");
    const closeBtn = document.getElementById("sidebarClose");

    function openSidebar() {
      if (sidebar) sidebar.classList.add("is-open");
      if (overlay) overlay.classList.add("is-open");
    }
    function closeSidebar() {
      if (sidebar) sidebar.classList.remove("is-open");
      if (overlay) overlay.classList.remove("is-open");
    }
    if (openBtn) openBtn.addEventListener("click", openSidebar);
    if (closeBtn) closeBtn.addEventListener("click", closeSidebar);
    if (overlay) overlay.addEventListener("click", closeSidebar);
  }

  function wireRefreshButton() {
    const btn = document.getElementById("refreshNowBtn");
    if (!btn) return;
    btn.addEventListener("click", function () {
      btn.classList.add("is-spinning");
      loadAllData(true).then(function () {
        setTimeout(function () { btn.classList.remove("is-spinning"); }, 400);
      });
    });
  }

  /* ============================== INICIO ============================== */

  document.addEventListener("DOMContentLoaded", function () {
    setChartDefaults();
    injectModuleTemplates();
    wireNav();
    wireSidebarMobile();
    wireRefreshButton();
    wireGlobalSearch();

    loadAllData(true).then(function () {
      setInterval(function () { loadAllData(false); }, CONFIG.refreshIntervalMs);
    });
  });

})();
