const PAGE_SIZE = 50;
const numberFormatter = new Intl.NumberFormat("ja-JP");
const decimalFormatter = new Intl.NumberFormat("ja-JP", {
  minimumFractionDigits: 1,
  maximumFractionDigits: 1,
});
const ratioFormatter = new Intl.NumberFormat("ja-JP", {
  minimumFractionDigits: 2,
  maximumFractionDigits: 2,
});
const percentFormatter = new Intl.NumberFormat("ja-JP", {
  minimumFractionDigits: 1,
  maximumFractionDigits: 1,
});

const WAGE_BAND_ORDER = [
  "1万円未満",
  "1.0万-1.5万円",
  "1.5万-2.0万円",
  "2.0万-3.0万円",
  "3.0万-5.0万円",
  "5.0万円以上",
];
const CAPACITY_BAND_ORDER = ["10名以下", "11-20名", "21-30名", "31-40名", "41名以上"];
const QUADRANT_ORDER = [
  "高工賃 × 高稼働",
  "高工賃 × 低稼働",
  "低工賃 × 高稼働",
  "低工賃 × 低稼働",
];
const STAFFING_QUADRANT_ORDER = [
  "高工賃 × 厚い人員",
  "高工賃 × 少ない人員",
  "低工賃 × 厚い人員",
  "低工賃 × 少ない人員",
];
const FILTER_PRESETS = {
  all: {},
  "osaka-city": {
    municipality: "大阪市",
  },
  "osaka-wam": {
    municipality: "大阪市",
    wamOnly: true,
  },
  "high-wage": {
    outlierFlag: "high",
  },
  "fix-priority": {
    municipality: "大阪市",
    wamOnly: true,
    staffingOutlier: "high",
  },
  unanswered: {
    responseStatus: "unanswered",
  },
};

const state = {
  records: [],
  filteredRecords: [],
  dashboard: null,
  sortKey: "wage_ratio_to_overall_mean",
  sortDirection: "desc",
  currentPage: 1,
  filters: {
    search: "",
    municipality: "all",
    corporationType: "all",
    responseStatus: "all",
    outlierFlag: "all",
    capacityBand: "all",
    quadrant: "all",
    wamMatch: "all",
    transport: "all",
    staffingOutlier: "all",
    primaryActivity: "all",
    mealSupport: "all",
    managerMulti: "all",
    newOnly: false,
    homeUseOnly: false,
    noufukuOnly: false,
    wamOnly: false,
  },
  selectedOfficeNo: null,
  activePreset: "all",
};

document.addEventListener("DOMContentLoaded", () => {
  init().catch((error) => {
    renderError(error);
  });
});

async function init() {
  const response = await fetch("./data/dashboard-data.json", { cache: "no-store" });
  if (!response.ok) {
    throw new Error("dashboard data could not be loaded");
  }

  const dashboard = await response.json();
  state.dashboard = dashboard;
  renderMeta(dashboard);
  renderQuality(dashboard);
  renderLoadingState();
  state.records = await loadDashboardRecords(dashboard);
  populateFilterOptions(state.records);
  bindEvents();
  syncPresetButtons();
  applyFilters();
}

async function loadDashboardRecords(dashboard) {
  if (Array.isArray(dashboard.records) && dashboard.records.length) {
    return dashboard.records;
  }

  const chunkPaths = dashboard.data_files?.record_chunks ?? [];
  if (!chunkPaths.length) {
    return [];
  }

  const chunks = await Promise.all(
    chunkPaths.map(async (path) => {
      const response = await fetch(`./${path}`, { cache: "no-store" });
      if (!response.ok) {
        throw new Error(`record chunk could not be loaded: ${path}`);
      }
      return response.json();
    })
  );
  return chunks.flat();
}

function bindEvents() {
  bindInput("searchInput", "input", (value) => {
    state.filters.search = value.trim();
  });
  bindInput("municipalitySelect", "change", (value) => {
    state.filters.municipality = value;
  });
  bindInput("corporationTypeSelect", "change", (value) => {
    state.filters.corporationType = value;
  });
  bindInput("responseStatusSelect", "change", (value) => {
    state.filters.responseStatus = value;
  });
  bindInput("outlierFlagSelect", "change", (value) => {
    state.filters.outlierFlag = value;
  });
  bindInput("capacityBandSelect", "change", (value) => {
    state.filters.capacityBand = value;
  });
  bindInput("quadrantSelect", "change", (value) => {
    state.filters.quadrant = value;
  });
  bindInput("wamMatchSelect", "change", (value) => {
    state.filters.wamMatch = value;
  });
  bindInput("transportSelect", "change", (value) => {
    state.filters.transport = value;
  });
  bindInput("staffingOutlierSelect", "change", (value) => {
    state.filters.staffingOutlier = value;
  });
  bindInput("primaryActivitySelect", "change", (value) => {
    state.filters.primaryActivity = value;
  });
  bindInput("mealSupportSelect", "change", (value) => {
    state.filters.mealSupport = value;
  });
  bindInput("managerMultiSelect", "change", (value) => {
    state.filters.managerMulti = value;
  });
  bindCheckbox("newOnlyCheckbox", "newOnly");
  bindCheckbox("homeUseOnlyCheckbox", "homeUseOnly");
  bindCheckbox("noufukuOnlyCheckbox", "noufukuOnly");
  bindCheckbox("wamOnlyCheckbox", "wamOnly");

  document.getElementById("resetFiltersButton").addEventListener("click", () => {
    resetFilters();
    applyFilters();
  });
  document.getElementById("downloadCsvButton").addEventListener("click", downloadFilteredCsv);
  document.getElementById("prevPageButton").addEventListener("click", () => changePage(-1));
  document.getElementById("nextPageButton").addEventListener("click", () => changePage(1));
  document.getElementById("focusSelectedButton").addEventListener("click", () => {
    const selectedRecord = getSelectedRecord();
    if (!selectedRecord) return;
    selectRecord(selectedRecord.office_no, { scroll: true });
  });

  document.querySelectorAll(".preset-chip[data-preset]").forEach((button) => {
    button.addEventListener("click", () => {
      const { preset } = button.dataset;
      if (!preset) return;
      applyPreset(preset);
    });
  });

  document.querySelectorAll("th button[data-sort-key]").forEach((button) => {
    button.addEventListener("click", () => {
      const { sortKey } = button.dataset;
      if (!sortKey) return;
      if (state.sortKey === sortKey) {
        state.sortDirection = state.sortDirection === "asc" ? "desc" : "asc";
      } else {
        state.sortKey = sortKey;
        state.sortDirection = defaultDirection(sortKey);
      }
      state.currentPage = 1;
      applyFilters();
    });
  });

  document.body.addEventListener("click", (event) => {
    const trigger = event.target.closest("[data-select-office]");
    if (!trigger) return;
    const officeNo = trigger.getAttribute("data-select-office");
    if (!officeNo) return;
    selectRecord(officeNo, { scroll: true });
  });
}

function bindInput(id, eventName, setter) {
  document.getElementById(id).addEventListener(eventName, (event) => {
    setter(event.target.value);
    state.activePreset = null;
    syncPresetButtons();
    state.currentPage = 1;
    applyFilters();
  });
}

function bindCheckbox(id, key) {
  document.getElementById(id).addEventListener("change", (event) => {
    state.filters[key] = event.target.checked;
    state.activePreset = null;
    syncPresetButtons();
    state.currentPage = 1;
    applyFilters();
  });
}

function resetFilters() {
  state.filters = {
    search: "",
    municipality: "all",
    corporationType: "all",
    responseStatus: "all",
    outlierFlag: "all",
    capacityBand: "all",
    quadrant: "all",
    wamMatch: "all",
    transport: "all",
    staffingOutlier: "all",
    primaryActivity: "all",
    mealSupport: "all",
    managerMulti: "all",
    newOnly: false,
    homeUseOnly: false,
    noufukuOnly: false,
    wamOnly: false,
  };
  state.activePreset = "all";

  [
    "searchInput",
    "municipalitySelect",
    "corporationTypeSelect",
    "responseStatusSelect",
    "outlierFlagSelect",
    "capacityBandSelect",
    "quadrantSelect",
    "wamMatchSelect",
    "transportSelect",
    "staffingOutlierSelect",
    "primaryActivitySelect",
    "mealSupportSelect",
    "managerMultiSelect",
  ].forEach((id) => {
    document.getElementById(id).value = id === "searchInput" ? "" : "all";
  });
  [
    "newOnlyCheckbox",
    "homeUseOnlyCheckbox",
    "noufukuOnlyCheckbox",
    "wamOnlyCheckbox",
  ].forEach((id) => {
    document.getElementById(id).checked = false;
  });
  syncPresetButtons();
  state.currentPage = 1;
}

function applyPreset(preset) {
  resetFilters();
  state.activePreset = preset;
  Object.assign(state.filters, FILTER_PRESETS[preset] ?? {});
  syncFilterControls();
  syncPresetButtons();
  applyFilters();
}

function syncFilterControls() {
  document.getElementById("searchInput").value = state.filters.search;
  document.getElementById("municipalitySelect").value = state.filters.municipality;
  document.getElementById("corporationTypeSelect").value = state.filters.corporationType;
  document.getElementById("responseStatusSelect").value = state.filters.responseStatus;
  document.getElementById("outlierFlagSelect").value = state.filters.outlierFlag;
  document.getElementById("capacityBandSelect").value = state.filters.capacityBand;
  document.getElementById("quadrantSelect").value = state.filters.quadrant;
  document.getElementById("wamMatchSelect").value = state.filters.wamMatch;
  document.getElementById("transportSelect").value = state.filters.transport;
  document.getElementById("staffingOutlierSelect").value = state.filters.staffingOutlier;
  document.getElementById("primaryActivitySelect").value = state.filters.primaryActivity;
  document.getElementById("mealSupportSelect").value = state.filters.mealSupport;
  document.getElementById("managerMultiSelect").value = state.filters.managerMulti;
  document.getElementById("newOnlyCheckbox").checked = state.filters.newOnly;
  document.getElementById("homeUseOnlyCheckbox").checked = state.filters.homeUseOnly;
  document.getElementById("noufukuOnlyCheckbox").checked = state.filters.noufukuOnly;
  document.getElementById("wamOnlyCheckbox").checked = state.filters.wamOnly;
}

function syncPresetButtons() {
  document.querySelectorAll(".preset-chip[data-preset]").forEach((button) => {
    button.classList.toggle("is-active", button.dataset.preset === state.activePreset);
  });
}

function syncSelectedRecord(filtered) {
  if (!filtered.length) {
    state.selectedOfficeNo = null;
    return;
  }
  const selectedExists = filtered.some((record) => String(record.office_no) === String(state.selectedOfficeNo));
  if (!selectedExists) {
    state.selectedOfficeNo = String(filtered[0].office_no);
  }
}

function getSelectedRecord() {
  if (state.selectedOfficeNo == null) return null;
  return (
    state.filteredRecords.find((record) => String(record.office_no) === String(state.selectedOfficeNo)) ??
    state.records.find((record) => String(record.office_no) === String(state.selectedOfficeNo)) ??
    null
  );
}

function selectRecord(officeNo, options = {}) {
  const record = state.filteredRecords.find((item) => String(item.office_no) === String(officeNo));
  if (!record) return;
  state.selectedOfficeNo = String(officeNo);
  renderCharts(state.filteredRecords);
  renderDetail(record);
  renderTable(state.filteredRecords);
  if (options.scroll) {
    document.getElementById("detailHeading").scrollIntoView({ behavior: "smooth", block: "start" });
  }
}

function applyFilters() {
  const filtered = state.records
    .filter((record) => matchesFilters(record, state.filters))
    .sort(compareRecords(state.sortKey, state.sortDirection));

  state.filteredRecords = filtered;
  const totalPages = Math.max(1, Math.ceil(filtered.length / PAGE_SIZE));
  state.currentPage = Math.min(state.currentPage, totalPages);
  syncSelectedRecord(filtered);

  renderActiveFilterSummary(filtered);
  renderInsights(filtered);
  renderStrategy(filtered);
  renderStats(filtered);
  renderBenchmarks(filtered);
  renderCharts(filtered);
  renderAnomalies(filtered);
  renderDetail(getSelectedRecord());
  renderTable(filtered);
}

function matchesFilters(record, filters) {
  if (filters.search) {
    const haystack = [
      record.office_no,
      record.municipality,
      record.corporation_name,
      record.office_name,
      record.corporation_type_label,
      record.remarks,
      record.wam_primary_activity_type,
      record.wam_primary_activity_detail,
    ]
      .filter(Boolean)
      .join(" ")
      .toLowerCase();
    if (!haystack.includes(filters.search.toLowerCase())) {
      return false;
    }
  }

  if (filters.municipality !== "all" && record.municipality !== filters.municipality) return false;
  if (filters.corporationType !== "all" && record.corporation_type_label !== filters.corporationType) {
    return false;
  }
  if (filters.responseStatus !== "all" && record.response_status !== filters.responseStatus) {
    return false;
  }
  if (filters.outlierFlag !== "all" && (record.wage_outlier_flag ?? "none") !== filters.outlierFlag) {
    return false;
  }
  if (filters.capacityBand !== "all" && record.capacity_band_label !== filters.capacityBand) {
    return false;
  }
  if (filters.quadrant !== "all" && record.market_position_quadrant !== filters.quadrant) {
    return false;
  }
  if (filters.wamMatch !== "all" && (record.wam_match_status ?? "unmatched") !== filters.wamMatch) {
    return false;
  }
  if (filters.transport !== "all") {
    const transportValue = record.wam_transport_available === true ? "true" : record.wam_transport_available === false ? "false" : "unknown";
    if (transportValue !== filters.transport) return false;
  }
  if (filters.staffingOutlier !== "all" && (record.wam_staffing_outlier_flag ?? "none") !== filters.staffingOutlier) {
    return false;
  }
  if (filters.primaryActivity !== "all" && (record.wam_primary_activity_type ?? "unknown") !== filters.primaryActivity) {
    return false;
  }
  if (filters.mealSupport !== "all") {
    const mealValue = record.wam_meal_support_addon === true ? "true" : record.wam_meal_support_addon === false ? "false" : "unknown";
    if (mealValue !== filters.mealSupport) return false;
  }
  if (filters.managerMulti !== "all") {
    const managerValue = record.wam_manager_multi_post === true ? "true" : record.wam_manager_multi_post === false ? "false" : "unknown";
    if (managerValue !== filters.managerMulti) return false;
  }
  if (filters.newOnly && !record.is_new_office) return false;
  if (filters.homeUseOnly && !record.home_use_active) return false;
  if (filters.noufukuOnly && !record.noufuku_active) return false;
  if (filters.wamOnly && record.wam_match_status !== "matched") return false;
  return true;
}

function compareRecords(sortKey, sortDirection) {
  return (left, right) => {
    const leftValue = comparableValue(left[sortKey]);
    const rightValue = comparableValue(right[sortKey]);
    if (leftValue == null && rightValue == null) return 0;
    if (leftValue == null) return 1;
    if (rightValue == null) return -1;

    let comparison = 0;
    if (typeof leftValue === "string" && typeof rightValue === "string") {
      comparison = leftValue.localeCompare(rightValue, "ja");
    } else {
      comparison = leftValue > rightValue ? 1 : leftValue < rightValue ? -1 : 0;
    }
    return sortDirection === "asc" ? comparison : -comparison;
  };
}

function comparableValue(value) {
  if (value == null || value === "") return null;
  if (typeof value === "boolean") return value ? 1 : 0;
  if (value === "matched") return 2;
  if (value === "unmatched") return 1;
  if (value === "high") return 3;
  if (value === "low") return 2;
  return value;
}

function defaultDirection(sortKey) {
  return [
    "office_no",
    "municipality",
    "corporation_type_label",
    "corporation_name",
    "office_name",
    "response_status",
    "market_position_quadrant",
    "wam_match_status",
    "wam_staffing_efficiency_quadrant",
  ].includes(sortKey)
    ? "asc"
    : "desc";
}

function populateFilterOptions(records) {
  populateSelect("municipalitySelect", "all", "すべての市町村", uniqueValues(records, "municipality"));
  populateSelect("corporationTypeSelect", "all", "すべての法人種別", uniqueValues(records, "corporation_type_label"));
  populateSelect("responseStatusSelect", "all", "すべての回答状態", ["answered", "annotated", "unanswered"]);
  populateSelect("outlierFlagSelect", "all", "すべての工賃異常値", ["high", "low", "none"]);
  populateSelect("capacityBandSelect", "all", "すべての定員帯", CAPACITY_BAND_ORDER);
  populateSelect("quadrantSelect", "all", "すべての運営ポジション", QUADRANT_ORDER);
  populateSelect("wamMatchSelect", "all", "すべてのWAM状態", ["matched", "unmatched"]);
  populateSelect("transportSelect", "all", "すべての送迎状態", ["true", "false", "unknown"]);
  populateSelect("staffingOutlierSelect", "all", "すべての人員異常値", ["high", "low", "none"]);
  populateSelect("primaryActivitySelect", "all", "すべての主活動", uniqueValues(records, "wam_primary_activity_type"));
  populateSelect("mealSupportSelect", "all", "すべての食事加算", ["true", "false", "unknown"]);
  populateSelect("managerMultiSelect", "all", "すべての管理者兼務", ["true", "false", "unknown"]);
}

function populateSelect(id, defaultValue, defaultLabel, values) {
  const root = document.getElementById(id);
  const options = [`<option value="${defaultValue}">${defaultLabel}</option>`];
  values.forEach((value) => {
    options.push(`<option value="${escapeHtml(value)}">${escapeHtml(labelForSelect(value))}</option>`);
  });
  root.innerHTML = options.join("");
}

function uniqueValues(records, key) {
  return [...new Set(records.map((record) => record[key]).filter(Boolean))].sort((left, right) =>
    String(left).localeCompare(String(right), "ja")
  );
}

function labelForSelect(value) {
  if (value === "answered") return "回答済み";
  if (value === "annotated") return "注記あり";
  if (value === "unanswered") return "未回答";
  if (value === "high") return "高い";
  if (value === "low") return "低い";
  if (value === "none") return "通常範囲";
  if (value === "matched") return "一致";
  if (value === "unmatched") return "未一致";
  if (value === "true") return "あり";
  if (value === "false") return "なし";
  if (value === "unknown") return "不明";
  return value;
}

function renderLoadingState() {
  const loadingCard = `<div class="loading-message">データを読み込み中...</div>`;
  document.getElementById("activeFilterSummary").textContent = "データを読み込み中...";
  document.getElementById("insightList").innerHTML = loadingCard;
  document.getElementById("growthList").innerHTML = loadingCard;
  document.getElementById("fixList").innerHTML = loadingCard;
  document.getElementById("reviewList").innerHTML = loadingCard;
  document.getElementById("statsGrid").innerHTML = loadingCard;
  document.getElementById("leaderboardList").innerHTML = loadingCard;
  document.getElementById("watchlistList").innerHTML = loadingCard;
  [
    "municipalityChart",
    "wageChart",
    "corporationChart",
    "capacityChart",
    "quadrantChart",
    "wamCoverageChart",
    "staffingQuadrantChart",
    "outlierChart",
    "staffingRoleChart",
    "featureChart",
    "wageBandChart",
    "utilizationScatter",
    "staffingScatter",
    "anomalyList",
    "detailContent",
  ].forEach((id) => {
    const element = document.getElementById(id);
    if (element) {
      element.innerHTML = loadingCard;
    }
  });
  document.getElementById("recordsTableBody").innerHTML = `<tr><td colspan="19">${loadingCard}</td></tr>`;
}

function renderMeta(dashboard) {
  const meta = dashboard.meta ?? {};
  const matchSummary = dashboard.summary?.wam_match_summary ?? {};
  const updatedAt = meta.generated_at ? new Date(meta.generated_at) : null;
  document.getElementById("updatedAt").textContent =
    updatedAt && !Number.isNaN(updatedAt.valueOf()) ? updatedAt.toLocaleString("ja-JP") : "-";
  document.getElementById("sourceWorkbook").textContent = meta.source_workbook ?? "-";
  document.getElementById("totalRecords").textContent = formatCount(meta.record_count ?? state.records.length);
  document.getElementById("wamMatchedCount").textContent = formatCount(matchSummary.matched_record_count ?? 0);
}

function renderQuality(dashboard) {
  const issuesRoot = document.getElementById("issuesList");
  const notesRoot = document.getElementById("notesList");
  const issues = dashboard.issues ?? [];
  const notes = dashboard.notes ?? [];
  const wageStats = dashboard.analytics?.overall_wage_stats ?? {};
  const wamMatchSummary = dashboard.analytics?.wam_match_summary ?? {};

  const summaryCards = [
    `<article class="note-card"><strong>工賃異常値ルール</strong><p>中央値 ${formatMaybeYen(
      wageStats.median
    )}。IQR に基づく外れ値判定を使っている。</p></article>`,
    `<article class="note-card"><strong>WAM一致率</strong><p>大阪市レコード ${formatCount(
      wamMatchSummary.matched_record_count ?? 0
    )} / ${formatCount(wamMatchSummary.workbook_osaka_record_count ?? 0)} 件を一致。</p></article>`,
  ];

  issuesRoot.innerHTML =
    summaryCards.join("") +
    issues
      .map((issue) => {
        const detail = issue.detail ?? `${issue.sheet}: ${formatCount(issue.count ?? 0)} 件`;
        return `<article class="note-card"><strong>${escapeHtml(issue.sheet)}</strong><p>${escapeHtml(detail)}</p></article>`;
      })
      .join("");

  notesRoot.innerHTML = notes
    .slice(0, 8)
    .map(
      (note) =>
        `<article class="note-card"><strong>注記 ${formatCount(note.source_row)}</strong><p>${escapeHtml(note.note_text)}</p></article>`
    )
    .join("");
}

function renderActiveFilterSummary(records) {
  const filters = [];
  if (state.filters.search) filters.push(`検索: ${state.filters.search}`);
  if (state.filters.municipality !== "all") filters.push(`市町村: ${state.filters.municipality}`);
  if (state.filters.corporationType !== "all") filters.push(`法人種別: ${state.filters.corporationType}`);
  if (state.filters.responseStatus !== "all") filters.push(`回答状態: ${labelForSelect(state.filters.responseStatus)}`);
  if (state.filters.outlierFlag !== "all") filters.push(`工賃異常: ${labelForSelect(state.filters.outlierFlag)}`);
  if (state.filters.capacityBand !== "all") filters.push(`定員帯: ${state.filters.capacityBand}`);
  if (state.filters.quadrant !== "all") filters.push(`工賃象限: ${state.filters.quadrant}`);
  if (state.filters.wamMatch !== "all") filters.push(`WAM: ${labelForSelect(state.filters.wamMatch)}`);
  if (state.filters.transport !== "all") filters.push(`送迎: ${labelForSelect(state.filters.transport)}`);
  if (state.filters.staffingOutlier !== "all") filters.push(`人員異常: ${labelForSelect(state.filters.staffingOutlier)}`);
  if (state.filters.primaryActivity !== "all") filters.push(`主活動: ${state.filters.primaryActivity}`);
  if (state.filters.mealSupport !== "all") filters.push(`食事加算: ${labelForSelect(state.filters.mealSupport)}`);
  if (state.filters.managerMulti !== "all") filters.push(`管理者兼務: ${labelForSelect(state.filters.managerMulti)}`);
  if (state.filters.newOnly) filters.push("新設のみ");
  if (state.filters.homeUseOnly) filters.push("在宅利用あり");
  if (state.filters.noufukuOnly) filters.push("農福連携あり");
  if (state.filters.wamOnly) filters.push("WAM一致のみ");

  document.getElementById("activeFilterSummary").textContent = filters.length
    ? `適用中: ${filters.join(" / ")}`
    : `${formatCount(records.length)} 件を表示中。フィルタ未指定。`;
}

function renderInsights(records) {
  const root = document.getElementById("insightList");
  if (!records.length) {
    root.innerHTML = document.getElementById("emptyStateTemplate").innerHTML;
    return;
  }

  const matched = matchedRecords(records);
  const wages = numericValues(records, "average_wage_yen");
  const wageStats = computeLocalStats(wages);
  const transportTrue = meanFor(matched.filter((record) => record.wam_transport_available === true), "average_wage_yen");
  const transportFalse = meanFor(matched.filter((record) => record.wam_transport_available === false), "average_wage_yen");
  const transportDelta =
    transportTrue != null && transportFalse != null ? transportTrue - transportFalse : null;
  const leanHigh = matched
    .filter((record) => record.wam_staffing_efficiency_quadrant === "高工賃 × 少ない人員")
    .sort((left, right) => (right.average_wage_yen ?? 0) - (left.average_wage_yen ?? 0))[0];
  const heavyLow = matched
    .filter((record) => record.wam_staffing_efficiency_quadrant === "低工賃 × 厚い人員")
    .sort((left, right) => (left.average_wage_yen ?? Infinity) - (right.average_wage_yen ?? Infinity))[0];
  const staffingOutlierCount = matched.filter((record) => record.wam_staffing_outlier_flag).length;

  const insights = [
    {
      title: "市場基準",
      body: `中央値 ${formatMaybeYen(wageStats.median)}。上位10%目安は ${formatMaybeYen(wageStats.p90)}。`,
    },
    {
      title: "WAM一致率",
      body: `現在の条件で WAM 詳細一致は ${formatCount(matched.length)} 件。人員・送迎・運営体制まで比較できる。`,
    },
    {
      title: "省人で高工賃",
      body: leanHigh
        ? `${leanHigh.office_name ?? "-"} は ${formatWageText(
            leanHigh.average_wage_yen
          )}、主要支援FTE/定員 ${formatPercent(leanHigh.wam_key_staff_fte_per_capacity)}。`
        : "高工賃 × 少ない人員の候補がまだ見えていない。",
    },
    {
      title: "送迎の影響",
      body:
        transportDelta == null
          ? "送迎あり/なしの比較に十分な一致件数がない。"
          : `送迎ありの平均工賃は ${formatMaybeYen(transportTrue)}、なしは ${formatMaybeYen(
              transportFalse
            )}。差分は ${formatSignedYen(transportDelta)}。`,
    },
    {
      title: "改善優先候補",
      body: heavyLow
        ? `${heavyLow.office_name ?? "-"} は ${formatWageText(
            heavyLow.average_wage_yen
          )}、主要支援FTE/定員 ${formatPercent(heavyLow.wam_key_staff_fte_per_capacity)}。人員設計の再点検余地が大きい。`
        : `人員異常値は ${formatCount(staffingOutlierCount)} 件。`,
    },
  ];

  root.innerHTML = insights
    .map(
      (insight) => `
        <article class="insight-card">
          <h3>${escapeHtml(insight.title)}</h3>
          <p>${escapeHtml(insight.body)}</p>
        </article>
      `
    )
    .join("");
}

function renderStrategy(records) {
  const rootSummary = document.getElementById("strategySummary");
  if (!records.length) {
    rootSummary.textContent = "条件に合うレコードがない";
    ["growthList", "fixList", "reviewList"].forEach((id) => {
      document.getElementById(id).innerHTML = document.getElementById("emptyStateTemplate").innerHTML;
    });
    return;
  }

  const growthCandidates = records
    .filter(
      (record) =>
        isNumber(record.average_wage_yen) &&
        isNumber(record.daily_user_capacity_ratio) &&
        record.wam_match_status === "matched" &&
        ((record.wam_staffing_efficiency_quadrant === "高工賃 × 少ない人員" &&
          record.daily_user_capacity_ratio < 0.9) ||
          (record.wage_ratio_to_overall_mean >= 1.15 && record.daily_user_capacity_ratio < 0.8))
    )
    .sort((left, right) => growthScore(right) - growthScore(left))
    .slice(0, 6);

  const fixCandidates = records
    .filter(
      (record) =>
        isNumber(record.average_wage_yen) &&
        (record.wam_staffing_efficiency_quadrant === "低工賃 × 厚い人員" ||
          (record.wam_staffing_outlier_flag === "high" && (record.wage_ratio_to_overall_mean ?? 0) < 0.95) ||
          (isNumber(record.daily_user_capacity_ratio) && record.daily_user_capacity_ratio < 0.55))
    )
    .sort((left, right) => fixScore(right) - fixScore(left))
    .slice(0, 6);

  const reviewCandidates = records
    .filter(
      (record) =>
        record.response_status === "unanswered" ||
        record.wam_match_status !== "matched" ||
        record.is_new_office ||
        Boolean(record.average_wage_error)
    )
    .sort((left, right) => reviewScore(right) - reviewScore(left))
    .slice(0, 6);

  rootSummary.textContent = `伸ばす ${formatCount(growthCandidates.length)} 件 / 改善 ${formatCount(
    fixCandidates.length
  )} 件 / 確認 ${formatCount(reviewCandidates.length)} 件`;

  renderStrategyList("growthList", growthCandidates, "伸ばす候補がまだない", buildGrowthReason);
  renderStrategyList("fixList", fixCandidates, "改善優先がまだない", buildFixReason);
  renderStrategyList("reviewList", reviewCandidates, "確認優先がまだない", buildReviewReason);
}

function renderStrategyList(id, records, emptyLabel, buildReason) {
  const root = document.getElementById(id);
  if (!records.length) {
    root.innerHTML = `<div class="empty-state"><h3>${escapeHtml(emptyLabel)}</h3></div>`;
    return;
  }

  root.innerHTML = records
    .map(
      (record) => `
        <button class="strategy-item" data-select-office="${escapeHtml(record.office_no)}" type="button">
          <span class="strategy-kicker">${escapeHtml(record.municipality ?? "-")} / No.${escapeHtml(record.office_no ?? "-")}</span>
          <strong>${escapeHtml(record.office_name ?? "-")}</strong>
          <span>${escapeHtml(record.corporation_name ?? "-")}</span>
          <p>${escapeHtml(buildReason(record))}</p>
        </button>
      `
    )
    .join("");
}

function renderStats(records) {
  const matched = matchedRecords(records);
  const wages = numericValues(records, "average_wage_yen");
  const wageStats = computeLocalStats(wages);
  const welfareStaffFte = meanFor(matched, "wam_welfare_staff_fte_total");
  const keyStaffPerCapacity = meanFor(matched, "wam_key_staff_fte_per_capacity");
  const transportRate = ratioOf(matched, (record) => record.wam_transport_available === true);
  const managerMultiRate = ratioOf(matched, (record) => record.wam_manager_multi_post === true);
  const heavyLowCount = matched.filter((record) => record.wam_staffing_efficiency_quadrant === "低工賃 × 厚い人員").length;

  const cards = [
    { label: "表示件数", value: formatCount(records.length), hint: `全 ${formatCount(state.records.length)} 件中` },
    { label: "平均工賃", value: formatMaybeYen(wageStats.mean), hint: "未回答を除外" },
    { label: "中央値", value: formatMaybeYen(wageStats.median), hint: "市場の中心値" },
    { label: "高工賃異常", value: formatCount(records.filter((record) => record.wage_outlier_flag === "high").length), hint: "上振れ確認" },
    { label: "WAM一致", value: formatCount(matched.length), hint: "人員詳細あり" },
    { label: "平均福祉職員FTE", value: formatFte(welfareStaffFte), hint: "一致レコードのみ" },
    { label: "平均主要支援FTE/定員", value: formatPercent(keyStaffPerCapacity), hint: "サービス管理+就労支援+職業指導+生活支援" },
    { label: "送迎実施率", value: formatPercent(transportRate), hint: "一致レコードのみ" },
    { label: "管理者兼務率", value: formatPercent(managerMultiRate), hint: "一致レコードのみ" },
    { label: "低工賃×厚い人員", value: formatCount(heavyLowCount), hint: "改善優先候補" },
  ];

  document.getElementById("statsGrid").innerHTML = cards
    .map(
      (card) => `
        <article class="stat-card">
          <p>${escapeHtml(card.label)}</p>
          <strong>${escapeHtml(card.value)}</strong>
          <em>${escapeHtml(card.hint)}</em>
        </article>
      `
    )
    .join("");
}

function renderBenchmarks(records) {
  const matched = matchedRecords(records);
  renderBenchmarkList(
    "leaderboardList",
    matched
      .filter((record) => record.wam_staffing_efficiency_quadrant === "高工賃 × 少ない人員")
      .sort((left, right) => {
        const wageDiff = (right.average_wage_yen ?? 0) - (left.average_wage_yen ?? 0);
        if (wageDiff !== 0) return wageDiff;
        return (left.wam_key_staff_fte_per_capacity ?? Infinity) - (right.wam_key_staff_fte_per_capacity ?? Infinity);
      })
      .slice(0, 8),
    "候補がまだない"
  );
  renderBenchmarkList(
    "watchlistList",
    matched
      .filter((record) => record.wam_staffing_efficiency_quadrant === "低工賃 × 厚い人員")
      .sort((left, right) => {
        const wageDiff = (left.average_wage_yen ?? Infinity) - (right.average_wage_yen ?? Infinity);
        if (wageDiff !== 0) return wageDiff;
        return (right.wam_key_staff_fte_per_capacity ?? 0) - (left.wam_key_staff_fte_per_capacity ?? 0);
      })
      .slice(0, 8),
    "候補がまだない"
  );
}

function renderBenchmarkList(id, records, emptyLabel) {
  const root = document.getElementById(id);
  if (!records.length) {
    root.innerHTML = `<div class="empty-state"><h3>${escapeHtml(emptyLabel)}</h3></div>`;
    return;
  }

  root.innerHTML = records
    .map(
      (record) => `
        <button class="benchmark-item" data-select-office="${escapeHtml(record.office_no)}" type="button">
          <h3>${escapeHtml(record.office_name ?? "-")}</h3>
          <p>${escapeHtml(record.municipality ?? "-")} / ${escapeHtml(record.corporation_name ?? "-")}</p>
          <div class="benchmark-meta">
            <span class="metric-chip">${escapeHtml(formatWageText(record.average_wage_yen))}</span>
            <span class="metric-chip">${escapeHtml(formatPercent(record.wam_key_staff_fte_per_capacity))}</span>
            <span class="metric-chip">${escapeHtml(formatBool(record.wam_transport_available))}</span>
            <span class="metric-chip">${escapeHtml(record.wam_primary_activity_type ?? "-")}</span>
          </div>
        </button>
      `
    )
    .join("");
}

function renderCharts(records) {
  renderBarChart("municipalityChart", topCounts(records, "municipality", 10), formatCountSuffix("件"));
  renderBarChart("wageChart", topAverageWageByMunicipality(records, 8), formatCountSuffix("円", true));
  renderBarChart("corporationChart", topCounts(records, "corporation_type_label", 6), formatCountSuffix("件"));
  renderBarChart(
    "capacityChart",
    averageByOrderedGroup(records, "capacity_band_label", "average_wage_yen", CAPACITY_BAND_ORDER),
    formatCountSuffix("円", true)
  );
  renderBarChart("quadrantChart", orderedCounts(records, "market_position_quadrant", QUADRANT_ORDER), formatCountSuffix("件"));
  renderBarChart("wamCoverageChart", wamCoverageBreakdown(records), formatCountSuffix("件"));
  renderBarChart("staffingQuadrantChart", orderedCounts(records, "wam_staffing_efficiency_quadrant", STAFFING_QUADRANT_ORDER), formatCountSuffix("件"));
  renderBarChart("outlierChart", combinedOutlierBreakdown(records), formatCountSuffix("件"));
  renderBarChart("staffingRoleChart", staffingRoleAverages(records), (value) => formatFte(value));
  renderBarChart("featureChart", featureWageComparison(records), formatCountSuffix("円", true));
  renderBarChart("wageBandChart", orderedCounts(records, "wage_band_label", WAGE_BAND_ORDER), formatCountSuffix("件"));
  renderScatterChart("utilizationScatter", records, {
    xKey: "daily_user_capacity_ratio",
    yKey: "average_wage_yen",
    xLabel: "稼働率",
    yLabel: "平均工賃",
    yFormatter: (value) => formatMaybeYen(value),
    xFormatter: (value) => formatPercent(value),
  });
  renderScatterChart("staffingScatter", matchedRecords(records), {
    xKey: "wam_key_staff_fte_per_capacity",
    yKey: "average_wage_yen",
    xLabel: "主要支援FTE/定員",
    yLabel: "平均工賃",
    yFormatter: (value) => formatMaybeYen(value),
    xFormatter: (value) => formatPercent(value),
  });
}

function renderBarChart(id, items, valueFormatter) {
  const root = document.getElementById(id);
  if (!items.length) {
    root.innerHTML = document.getElementById("emptyStateTemplate").innerHTML;
    return;
  }
  const maxValue = Math.max(...items.map((item) => item.value), 1);
  root.innerHTML = items
    .map((item) => {
      const width = Math.max((item.value / maxValue) * 100, 3);
      return `
        <div class="bar-row">
          <span class="bar-label">${escapeHtml(item.label)}</span>
          <div class="bar-track"><div class="bar-fill" style="width: ${width}%"></div></div>
          <span class="bar-value">${escapeHtml(valueFormatter(item.value))}</span>
        </div>
      `;
    })
    .join("");
}

function renderScatterChart(id, records, config) {
  const root = document.getElementById(id);
  const points = records.filter(
    (record) => isNumber(record[config.xKey]) && isNumber(record[config.yKey])
  );
  if (!points.length) {
    root.innerHTML = document.getElementById("emptyStateTemplate").innerHTML;
    return;
  }

  const xValues = points.map((record) => record[config.xKey]);
  const yValues = points.map((record) => record[config.yKey]);
  const xMin = Math.min(...xValues);
  const xMax = Math.max(...xValues);
  const yMin = Math.min(...yValues);
  const yMax = Math.max(...yValues);
  const width = 320;
  const height = 220;
  const padding = 22;
  const spanX = xMax - xMin || 1;
  const spanY = yMax - yMin || 1;
  const selectedOfficeNo = String(state.selectedOfficeNo ?? "");

  const circles = points
    .map((record) => {
      const x = padding + ((record[config.xKey] - xMin) / spanX) * (width - padding * 2);
      const y = height - padding - ((record[config.yKey] - yMin) / spanY) * (height - padding * 2);
      const isSelected = String(record.office_no) === selectedOfficeNo;
      const fill =
        record.wage_outlier_flag === "high" || record.wam_staffing_outlier_flag === "high"
          ? "#c86f43"
          : record.wam_match_status === "matched"
            ? "#0f7c79"
            : "#94a2a5";
      return `<circle cx="${x.toFixed(1)}" cy="${y.toFixed(1)}" r="${isSelected ? 5.2 : 3.2}" fill="${fill}" fill-opacity="${
        isSelected ? "1" : "0.72"
      }" stroke="${isSelected ? "#233033" : "transparent"}" stroke-width="${isSelected ? "1.5" : "0"}"></circle>`;
    })
    .join("");

  root.innerHTML = `
    <div class="scatter-wrap">
      <svg viewBox="0 0 ${width} ${height}" role="img" aria-label="${escapeHtml(config.yLabel)} と ${escapeHtml(config.xLabel)} の散布図">
        <line x1="${padding}" y1="${height - padding}" x2="${width - padding}" y2="${height - padding}" class="axis-line"></line>
        <line x1="${padding}" y1="${padding}" x2="${padding}" y2="${height - padding}" class="axis-line"></line>
        ${circles}
      </svg>
      <div class="scatter-meta">
        <span>${escapeHtml(config.xLabel)}: ${escapeHtml(config.xFormatter(xMin))} - ${escapeHtml(config.xFormatter(xMax))}</span>
        <span>${escapeHtml(config.yLabel)}: ${escapeHtml(config.yFormatter(yMin))} - ${escapeHtml(config.yFormatter(yMax))}</span>
      </div>
    </div>
  `;
}

function renderAnomalies(records) {
  const root = document.getElementById("anomalyList");
  const summaryRoot = document.getElementById("anomalySummary");
  const anomalies = records
    .filter((record) => record.wage_outlier_flag || record.wam_staffing_outlier_flag)
    .sort((left, right) => anomalyScore(right) - anomalyScore(left))
    .slice(0, 12);

  summaryRoot.textContent = anomalies.length
    ? `${formatCount(anomalies.length)} 件を抽出`
    : "条件に一致する異常値はない";

  if (!anomalies.length) {
    root.innerHTML = document.getElementById("emptyStateTemplate").innerHTML;
    return;
  }

  root.innerHTML = anomalies
    .map(
      (record) => `
        <button
          class="anomaly-card"
          data-select-office="${escapeHtml(record.office_no)}"
          data-flag="${escapeHtml(record.wage_outlier_flag ?? record.wam_staffing_outlier_flag ?? "none")}"
          type="button"
        >
          <div>
            <p class="section-kicker">${escapeHtml(record.municipality ?? "-")} / No.${escapeHtml(record.office_no ?? "-")}</p>
            <h3>${escapeHtml(record.office_name ?? "-")}</h3>
            <p>${escapeHtml(record.corporation_name ?? "-")}</p>
          </div>
          <div class="anomaly-meta">
            ${record.wage_outlier_flag ? `<span class="metric-chip">${escapeHtml(`工賃 ${labelForSelect(record.wage_outlier_flag)}`)}</span>` : ""}
            ${record.wam_staffing_outlier_flag ? `<span class="metric-chip">${escapeHtml(`人員 ${labelForSelect(record.wam_staffing_outlier_flag)}`)}</span>` : ""}
            <span class="metric-chip">${escapeHtml(formatWageText(record.average_wage_yen))}</span>
            <span class="metric-chip">${escapeHtml(formatPercent(record.wam_key_staff_fte_per_capacity))}</span>
          </div>
          <p>全体平均比 ${escapeHtml(formatRatio(record.wage_ratio_to_overall_mean))} / 主要支援FTE/定員 ${escapeHtml(formatPercent(record.wam_key_staff_fte_per_capacity))}</p>
          <p>${escapeHtml(record.wam_staffing_efficiency_quadrant ?? "人員象限なし")} / ${escapeHtml(record.wam_primary_activity_type ?? "活動不明")}</p>
        </button>
      `
    )
    .join("");
}

function renderDetail(record) {
  const root = document.getElementById("detailContent");
  if (!record) {
    root.innerHTML = document.getElementById("emptyStateTemplate").innerHTML;
    return;
  }

  const actionNotes = buildActionNotes(record);
  const officeUrl = safeExternalUrl(record.wam_office_url);
  root.innerHTML = `
    <div class="detail-hero">
      <div>
        <p class="section-kicker">${escapeHtml(record.municipality ?? "-")} / No.${escapeHtml(record.office_no ?? "-")}</p>
        <h3>${escapeHtml(record.office_name ?? "-")}</h3>
        <p class="detail-subtitle">${escapeHtml(record.corporation_name ?? "-")} / ${escapeHtml(record.corporation_type_label ?? "-")}</p>
      </div>
      <div class="detail-cta">
        ${officeUrl ? `<a class="solid-button link-button" href="${escapeAttribute(officeUrl)}" target="_blank" rel="noreferrer">事業所ページ</a>` : ""}
      </div>
    </div>
    <div class="detail-kpi-grid">
      ${detailKpi("平均工賃", formatWageText(record.average_wage_yen), `全体平均比 ${formatRatio(record.wage_ratio_to_overall_mean)}`)}
      ${detailKpi("稼働率", formatPercent(record.daily_user_capacity_ratio), `定員 ${formatNullable(record.capacity)} 名 / 平均利用 ${formatNumber(record.average_daily_users)} 名`)}
      ${detailKpi(
        "主要支援FTE/定員",
        formatPercent(record.wam_key_staff_fte_per_capacity),
        record.wam_match_status === "matched" ? `人員象限 ${record.wam_staffing_efficiency_quadrant ?? "-"}` : "WAM未一致"
      )}
      ${detailKpi("WAM月額工賃", formatMaybeYen(record.wam_average_wage_monthly_yen), `Excelとの差 ${formatSignedYen(record.wam_average_wage_gap_yen)}`)}
    </div>
    <div class="detail-grid">
      <article class="detail-card">
        <h3>ベンチマーク</h3>
        <ul class="detail-list">
          <li>市町村平均比: ${escapeHtml(formatRatio(record.wage_ratio_to_municipality_mean))}</li>
          <li>法人種別平均比: ${escapeHtml(formatRatio(record.wage_ratio_to_corporation_type_mean))}</li>
          <li>定員帯平均比: ${escapeHtml(formatRatio(record.wage_ratio_to_capacity_band_mean))}</li>
          <li>工賃帯: ${escapeHtml(record.wage_band_label ?? "-")}</li>
          <li>運営ポジション: ${escapeHtml(record.market_position_quadrant ?? "-")}</li>
        </ul>
      </article>
      <article class="detail-card">
        <h3>WAM運営体制</h3>
        <ul class="detail-list">
          <li>WAM一致: ${escapeHtml(record.wam_match_status === "matched" ? `一致 / ${record.wam_match_confidence ?? "-"}` : "未一致")}</li>
          <li>送迎: ${escapeHtml(formatBool(record.wam_transport_available))}</li>
          <li>食事加算: ${escapeHtml(formatBool(record.wam_meal_support_addon))}</li>
          <li>管理者兼務: ${escapeHtml(formatBool(record.wam_manager_multi_post))}</li>
          <li>主活動: ${escapeHtml(record.wam_primary_activity_type ?? "-")}</li>
          <li>サービス管理責任者FTE: ${escapeHtml(formatFte(record.wam_service_manager_fte))}</li>
        </ul>
      </article>
      <article class="detail-card">
        <h3>注視ポイント</h3>
        <ul class="detail-list">
          ${actionNotes.map((note) => `<li>${escapeHtml(note)}</li>`).join("")}
        </ul>
      </article>
      <article class="detail-card">
        <h3>基本情報</h3>
        <ul class="detail-list">
          <li>住所: ${escapeHtml(composeAddress(record))}</li>
          <li>電話: ${escapeHtml(record.wam_office_phone ?? "-")}</li>
          <li>事業所番号: ${escapeHtml(record.wam_office_number ?? "-")}</li>
          <li>WAM定員: ${escapeHtml(formatNullable(record.wam_office_capacity))}</li>
          <li>新設: ${escapeHtml(record.is_new_office ? "あり" : "なし")}</li>
          <li>備考: ${escapeHtml(record.remarks ?? "-")}</li>
        </ul>
      </article>
    </div>
  `;
}

function renderTable(records) {
  const tableBody = document.getElementById("recordsTableBody");
  const pageCount = Math.max(1, Math.ceil(records.length / PAGE_SIZE));
  const startIndex = (state.currentPage - 1) * PAGE_SIZE;
  const pagedRecords = records.slice(startIndex, startIndex + PAGE_SIZE);

  document.getElementById("tableSummary").textContent = `${formatCount(records.length)} 件 / ${formatCount(pageCount)} ページ`;
  document.getElementById("pageSummary").textContent = `${formatCount(state.currentPage)} / ${formatCount(pageCount)} ページ`;
  document.getElementById("prevPageButton").disabled = state.currentPage <= 1;
  document.getElementById("nextPageButton").disabled = state.currentPage >= pageCount;

  if (!pagedRecords.length) {
    tableBody.innerHTML = `<tr><td colspan="19">${document.getElementById("emptyStateTemplate").innerHTML}</td></tr>`;
    return;
  }

  tableBody.innerHTML = pagedRecords
    .map(
      (record) => `
        <tr class="${rowClass(record)}">
          <td class="numeric">${formatNullable(record.office_no)}</td>
          <td>${escapeHtml(record.municipality ?? "-")}</td>
          <td>${escapeHtml(record.corporation_type_label ?? "-")}</td>
          <td>${escapeHtml(record.corporation_name ?? "-")}</td>
          <td>${escapeHtml(record.office_name ?? "-")}</td>
          <td class="numeric">${formatWage(record.average_wage_yen, record.average_wage_error)}</td>
          <td class="numeric">${formatRatio(record.wage_ratio_to_overall_mean)}</td>
          <td>${outlierBadge(record.wage_outlier_flag, record.wage_outlier_severity)}</td>
          <td class="numeric">${formatPercent(record.daily_user_capacity_ratio)}</td>
          <td>${matchBadge(record.wam_match_status, record.wam_match_confidence)}</td>
          <td class="numeric">${formatFte(record.wam_welfare_staff_fte_total)}</td>
          <td class="numeric">${formatPercent(record.wam_key_staff_fte_per_capacity)}</td>
          <td>${booleanFlag(record.wam_transport_available)}</td>
          <td>${quadrantBadge(record.wam_staffing_efficiency_quadrant)}</td>
          <td>${staffingOutlierBadge(record.wam_staffing_outlier_flag, record.wam_staffing_outlier_severity)}</td>
          <td>${statusBadge(record.response_status)}</td>
          <td>${booleanFlag(record.is_new_office)}</td>
          <td>${escapeHtml(record.remarks ?? "-")}</td>
          <td><button class="table-link" data-select-office="${escapeHtml(record.office_no)}" type="button">表示</button></td>
        </tr>
      `
    )
    .join("");
}

function changePage(step) {
  const totalPages = Math.max(1, Math.ceil(state.filteredRecords.length / PAGE_SIZE));
  const nextPage = Math.min(Math.max(state.currentPage + step, 1), totalPages);
  if (nextPage === state.currentPage) return;
  state.currentPage = nextPage;
  renderTable(state.filteredRecords);
}

function rowClass(record) {
  const classes = [];
  if (String(record.office_no) === String(state.selectedOfficeNo)) classes.push("row-selected");
  if (record.wage_outlier_flag === "high" || record.wam_staffing_outlier_flag === "high") classes.push("row-high-outlier");
  if (record.wage_outlier_flag === "low" || record.wam_staffing_outlier_flag === "low") classes.push("row-low-outlier");
  return classes.join(" ");
}

function matchedRecords(records) {
  return records.filter((record) => record.wam_match_status === "matched" && record.wam_fetch_status === "ok");
}

function numericValues(records, key) {
  return records.map((record) => record[key]).filter(isNumber);
}

function meanFor(records, key) {
  const values = numericValues(records, key);
  if (!values.length) return null;
  return values.reduce((sum, value) => sum + value, 0) / values.length;
}

function ratioOf(records, predicate) {
  if (!records.length) return null;
  return records.filter(predicate).length / records.length;
}

function computeLocalStats(values) {
  if (!values.length) return { mean: null, median: null, p90: null };
  const sorted = [...values].sort((left, right) => left - right);
  return {
    mean: sorted.reduce((sum, value) => sum + value, 0) / sorted.length,
    median: median(sorted),
    p90: percentile(sorted, 0.9),
  };
}

function percentile(sortedValues, ratio) {
  if (!sortedValues.length) return null;
  if (sortedValues.length === 1) return sortedValues[0];
  const position = (sortedValues.length - 1) * ratio;
  const lower = Math.floor(position);
  const upper = Math.ceil(position);
  if (lower === upper) return sortedValues[lower];
  return sortedValues[lower] + (sortedValues[upper] - sortedValues[lower]) * (position - lower);
}

function median(values) {
  const sorted = [...values].sort((left, right) => left - right);
  const middle = Math.floor(sorted.length / 2);
  return sorted.length % 2 === 0 ? (sorted[middle - 1] + sorted[middle]) / 2 : sorted[middle];
}

function averageByOrderedGroup(records, groupKey, metricKey, order) {
  return order
    .map((label) => {
      const values = records
        .filter((record) => record[groupKey] === label && isNumber(record[metricKey]))
        .map((record) => record[metricKey]);
      if (!values.length) return null;
      return { label, value: values.reduce((sum, value) => sum + value, 0) / values.length };
    })
    .filter(Boolean);
}

function orderedCounts(records, key, order) {
  const counter = new Map();
  records.forEach((record) => {
    const value = record[key];
    if (!value) return;
    counter.set(value, (counter.get(value) ?? 0) + 1);
  });
  return order.map((label) => ({ label, value: counter.get(label) ?? 0 })).filter((item) => item.value > 0);
}

function topCounts(records, key, limit) {
  const counter = new Map();
  records.forEach((record) => {
    const value = record[key];
    if (!value) return;
    counter.set(value, (counter.get(value) ?? 0) + 1);
  });
  return [...counter.entries()]
    .sort((left, right) => right[1] - left[1])
    .slice(0, limit)
    .map(([label, value]) => ({ label, value }));
}

function topAverageWageByMunicipality(records, limit) {
  const buckets = new Map();
  records.forEach((record) => {
    if (!record.municipality || !isNumber(record.average_wage_yen)) return;
    const bucket = buckets.get(record.municipality) ?? [];
    bucket.push(record.average_wage_yen);
    buckets.set(record.municipality, bucket);
  });
  return [...buckets.entries()]
    .filter(([, values]) => values.length >= 5)
    .map(([label, values]) => ({ label, value: values.reduce((sum, value) => sum + value, 0) / values.length }))
    .sort((left, right) => right.value - left.value)
    .slice(0, limit);
}

function wamCoverageBreakdown(records) {
  const matched = records.filter((record) => record.wam_match_status === "matched").length;
  const unmatched = records.length - matched;
  return [
    { label: "WAM一致", value: matched },
    { label: "WAM未一致", value: unmatched },
  ];
}

function combinedOutlierBreakdown(records) {
  return [
    { label: "工賃高", value: records.filter((record) => record.wage_outlier_flag === "high").length },
    { label: "工賃低", value: records.filter((record) => record.wage_outlier_flag === "low").length },
    { label: "人員厚い", value: records.filter((record) => record.wam_staffing_outlier_flag === "high").length },
    { label: "人員薄い", value: records.filter((record) => record.wam_staffing_outlier_flag === "low").length },
  ].filter((item) => item.value > 0);
}

function staffingRoleAverages(records) {
  const matched = matchedRecords(records);
  const rows = [
    { label: "サビ管", value: meanFor(matched, "wam_service_manager_fte") },
    { label: "就労支援員", value: meanFor(matched, "wam_employment_support_fte") },
    { label: "職業指導員", value: meanFor(matched, "wam_vocational_instructor_fte") },
    { label: "生活支援員", value: meanFor(matched, "wam_life_support_fte") },
  ];
  return rows.filter((row) => isNumber(row.value));
}

function featureWageComparison(records) {
  const matched = matchedRecords(records);
  const rows = [
    {
      label: "送迎あり",
      value: meanFor(matched.filter((record) => record.wam_transport_available === true), "average_wage_yen"),
    },
    {
      label: "食事加算あり",
      value: meanFor(matched.filter((record) => record.wam_meal_support_addon === true), "average_wage_yen"),
    },
    {
      label: "地域協働あり",
      value: meanFor(matched.filter((record) => record.wam_regional_collaboration_addon === true), "average_wage_yen"),
    },
    {
      label: "管理者兼務あり",
      value: meanFor(matched.filter((record) => record.wam_manager_multi_post === true), "average_wage_yen"),
    },
  ];
  return rows.filter((row) => isNumber(row.value));
}

function anomalyScore(record) {
  const wageScore = Math.abs(record.wage_z_score ?? 0) * 5 + (record.wage_ratio_to_overall_mean ?? 0) * 2;
  const staffingBase = record.wam_key_staff_fte_per_capacity ?? 0;
  const staffingScore =
    record.wam_staffing_outlier_flag === "high" ? staffingBase * 40 : record.wam_staffing_outlier_flag === "low" ? 30 : 0;
  return wageScore + staffingScore;
}

function growthScore(record) {
  return (
    (record.wage_ratio_to_overall_mean ?? 0) * 70 +
    (1 - (record.daily_user_capacity_ratio ?? 1)) * 25 +
    (record.wam_staffing_efficiency_quadrant === "高工賃 × 少ない人員" ? 12 : 0)
  );
}

function fixScore(record) {
  return (
    Math.max(1 - (record.wage_ratio_to_overall_mean ?? 1), 0) * 80 +
    ((record.wam_key_staff_fte_per_capacity ?? 0) * 100) +
    Math.max(0.7 - (record.daily_user_capacity_ratio ?? 0.7), 0) * 25
  );
}

function reviewScore(record) {
  let score = 0;
  if (record.response_status === "unanswered") score += 50;
  if (record.wam_match_status !== "matched") score += 30;
  if (record.is_new_office) score += 20;
  if (record.average_wage_error) score += 15;
  return score;
}

function buildGrowthReason(record) {
  return `工賃 ${formatWageText(record.average_wage_yen)} / 稼働率 ${formatPercent(
    record.daily_user_capacity_ratio
  )} / 主要支援FTE ${formatPercent(record.wam_key_staff_fte_per_capacity)}`;
}

function buildFixReason(record) {
  return `平均比 ${formatRatio(record.wage_ratio_to_overall_mean)} / 稼働率 ${formatPercent(
    record.daily_user_capacity_ratio
  )} / 人員象限 ${record.wam_staffing_efficiency_quadrant ?? "-"}`;
}

function buildReviewReason(record) {
  const flags = [];
  if (record.response_status === "unanswered") flags.push("未回答");
  if (record.wam_match_status !== "matched") flags.push("WAM未一致");
  if (record.is_new_office) flags.push("新設");
  if (record.average_wage_error) flags.push(record.average_wage_error);
  return flags.join(" / ") || "確認推奨";
}

function buildActionNotes(record) {
  const notes = [];
  if (record.wam_match_status !== "matched") {
    notes.push("WAM未一致で人員分析が欠ける。名称揺れか事業所追加有無を確認したい。");
  }
  if ((record.wage_ratio_to_municipality_mean ?? 1) < 0.9) {
    notes.push(`市町村平均比 ${formatRatio(record.wage_ratio_to_municipality_mean)}。地域内の工賃競争力を点検したい。`);
  }
  if (isNumber(record.daily_user_capacity_ratio) && record.daily_user_capacity_ratio < 0.6) {
    notes.push(`稼働率 ${formatPercent(record.daily_user_capacity_ratio)}。体験導線、紹介元、欠席率を見直す余地が大きい。`);
  }
  if (record.wam_staffing_efficiency_quadrant === "低工賃 × 厚い人員") {
    notes.push("人員が厚い割に工賃が伸びていない。役割配置と作業単価の両面を再点検したい。");
  }
  if (record.wage_outlier_flag === "high") {
    notes.push("高工賃の上振れ事業所である。主活動と取引構造を横展開候補として見る価値がある。");
  }
  if (record.is_new_office) {
    notes.push("新設事業所である。立ち上がり期なので単月値ではなく定員充足と工賃立ち上がりを追いたい。");
  }
  if (!notes.length) {
    notes.push("大きな警戒シグナルは薄い。市平均比、稼働率、支援体制のバランスを定点観測したい。");
  }
  return notes.slice(0, 4);
}

function detailKpi(label, value, hint) {
  return `
    <article class="detail-kpi">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(value)}</strong>
      <em>${escapeHtml(hint)}</em>
    </article>
  `;
}

function composeAddress(record) {
  const parts = [record.wam_office_address_city, record.wam_office_address_line].filter(Boolean);
  return parts.length ? parts.join("") : "-";
}

function downloadFilteredCsv() {
  const headers = [
    ["office_no", "No."],
    ["municipality", "市町村"],
    ["corporation_type_label", "法人種別"],
    ["corporation_name", "法人名"],
    ["office_name", "事業所名"],
    ["average_wage_yen", "平均工賃"],
    ["wage_ratio_to_overall_mean", "全体平均比"],
    ["wage_ratio_to_municipality_mean", "市町村平均比"],
    ["wage_ratio_to_capacity_band_mean", "定員帯平均比"],
    ["daily_user_capacity_ratio", "稼働率"],
    ["wage_outlier_flag", "工賃異常値"],
    ["wam_match_status", "WAM状態"],
    ["wam_welfare_staff_fte_total", "福祉職員FTE"],
    ["wam_key_staff_fte_per_capacity", "主要支援FTE/定員"],
    ["wam_transport_available", "送迎"],
    ["wam_meal_support_addon", "食事加算"],
    ["wam_manager_multi_post", "管理者兼務"],
    ["wam_staffing_efficiency_quadrant", "人員象限"],
    ["wam_staffing_outlier_flag", "人員異常値"],
    ["wam_primary_activity_type", "主活動種別"],
    ["wam_office_number", "WAM事業所番号"],
    ["wam_office_url", "事業所URL"],
    ["remarks", "備考"],
  ];

  const lines = [headers.map(([, label]) => escapeCsvCell(label)).join(",")];
  state.filteredRecords.forEach((record) => {
    const row = headers.map(([key]) => {
      const value = record[key];
      if (
        key === "response_status" ||
        key === "wage_outlier_flag" ||
        key === "wam_staffing_outlier_flag"
      ) {
        return escapeCsvCell(labelForSelect(value ?? "none"));
      }
      if (key === "wam_match_status") {
        return escapeCsvCell(labelForSelect(value ?? "unmatched"));
      }
      if (typeof value === "boolean") {
        return escapeCsvCell(value ? "あり" : "なし");
      }
      return escapeCsvCell(value ?? "");
    });
    lines.push(row.join(","));
  });

  const blob = new Blob([`\uFEFF${lines.join("\n")}`], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "r6kouchin-dashboard-filtered.csv";
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function statusBadge(status) {
  const label = labelForSelect(status ?? "");
  const variant = status ?? "answered";
  return `<span class="badge badge-${escapeHtml(variant)}">${escapeHtml(label)}</span>`;
}

function outlierBadge(flag, severity) {
  if (!flag) return `<span class="badge badge-answered">通常範囲</span>`;
  const variant = flag === "high" ? "unanswered" : "answered";
  const label = `工賃 ${labelForSelect(flag)}${severity ? ` / ${severity}` : ""}`;
  return `<span class="badge badge-${variant}">${escapeHtml(label)}</span>`;
}

function staffingOutlierBadge(flag, severity) {
  if (!flag) return `<span class="badge badge-answered">通常範囲</span>`;
  const variant = flag === "high" ? "annotated" : "answered";
  const label = `人員 ${labelForSelect(flag)}${severity ? ` / ${severity}` : ""}`;
  return `<span class="badge badge-${variant}">${escapeHtml(label)}</span>`;
}

function matchBadge(status, confidence) {
  if (status !== "matched") return `<span class="badge badge-unanswered">未一致</span>`;
  return `<span class="badge badge-annotated">一致${confidence ? ` / ${escapeHtml(confidence)}` : ""}</span>`;
}

function quadrantBadge(value) {
  if (!value) return `<span class="flag-muted">-</span>`;
  return `<span class="metric-chip">${escapeHtml(value)}</span>`;
}

function booleanFlag(value) {
  if (value === true) return `<span class="flag">○</span>`;
  if (value === false) return `<span class="flag-muted">-</span>`;
  return `<span class="flag-muted">?</span>`;
}

function formatBool(value) {
  if (value === true) return "あり";
  if (value === false) return "なし";
  return "不明";
}

function formatWage(value, error) {
  if (error) return `<span class="flag-muted">${escapeHtml(error)}</span>`;
  return isNumber(value) ? `${formatCount(Math.round(value))}円` : "-";
}

function formatWageText(value) {
  return isNumber(value) ? `${formatCount(Math.round(value))}円` : "-";
}

function formatFte(value) {
  return isNumber(value) ? `${decimalFormatter.format(value)} FTE` : "-";
}

function formatNumber(value) {
  return isNumber(value) ? ratioFormatter.format(value) : "-";
}

function formatRatio(value) {
  return isNumber(value) ? `${ratioFormatter.format(value)}x` : "-";
}

function formatPercent(value) {
  return isNumber(value) ? `${percentFormatter.format(value * 100)}%` : "-";
}

function formatMaybeYen(value) {
  return isNumber(value) ? `${formatCount(Math.round(value))}円` : "-";
}

function formatSignedYen(value) {
  if (!isNumber(value)) return "-";
  return `${value >= 0 ? "+" : ""}${formatCount(Math.round(value))}円`;
}

function formatCount(value) {
  return numberFormatter.format(value);
}

function formatNullable(value) {
  return value == null || value === "" ? "-" : escapeHtml(String(value));
}

function formatCountSuffix(unit, roundValue = false) {
  return (value) => {
    if (!isNumber(value)) return "-";
    if (unit === "件" || roundValue) {
      return `${formatCount(Math.round(value))} ${unit}`;
    }
    return `${ratioFormatter.format(value)} ${unit}`;
  };
}

function isNumber(value) {
  return typeof value === "number" && Number.isFinite(value);
}

function renderError(error) {
  document.body.innerHTML = `
    <main class="page-shell">
      <section class="panel">
        <div class="empty-state">
          <h3>ダッシュボードを読み込めない</h3>
          <p>${escapeHtml(error?.message ?? "unknown error")}</p>
        </div>
      </section>
    </main>
  `;
}

function safeExternalUrl(value) {
  if (!value) return null;
  try {
    const url = new URL(String(value), window.location.href);
    return ["http:", "https:"].includes(url.protocol) ? url.href : null;
  } catch {
    return null;
  }
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function escapeAttribute(value) {
  return escapeHtml(value).replaceAll("`", "&#96;");
}

function escapeCsvCell(value) {
  const text = String(value ?? "");
  if (/[",\n]/.test(text)) {
    return `"${text.replaceAll('"', '""')}"`;
  }
  return text;
}
