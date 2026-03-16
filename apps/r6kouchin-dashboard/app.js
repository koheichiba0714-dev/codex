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
const CORPORATION_SORT_OPTIONS = [
  { key: "municipality", label: "市区順" },
  { key: "wage", label: "工賃順" },
  { key: "utilization", label: "利用率順" },
];

/* ─────────────────────────────────────────────
   報酬算定区分（令和6年度 就労継続支援B型）
   工賃月額に応じて9区分。区分が上がるほど
   事業所が受け取る報酬単価（1日あたり）が高い。
   ───────────────────────────────────────────── */
const WAGE_TIER_TABLE = [
  { min: 0,     max: 5000,       label: "区分1", tierNo: 1, unitYen: 567 },
  { min: 5000,  max: 10000,      label: "区分2", tierNo: 2, unitYen: 590 },
  { min: 10000, max: 15000,      label: "区分3", tierNo: 3, unitYen: 612 },
  { min: 15000, max: 20000,      label: "区分4", tierNo: 4, unitYen: 646 },
  { min: 20000, max: 25000,      label: "区分5", tierNo: 5, unitYen: 671 },
  { min: 25000, max: 30000,      label: "区分6", tierNo: 6, unitYen: 703 },
  { min: 30000, max: 35000,      label: "区分7", tierNo: 7, unitYen: 736 },
  { min: 35000, max: 45000,      label: "区分8", tierNo: 8, unitYen: 757 },
  { min: 45000, max: Infinity,   label: "区分9", tierNo: 9, unitYen: 764 },
];

/* 法定人員配置基準（令和6年度改定）
   職業指導員＋生活支援員の合計 ÷ 利用者数
   6:1  → サービス費(Ⅰ)/(Ⅳ) — R6新設・最も手厚い
   7.5:1 → サービス費(Ⅱ)/(Ⅴ) — 標準
   10:1  → 最低基準 */
const STAFFING_TIERS = [
  { ratio: 6,   label: "6:1（手厚い）", serviceFee: "Ⅰ/Ⅳ" },
  { ratio: 7.5, label: "7.5:1（標準）", serviceFee: "Ⅱ/Ⅴ" },
  { ratio: 10,  label: "10:1（最低基準）", serviceFee: "-" },
];
const LEGAL_STAFFING_RATIO_MIN = 10;   // 最低基準
const LEGAL_STAFFING_RATIO_STD = 7.5;  // 標準
const LEGAL_STAFFING_RATIO_HIGH = 6;   // R6新設・手厚い

/* 大阪府の工賃向上目標（令和6年度） */
const OSAKA_WAGE_TARGET_YEN = 15000;

/* ─────────────────────────────────────────────
   解析関数群（データサイエンス＋制度知識）
   ───────────────────────────────────────────── */

/** 報酬算定区分を返す */
function getWageTier(wageYen) {
  if (!isNumber(wageYen)) return null;
  return WAGE_TIER_TABLE.find((t) => wageYen >= t.min && wageYen < t.max) ?? WAGE_TIER_TABLE[0];
}

/** 次の区分までの距離と収入インパクトを返す */
function getWageTierUpInfo(record) {
  const wage = record.average_wage_yen;
  const tier = getWageTier(wage);
  if (!tier || !isNumber(wage)) return null;
  const nextTier = WAGE_TIER_TABLE.find((t) => t.tierNo === tier.tierNo + 1);
  if (!nextTier) return { current: tier, next: null, gapYen: 0, revenueImpact: 0 };
  const gapYen = Math.max(0, nextTier.min - wage);
  const avgDailyUsers = record.average_daily_users ?? 0;
  const unitGain = nextTier.unitYen - tier.unitYen;
  // 月の稼働日約22日 × 利用者数 × 単価差
  const monthlyRevenueImpact = Math.round(unitGain * avgDailyUsers * 22);
  return { current: tier, next: nextTier, gapYen: Math.round(gapYen), revenueImpact: monthlyRevenueImpact, unitGain };
}

/** ソート済み配列に対してパーセンタイル順位 (0-100) を返す */
function computePercentileRank(value, sortedValues) {
  if (!isNumber(value) || !sortedValues.length) return null;
  let count = 0;
  for (const v of sortedValues) {
    if (v < value) count++;
    else break;
  }
  return Math.round((count / sortedValues.length) * 100);
}

/** 人員配置基準の充足度を返す（R6年度3段階対応） */
function getStaffingComplianceLevel(record) {
  if (record.wam_match_status !== "matched") return null;
  const staffFte = record.wam_welfare_staff_fte_total;
  const capacity = record.wam_office_capacity ?? record.capacity;
  if (!isNumber(staffFte) || !isNumber(capacity) || capacity <= 0) return null;
  const requiredMin = capacity / LEGAL_STAFFING_RATIO_MIN;   // 10:1
  const requiredStd = capacity / LEGAL_STAFFING_RATIO_STD;   // 7.5:1
  const requiredHigh = capacity / LEGAL_STAFFING_RATIO_HIGH;  // 6:1
  const ratioToMin = staffFte / requiredMin;
  let level, label, qualifiedTier, advice;
  if (staffFte >= requiredHigh) {
    level = "tier_6_1"; label = "6:1達成";
    qualifiedTier = "サービス費(Ⅰ)/(Ⅳ)";
    advice = "令和6年度新設の最も手厚い6:1基準を達成。最高報酬区分の対象。工賃向上の成果が問われる配置水準。";
  } else if (staffFte >= requiredStd) {
    level = "tier_7_5_1"; label = "7.5:1達成";
    qualifiedTier = "サービス費(Ⅱ)/(Ⅴ)";
    const gap = requiredHigh - staffFte;
    advice = `標準の7.5:1をクリア。あと常勤換算 ${ratioFormatter.format(Math.max(gap, 0))} 人で6:1に到達し報酬単価アップ。`;
  } else if (staffFte >= requiredMin) {
    level = "tier_10_1"; label = "10:1達成";
    qualifiedTier = "最低基準";
    const gap = requiredStd - staffFte;
    advice = `最低基準の10:1はクリアだが余裕がない。あと常勤換算 ${ratioFormatter.format(Math.max(gap, 0))} 人で7.5:1に到達。`;
  } else {
    level = "critical"; label = "基準未達の可能性";
    qualifiedTier = "要確認";
    advice = "10:1の最低基準を満たしていない可能性がある。指定取消リスクあり、至急人員補充を検討すべき。";
  }
  return { ratioToMin, requiredMin, requiredStd, requiredHigh, staffFte, level, label, qualifiedTier, advice };
}

/** ピアベンチマーク: 同じ主活動 or 同じ定員帯の事業所と比較 */
function computePeerBenchmark(record, allRecords) {
  const result = {};
  // 主活動種別ピア
  const actType = record.wam_primary_activity_type;
  if (actType) {
    const peers = allRecords.filter(
      (r) => r.wam_primary_activity_type === actType && isNumber(r.average_wage_yen)
    );
    if (peers.length >= 3) {
      const wages = peers.map((r) => r.average_wage_yen).sort((a, b) => a - b);
      const rank = peers.filter((r) => r.average_wage_yen > record.average_wage_yen).length + 1;
      result.activity = {
        type: actType,
        count: peers.length,
        rank,
        median: median(wages),
        mean: wages.reduce((s, v) => s + v, 0) / wages.length,
        myWage: record.average_wage_yen,
      };
    }
  }
  // 定員帯ピア
  const capBand = record.capacity_band_label;
  if (capBand) {
    const peers = allRecords.filter(
      (r) => r.capacity_band_label === capBand && isNumber(r.average_wage_yen)
    );
    if (peers.length >= 3) {
      const wages = peers.map((r) => r.average_wage_yen).sort((a, b) => a - b);
      const rank = peers.filter((r) => r.average_wage_yen > record.average_wage_yen).length + 1;
      result.capacity = {
        band: capBand,
        count: peers.length,
        rank,
        median: median(wages),
        mean: wages.reduce((s, v) => s + v, 0) / wages.length,
      };
    }
  }
  return result;
}

/** 利用者1人あたり支援職員数 */
function staffPerActualUser(record) {
  const staff = record.wam_welfare_staff_fte_total;
  const users = record.average_daily_users;
  if (!isNumber(staff) || !isNumber(users) || users <= 0) return null;
  return staff / users;
}

/** 利用率1%改善あたりの月間収入増加額（経営者向け） */
function revenuePerUtilizationPoint(record) {
  const tier = getWageTier(record.average_wage_yen);
  const capacity = record.capacity ?? record.wam_office_capacity;
  if (!tier || !isNumber(capacity) || capacity <= 0) return null;
  // 1%改善 = 定員 × 0.01 人分の利用者増 × 単価 × 22日
  return Math.round(capacity * 0.01 * tier.unitYen * 22);
}

/** 加算取り漏れチェック（経営者向け） */
function checkMissedAddons(record) {
  const missed = [];
  if (record.wam_match_status !== "matched") return missed;

  // 目標工賃達成指導員配置加算: 工賃が前年度より上がっている or 一定額以上の場合に取得できる
  const tier = getWageTier(record.average_wage_yen);
  if (tier && tier.tierNo >= 4 && record.wam_staffing_efficiency_quadrant !== "低工賃 × 少ない人員") {
    missed.push({
      name: "目標工賃達成指導員配置加算",
      hint: `工賃が${formatWageText(record.average_wage_yen)}と比較的高く、人員も確保されている。目標工賃達成指導員の配置で月額約7万円/人の加算が見込める。`,
    });
  }

  // 送迎加算: 送迎なしだが利用率が低い場合
  if (record.wam_transport_available === false && isNumber(record.daily_user_capacity_ratio) && record.daily_user_capacity_ratio < 0.7) {
    missed.push({
      name: "送迎体制の整備",
      hint: "送迎なしかつ利用率低め。送迎実施で利用率改善が期待できる（送迎加算: 片道21単位/日）。",
    });
  }

  // 食事加算: なしの場合
  if (record.wam_meal_support_addon === false) {
    missed.push({
      name: "食事提供体制加算",
      hint: "食事提供なし。食事加算（30単位/日）は利用者満足度と定着率の向上にもつながる。",
    });
  }

  // 在宅利用: あるが在宅率が低い場合
  if (record.home_use_active === true && isNumber(record.home_use_user_ratio_decimal) && record.home_use_user_ratio_decimal < 0.05) {
    missed.push({
      name: "在宅利用支援の強化",
      hint: "在宅利用ありだが実績が少ない。在宅時生活支援加算（15単位/日）を算定できるケースを増やす余地あり。",
    });
  }

  return missed;
}

/** 好事例の主活動詳細を抽出（スタッフ向け） */
function topPerformerActivities(records, limit = 5) {
  return records
    .filter((r) => isNumber(r.average_wage_yen) && r.wam_primary_activity_detail && r.wage_outlier_flag === "high")
    .sort((a, b) => (b.average_wage_yen ?? 0) - (a.average_wage_yen ?? 0))
    .slice(0, limit)
    .map((r) => ({
      name: r.office_name,
      wage: r.average_wage_yen,
      activity: r.wam_primary_activity_type,
      detail: r.wam_primary_activity_detail,
      municipality: r.municipality,
    }));
}

/** 工賃の標準偏差を計算 */
function computeStdDev(values) {
  if (values.length < 2) return null;
  const mean = values.reduce((s, v) => s + v, 0) / values.length;
  const variance = values.reduce((s, v) => s + (v - mean) ** 2, 0) / (values.length - 1);
  return Math.sqrt(variance);
}

/** ジニ係数（工賃の格差指標） */
function computeGini(values) {
  if (values.length < 2) return null;
  const sorted = [...values].sort((a, b) => a - b);
  const n = sorted.length;
  const mean = sorted.reduce((s, v) => s + v, 0) / n;
  if (mean === 0) return 0;
  let sumDiff = 0;
  for (let i = 0; i < n; i++) {
    sumDiff += (2 * (i + 1) - n - 1) * sorted[i];
  }
  return sumDiff / (n * n * mean);
}

/** 大阪府目標(15,000円)達成率 */
function osakaTargetRate(records) {
  const withWage = records.filter((r) => isNumber(r.average_wage_yen));
  if (!withWage.length) return null;
  return withWage.filter((r) => r.average_wage_yen >= OSAKA_WAGE_TARGET_YEN).length / withWage.length;
}

/** 報酬算定区分の分布 */
function wageTierDistribution(records) {
  const dist = new Map();
  WAGE_TIER_TABLE.forEach((t) => dist.set(t.label, 0));
  records.forEach((r) => {
    const tier = getWageTier(r.average_wage_yen);
    if (tier) dist.set(tier.label, (dist.get(tier.label) ?? 0) + 1);
  });
  return WAGE_TIER_TABLE.map((t) => ({ label: t.label, value: dist.get(t.label) ?? 0 })).filter((item) => item.value > 0);
}
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
  "high-wage-low-util": {
    quadrant: "高工賃 × 低稼働",
  },
  "fix-priority": {
    wamOnly: true,
    quadrant: "低工賃 × 低稼働",
    staffingOutlier: "high",
  },
  "high-wage-high-util": {
    quadrant: "高工賃 × 高稼働",
  },
  "work-shortage": {
    workShortageRisk: "likely",
  },
  unanswered: {
    responseStatus: "unanswered",
  },
};
const INITIAL_PRESET = "osaka-city";

function createDefaultFilters() {
  return {
    search: "",
    municipality: "all",
    area: "all",
    corporationType: "all",
    responseStatus: "all",
    outlierFlag: "all",
    capacityBand: "all",
    quadrant: "all",
    wamMatch: "all",
    workShortageRisk: "all",
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
}

function cloneFilters(filters) {
  return { ...filters };
}

function createPresetFilters(preset) {
  return { ...createDefaultFilters(), ...(FILTER_PRESETS[preset] ?? {}) };
}

const state = {
  records: [],
  filteredRecords: [],
  dashboard: null,
  sortKey: "wage_ratio_to_overall_mean",
  sortDirection: "desc",
  currentPage: 1,
  filters: createPresetFilters(INITIAL_PRESET),
  draftFilters: createPresetFilters(INITIAL_PRESET),
  selectedOfficeNo: "339",
  selectedCorporationKey: null,
  selectedCorporationLabel: null,
  selectedCorporationSort: "municipality",
  activePreset: INITIAL_PRESET,
};

document.addEventListener("DOMContentLoaded", () => {
  init().catch((error) => {
    renderError(error);
  });
});

async function init() {
  bindSectionNav();
  bindFilterDialog();
  bindGuideDialog();
  bindDetailDialog();
  bindCorporationDialog();
  bindPanelToggles();
  bindMobileSidebar();
  const response = await fetch("./data/dashboard-data.json", { cache: "no-store" });
  if (!response.ok) {
    throw new Error("dashboard data could not be loaded");
  }

  const dashboard = await response.json();
  state.dashboard = dashboard;
  renderLoadingState();
  state.records = await loadDashboardRecords(dashboard);
  renderMeta(dashboard);
  renderQuality(dashboard, state.records);
  populateFilterOptions(state.records);
  bindEvents();
  syncPresetButtons();
  applyFilters();
}

function bindSectionNav() {
  const links = Array.from(document.querySelectorAll(".sidebar .nav-link[href^='#']"));
  if (!links.length) return;

  const sections = links
    .map((link) => {
      const target = document.querySelector(link.getAttribute("href"));
      return target ? { id: target.id, link, target } : null;
    })
    .filter(Boolean);

  if (!sections.length) return;

  const activate = (id) => {
    sections.forEach(({ id: sectionId, link }) => {
      link.classList.toggle("is-active", sectionId === id);
    });
  };

  links.forEach((link) => {
    link.addEventListener("click", () => {
      const targetId = link.getAttribute("href")?.slice(1);
      if (!targetId) return;
      const target = document.getElementById(targetId);
      target?.closest(".panel.is-collapsed")?.classList.remove("is-collapsed");
      syncPanelToggleButtons();
      activate(targetId);
    });
  });

  if (!("IntersectionObserver" in window)) return;

  const observer = new IntersectionObserver(
    (entries) => {
      const visible = entries
        .filter((entry) => entry.isIntersecting)
        .sort((left, right) => left.boundingClientRect.top - right.boundingClientRect.top)[0];
      if (visible) {
        activate(visible.target.id);
      }
    },
    {
      rootMargin: "-18% 0px -66% 0px",
      threshold: [0, 0.3, 0.8],
    }
  );

  sections.forEach(({ target }) => observer.observe(target));
}

function bindFilterDialog() {
  const dialog = document.getElementById("filtersDialog");
  if (!dialog) return;

  ["openFiltersButton", "openFiltersPanelButton"].forEach((id) => {
    document.getElementById(id)?.addEventListener("click", () => {
      openFiltersDialog(dialog);
    });
  });

  document.getElementById("closeFiltersButton")?.addEventListener("click", () => {
    closeFiltersDialog(dialog);
  });

  document.getElementById("resetFiltersDialogButton")?.addEventListener("click", () => {
    resetDraftFilters();
  });

  document.getElementById("saveFiltersButton")?.addEventListener("click", () => {
    saveDraftFilters(dialog);
  });

  dialog.addEventListener("click", (event) => {
    if (event.target === dialog) {
      closeFiltersDialog(dialog);
    }
  });
}

function openFiltersDialog(dialog = document.getElementById("filtersDialog")) {
  if (!dialog) return;
  state.draftFilters = cloneFilters(state.filters);
  syncFilterControls(state.draftFilters);
  if (typeof dialog.showModal === "function") {
    if (!dialog.open) {
      dialog.showModal();
    }
  } else {
    dialog.setAttribute("open", "open");
  }
  window.requestAnimationFrame(() => {
    document.getElementById("searchInput")?.focus();
  });
}

function closeFiltersDialog(dialog = document.getElementById("filtersDialog")) {
  if (!dialog) return;
  if (typeof dialog.close === "function" && dialog.open) {
    dialog.close();
    return;
  }
  dialog.removeAttribute("open");
}

function bindGuideDialog() {
  const dialog = document.getElementById("guideDialog");
  if (!dialog) return;

  ["openGuideButton", "openGuidePanelButton"].forEach((id) => {
    document.getElementById(id)?.addEventListener("click", () => {
      openDialogElement(dialog);
    });
  });

  document.getElementById("closeGuideButton")?.addEventListener("click", () => {
    closeDialogElement(dialog);
  });

  dialog.addEventListener("click", (event) => {
    if (event.target === dialog) {
      closeDialogElement(dialog);
    }
  });
}

function bindDetailDialog() {
  const dialog = document.getElementById("detailDialog");
  if (!dialog) return;

  document.getElementById("openSelectedDetailButton")?.addEventListener("click", () => {
    openDetailDialog();
  });

  document.getElementById("closeDetailButton")?.addEventListener("click", () => {
    closeDialogElement(dialog);
  });

  dialog.addEventListener("click", (event) => {
    if (event.target === dialog) {
      closeDialogElement(dialog);
    }
  });
}

function bindCorporationDialog() {
  const dialog = document.getElementById("corporationDialog");
  if (!dialog) return;

  document.getElementById("closeCorporationButton")?.addEventListener("click", () => {
    closeDialogElement(dialog);
  });

  dialog.addEventListener("click", (event) => {
    if (event.target === dialog) {
      closeDialogElement(dialog);
    }
  });
}

function bindPanelToggles() {
  document.querySelectorAll(".panel[data-collapsible]").forEach((panel) => {
    panel.querySelector(".panel-collapse-button")?.addEventListener("click", () => {
      panel.classList.toggle("is-collapsed");
      syncPanelToggleButtons();
    });
  });
  syncPanelToggleButtons();
}

function syncPanelToggleButtons() {
  document.querySelectorAll(".panel[data-collapsible]").forEach((panel) => {
    const button = panel.querySelector(".panel-collapse-button");
    if (!button) return;
    const expanded = !panel.classList.contains("is-collapsed");
    button.textContent = expanded ? "閉じる" : "開く";
    button.setAttribute("aria-expanded", expanded ? "true" : "false");
  });
}

function bindMobileSidebar() {
  const sidebar = document.getElementById("sidebar");
  const toggle = document.getElementById("mobileMenuToggle");
  const overlay = document.getElementById("sidebarOverlay");
  if (!sidebar || !toggle) return;

  const closeSidebar = () => { sidebar.classList.remove("is-open"); };
  toggle.addEventListener("click", () => { sidebar.classList.toggle("is-open"); });
  if (overlay) overlay.addEventListener("click", closeSidebar);

  document.querySelectorAll(".sidebar .nav-link").forEach((link) => {
    link.addEventListener("click", closeSidebar);
  });
}

function openDialogElement(dialog) {
  if (!dialog) return;
  if (typeof dialog.showModal === "function") {
    if (!dialog.open) {
      dialog.showModal();
    }
  } else {
    dialog.setAttribute("open", "open");
  }
}

function closeDialogElement(dialog) {
  if (!dialog) return;
  if (typeof dialog.close === "function" && dialog.open) {
    dialog.close();
    return;
  }
  dialog.removeAttribute("open");
}

function openDetailDialog() {
  const record = getSelectedRecord();
  if (!record) return;
  renderDetail(record);
  openDialogElement(document.getElementById("detailDialog"));
}

function openCorporationDialog(corporationName) {
  const label = String(corporationName ?? "").trim();
  if (!label) return;
  renderCorporationDialog(label);
  closeDialogElement(document.getElementById("detailDialog"));
  openDialogElement(document.getElementById("corporationDialog"));
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
  bindInput("searchInput", "input", (filters, value) => {
    filters.search = value.trim();
  });
  bindInput("municipalitySelect", "change", (filters, value) => {
    filters.municipality = value;
  });
  bindInput("areaSelect", "change", (filters, value) => {
    filters.area = value;
  });
  bindInput("corporationTypeSelect", "change", (filters, value) => {
    filters.corporationType = value;
  });
  bindInput("responseStatusSelect", "change", (filters, value) => {
    filters.responseStatus = value;
  });
  bindInput("outlierFlagSelect", "change", (filters, value) => {
    filters.outlierFlag = value;
  });
  bindInput("capacityBandSelect", "change", (filters, value) => {
    filters.capacityBand = value;
  });
  bindInput("quadrantSelect", "change", (filters, value) => {
    filters.quadrant = value;
  });
  bindInput("wamMatchSelect", "change", (filters, value) => {
    filters.wamMatch = value;
  });
  bindInput("workShortageRiskSelect", "change", (filters, value) => {
    filters.workShortageRisk = value;
  });
  bindInput("transportSelect", "change", (filters, value) => {
    filters.transport = value;
  });
  bindInput("staffingOutlierSelect", "change", (filters, value) => {
    filters.staffingOutlier = value;
  });
  bindInput("primaryActivitySelect", "change", (filters, value) => {
    filters.primaryActivity = value;
  });
  bindInput("mealSupportSelect", "change", (filters, value) => {
    filters.mealSupport = value;
  });
  bindInput("managerMultiSelect", "change", (filters, value) => {
    filters.managerMulti = value;
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
    selectRecord(selectedRecord.office_no, { openDetail: false });
    document.getElementById("tableHeading")?.scrollIntoView({ behavior: "smooth", block: "start" });
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
    const corporationSortTrigger = event.target.closest("[data-corporation-sort]");
    if (corporationSortTrigger) {
      const sortKey = corporationSortTrigger.getAttribute("data-corporation-sort");
      if (sortKey && state.selectedCorporationLabel) {
        state.selectedCorporationSort = sortKey;
        renderCorporationDialog(state.selectedCorporationLabel);
      }
      return;
    }

    const corporationTrigger = event.target.closest("[data-open-corporation]");
    if (corporationTrigger) {
      const corporationName = corporationTrigger.getAttribute("data-open-corporation");
      if (!corporationName) return;
      openCorporationDialog(corporationName);
      return;
    }

    const trigger = event.target.closest("[data-select-office]");
    if (!trigger) return;
    const officeNo = trigger.getAttribute("data-select-office");
    if (!officeNo) return;
    if (trigger.closest("#corporationDialog")) {
      closeDialogElement(document.getElementById("corporationDialog"));
    }
    selectRecord(officeNo, { openDetail: true });
  });
}

function bindInput(id, eventName, setter) {
  document.getElementById(id).addEventListener(eventName, (event) => {
    const draftFilters = state.draftFilters ?? state.filters;
    setter(draftFilters, event.target.value);
  });
}

function bindCheckbox(id, key) {
  document.getElementById(id).addEventListener("change", (event) => {
    const draftFilters = state.draftFilters ?? state.filters;
    draftFilters[key] = event.target.checked;
  });
}

function resetDraftFilters() {
  state.draftFilters = createPresetFilters(INITIAL_PRESET);
  syncFilterControls(state.draftFilters);
}

function resetFilters() {
  state.filters = createPresetFilters(INITIAL_PRESET);
  state.draftFilters = cloneFilters(state.filters);
  state.activePreset = INITIAL_PRESET;
  syncFilterControls(state.draftFilters);
  syncPresetButtons();
  state.currentPage = 1;
}

function saveDraftFilters(dialog = document.getElementById("filtersDialog")) {
  state.filters = cloneFilters(state.draftFilters ?? createDefaultFilters());
  state.activePreset = detectActivePreset(state.filters);
  syncPresetButtons();
  state.currentPage = 1;
  applyFilters();
  closeFiltersDialog(dialog);
}

function detectActivePreset(filters) {
  const defaults = createDefaultFilters();
  if (areFiltersEqual(filters, defaults)) {
    return "all";
  }
  return (
    Object.entries(FILTER_PRESETS).find(([, presetFilters]) =>
      areFiltersEqual(filters, { ...defaults, ...presetFilters })
    )?.[0] ?? null
  );
}

function areFiltersEqual(left, right) {
  return Object.keys(createDefaultFilters()).every((key) => left[key] === right[key]);
}

function applyPreset(preset) {
  state.filters = { ...createDefaultFilters(), ...(FILTER_PRESETS[preset] ?? {}) };
  state.draftFilters = cloneFilters(state.filters);
  state.activePreset = preset;
  syncFilterControls(state.draftFilters);
  syncPresetButtons();
  state.currentPage = 1;
  applyFilters();
}

function syncFilterControls(filters = state.filters) {
  document.getElementById("searchInput").value = filters.search;
  document.getElementById("municipalitySelect").value = filters.municipality;
  document.getElementById("areaSelect").value = filters.area;
  document.getElementById("corporationTypeSelect").value = filters.corporationType;
  document.getElementById("responseStatusSelect").value = filters.responseStatus;
  document.getElementById("outlierFlagSelect").value = filters.outlierFlag;
  document.getElementById("capacityBandSelect").value = filters.capacityBand;
  document.getElementById("quadrantSelect").value = filters.quadrant;
  document.getElementById("wamMatchSelect").value = filters.wamMatch;
  document.getElementById("workShortageRiskSelect").value = filters.workShortageRisk;
  document.getElementById("transportSelect").value = filters.transport;
  document.getElementById("staffingOutlierSelect").value = filters.staffingOutlier;
  document.getElementById("primaryActivitySelect").value = filters.primaryActivity;
  document.getElementById("mealSupportSelect").value = filters.mealSupport;
  document.getElementById("managerMultiSelect").value = filters.managerMulti;
  document.getElementById("newOnlyCheckbox").checked = filters.newOnly;
  document.getElementById("homeUseOnlyCheckbox").checked = filters.homeUseOnly;
  document.getElementById("noufukuOnlyCheckbox").checked = filters.noufukuOnly;
  document.getElementById("wamOnlyCheckbox").checked = filters.wamOnly;
}

function syncPresetButtons() {
  document.querySelectorAll(".preset-chip[data-preset]").forEach((button) => {
    button.classList.toggle("is-active", button.dataset.preset === state.activePreset);
  });
}

const MY_OFFICE_NO = "339";

function syncSelectedRecord(filtered) {
  if (!filtered.length) {
    state.selectedOfficeNo = null;
    return;
  }
  const selectedExists = filtered.some((record) => String(record.office_no) === String(state.selectedOfficeNo));
  if (!selectedExists) {
    const myOffice = filtered.find((record) => String(record.office_no) === MY_OFFICE_NO);
    state.selectedOfficeNo = myOffice ? MY_OFFICE_NO : String(filtered[0].office_no);
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

function normalizeCorporationName(value) {
  return String(value ?? "")
    .normalize("NFKC")
    .trim()
    .replace(/\s+/g, "")
    .toLowerCase();
}

function getCorporationKey(value) {
  const key = normalizeCorporationName(value);
  return key || null;
}

function corporationRecords(corporationName) {
  const key = getCorporationKey(corporationName);
  if (!key) return [];
  return state.records.filter((record) => getCorporationKey(record.corporation_name) === key);
}

function sortCorporationRecords(records, sortKey = state.selectedCorporationSort) {
  const list = [...records];
  return list.sort((left, right) => {
    if (sortKey === "wage") {
      const wageDiff = (right.average_wage_yen ?? Number.NEGATIVE_INFINITY) - (left.average_wage_yen ?? Number.NEGATIVE_INFINITY);
      if (Number.isFinite(wageDiff) && wageDiff !== 0) return wageDiff;
    }
    if (sortKey === "utilization") {
      const utilDiff =
        (right.daily_user_capacity_ratio ?? Number.NEGATIVE_INFINITY) -
        (left.daily_user_capacity_ratio ?? Number.NEGATIVE_INFINITY);
      if (Number.isFinite(utilDiff) && utilDiff !== 0) return utilDiff;
    }
    const municipalityDiff = String(left.municipality ?? "").localeCompare(String(right.municipality ?? ""), "ja");
    if (municipalityDiff !== 0) return municipalityDiff;
    const officeDiff = String(left.office_name ?? "").localeCompare(String(right.office_name ?? ""), "ja");
    if (officeDiff !== 0) return officeDiff;
    return Number(left.office_no ?? 0) - Number(right.office_no ?? 0);
  });
}

function corporationLinkButton(corporationName, className = "entity-link-button") {
  const label = String(corporationName ?? "").trim();
  if (!label) return escapeHtml("-");
  return `<button class="${escapeAttribute(className)}" type="button" data-open-corporation="${escapeAttribute(label)}">${escapeHtml(label)}</button>`;
}

function selectRecord(officeNo, options = {}) {
  const record =
    state.filteredRecords.find((item) => String(item.office_no) === String(officeNo)) ??
    state.records.find((item) => String(item.office_no) === String(officeNo));
  if (!record) return;
  state.selectedOfficeNo = String(officeNo);
  renderCharts(state.filteredRecords);
  renderDetail(record);
  renderTable(state.filteredRecords);
  if (options.openDetail) {
    openDialogElement(document.getElementById("detailDialog"));
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
      getAreaLabel(record),
      record.corporation_name,
      record.office_name,
      record.corporation_type_label,
      record.remarks,
      record.wam_primary_activity_type,
      record.wam_primary_activity_detail,
      record.wam_office_address_city,
      record.wam_office_address_line,
      record.wam_office_number,
      composeAddress(record),
    ]
      .filter(Boolean)
      .join(" ")
      .toLowerCase();
    if (!haystack.includes(filters.search.toLowerCase())) {
      return false;
    }
  }

  if (filters.municipality !== "all" && record.municipality !== filters.municipality) return false;
  if (filters.area !== "all" && getAreaLabel(record) !== filters.area) return false;
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
  if (filters.workShortageRisk === "likely" && !hasWorkShortageRisk(record)) {
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
  populateSelect("areaSelect", "all", "すべてのエリア", uniqueAreaValues(records));
  populateSelect("corporationTypeSelect", "all", "すべての法人種別", uniqueValues(records, "corporation_type_label"));
  populateSelect("responseStatusSelect", "all", "すべての回答状態", ["answered", "annotated", "unanswered"]);
  populateSelect("outlierFlagSelect", "all", "すべての工賃水準", ["high", "low", "none"]);
  populateSelect("capacityBandSelect", "all", "すべての定員帯", CAPACITY_BAND_ORDER);
  populateSelect("quadrantSelect", "all", "すべての工賃と利用の位置", QUADRANT_ORDER);
  populateSelect("wamMatchSelect", "all", "すべての人員詳細", ["matched", "unmatched"]);
  populateSelect("workShortageRiskSelect", "all", "すべての仕事状況", ["likely"]);
  populateSelect("transportSelect", "all", "すべての送迎", ["true", "false", "unknown"]);
  populateSelect("staffingOutlierSelect", "all", "すべての人員配置", ["high", "low", "none"]);
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

function uniqueAreaValues(records) {
  return [...new Set(records.map((record) => getAreaLabel(record)).filter(Boolean))].sort((left, right) =>
    String(left).localeCompare(String(right), "ja")
  );
}

function labelForSelect(value) {
  if (value === "answered") return "回答済み";
  if (value === "annotated") return "注記あり";
  if (value === "unanswered") return "未回答";
  if (value === "likely") return "仕事不足の可能性あり";
  if (value === "high") return "高め";
  if (value === "low") return "低め";
  if (value === "none") return "通常";
  if (value === "matched") return "人員詳細あり";
  if (value === "unmatched") return "人員詳細なし";
  if (value === "true") return "あり";
  if (value === "false") return "なし";
  if (value === "unknown") return "不明";
  if (value === "高工賃 × 高稼働") return "工賃高め・利用多め";
  if (value === "高工賃 × 低稼働") return "工賃高め・利用少なめ";
  if (value === "低工賃 × 高稼働") return "工賃低め・利用多め";
  if (value === "低工賃 × 低稼働") return "工賃低め・利用少なめ";
  if (value === "高工賃 × 厚い人員") return "工賃高め・人員厚め";
  if (value === "高工賃 × 少ない人員") return "工賃高め・人員少なめ";
  if (value === "低工賃 × 厚い人員") return "工賃低め・人員過剰";
  if (value === "低工賃 × 少ない人員") return "工賃低め・人員少なめ";
  return value;
}

function renderLoadingState() {
  const loadingCard = `<div class="loading-message">データを読み込み中...</div>`;
  document.getElementById("openSelectedDetailButton").disabled = true;
  document.getElementById("focusSelectedButton").disabled = true;
  document.getElementById("activeFilterSummary").textContent = "データを読み込み中...";
  const filterTags = document.getElementById("activeFilterTags");
  if (filterTags) {
    filterTags.innerHTML = "";
  }
  document.getElementById("insightList").innerHTML = loadingCard;
  document.getElementById("growthList").innerHTML = loadingCard;
  document.getElementById("fixList").innerHTML = loadingCard;
  document.getElementById("highHighList").innerHTML = loadingCard;
  document.getElementById("statsGrid").innerHTML = loadingCard;
  document.getElementById("recordsCardList").innerHTML = loadingCard;
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
    "detailDialogContent",
  ].forEach((id) => {
    const element = document.getElementById(id);
    if (element) {
      element.innerHTML = loadingCard;
    }
  });
  document.getElementById("recordsTableBody").innerHTML = `<tr><td colspan="13">${loadingCard}</td></tr>`;
}

function renderMeta(dashboard) {
  const matchSummary = dashboard.summary?.wam_match_summary ?? {};
  const updatedAt = dashboard.meta?.generated_at ? new Date(dashboard.meta.generated_at) : null;
  document.getElementById("updatedAt").textContent =
    updatedAt && !Number.isNaN(updatedAt.valueOf()) ? updatedAt.toLocaleString("ja-JP") : "-";
  document.getElementById("totalRecords").textContent = formatCount(state.records.length);
  document.getElementById("wamMatchedCount").textContent = formatCount(matchSummary.matched_record_count ?? 0);
}

function renderQuality(dashboard, records = state.records) {
  const issuesRoot = document.getElementById("issuesList");
  const notesRoot = document.getElementById("notesList");
  const issues = dashboard.issues ?? [];
  const notes = dashboard.notes ?? [];
  const wageStats = computeLocalStats(numericValues(records, "average_wage_yen"));
  const wamMatchSummary = dashboard.analytics?.wam_match_summary ?? {};

  const summaryCards = [
    `<article class="note-card"><strong>表示対象</strong><p>工賃データは大阪府全域 ${formatCount(
      records.length
    )} 件を表示している。大阪市以外は参考比較で見られる。</p></article>`,
    `<article class="note-card"><strong>人員詳細あり とは</strong><p>福祉医療機構の公開情報と結び付いた事業所で、大阪市レコード ${formatCount(
      wamMatchSummary.matched_record_count ?? 0
    )} / ${formatCount(wamMatchSummary.workbook_osaka_record_count ?? 0)} 件が対象である。</p></article>`,
    `<article class="note-card"><strong>差が大きい事業所の見つけ方</strong><p>表示中データの真ん中あたりの工賃は ${formatMaybeYen(
      wageStats.median
    )}。そこから大きく離れた工賃を確認候補として表示している。</p></article>`,
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
  if (state.filters.area !== "all") filters.push(`エリア: ${state.filters.area}`);
  if (state.filters.corporationType !== "all") filters.push(`法人種別: ${state.filters.corporationType}`);
  if (state.filters.responseStatus !== "all") filters.push(`回答状態: ${labelForSelect(state.filters.responseStatus)}`);
  if (state.filters.outlierFlag !== "all") filters.push(`工賃水準: ${labelForSelect(state.filters.outlierFlag)}`);
  if (state.filters.capacityBand !== "all") filters.push(`定員帯: ${state.filters.capacityBand}`);
  if (state.filters.quadrant !== "all") filters.push(`工賃と利用の位置: ${labelForSelect(state.filters.quadrant)}`);
  if (state.filters.wamMatch !== "all") filters.push(`人員詳細: ${labelForSelect(state.filters.wamMatch)}`);
  if (state.filters.workShortageRisk !== "all") filters.push(`仕事状況: ${labelForSelect(state.filters.workShortageRisk)}`);
  if (state.filters.transport !== "all") filters.push(`送迎: ${labelForSelect(state.filters.transport)}`);
  if (state.filters.staffingOutlier !== "all") filters.push(`人員配置: ${labelForSelect(state.filters.staffingOutlier)}`);
  if (state.filters.primaryActivity !== "all") filters.push(`主活動: ${state.filters.primaryActivity}`);
  if (state.filters.mealSupport !== "all") filters.push(`食事加算: ${labelForSelect(state.filters.mealSupport)}`);
  if (state.filters.managerMulti !== "all") filters.push(`管理者兼務: ${labelForSelect(state.filters.managerMulti)}`);
  if (state.filters.newOnly) filters.push("新設のみ");
  if (state.filters.homeUseOnly) filters.push("在宅利用あり");
  if (state.filters.noufukuOnly) filters.push("農福連携あり");
  if (state.filters.wamOnly) filters.push("人員詳細ありのみ");

  document.getElementById("activeFilterSummary").textContent = filters.length
    ? `${formatCount(records.length)} 件を表示中。`
    : `${formatCount(records.length)} 件を表示中。追加条件なし。`;
  const tagsRoot = document.getElementById("activeFilterTags");
  if (tagsRoot) {
    const tags = [
      ...(state.filters.municipality === "all" ? ["対象: 大阪府全域"] : []),
      ...(filters.length ? filters : ["条件なし"]),
    ];
    tagsRoot.innerHTML = tags
      .map((filter) => `<span class="active-filter-chip">${escapeHtml(filter)}</span>`)
      .join("");
  }
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
  const sortedWages = [...wages].sort((a, b) => a - b);

  // 報酬算定区分の分布
  const tierDist = wageTierDistribution(records);
  const modeTier = tierDist.reduce((best, item) => (item.value > (best?.value ?? 0) ? item : best), null);
  const nearRankUp = records.filter((r) => {
    const info = getWageTierUpInfo(r);
    return info && info.next && info.gapYen > 0 && info.gapYen <= 3000;
  });
  const totalRevenueImpact = nearRankUp.reduce((sum, r) => sum + (getWageTierUpInfo(r)?.revenueImpact ?? 0), 0);

  // 大阪府目標
  const targetRate = osakaTargetRate(records);

  // 送迎効果
  const transportTrue = meanFor(matched.filter((r) => r.wam_transport_available === true), "average_wage_yen");
  const transportFalse = meanFor(matched.filter((r) => r.wam_transport_available === false), "average_wage_yen");
  const transportDelta = transportTrue != null && transportFalse != null ? transportTrue - transportFalse : null;

  // 人員配置基準分布
  const tier6Count = matched.filter((r) => { const c = getStaffingComplianceLevel(r); return c && c.level === "tier_6_1"; }).length;
  const tier75Count = matched.filter((r) => { const c = getStaffingComplianceLevel(r); return c && c.level === "tier_7_5_1"; }).length;

  // 好事例
  const topPerformers = topPerformerActivities(records, 3);

  // 利用率1%改善のインパクト
  const lowUtilRecords = records.filter((r) => isNumber(r.daily_user_capacity_ratio) && r.daily_user_capacity_ratio < 0.7);
  const avgRevPerPoint = lowUtilRecords.length
    ? Math.round(lowUtilRecords.reduce((s, r) => s + (revenuePerUtilizationPoint(r) ?? 0), 0) / lowUtilRecords.length)
    : null;

  const insights = [
    {
      title: "📊 報酬算定区分の分布",
      body: modeTier
        ? `最多は ${modeTier.label}（${formatCount(modeTier.value)}件）。ランクアップまで3,000円以内の事業所が ${formatCount(nearRankUp.length)} 件あり、全体の潜在増収は月 ${formatMaybeYen(totalRevenueImpact)} と試算される。`
        : "工賃データが不足しており分析できない。",
    },
    {
      title: "🎯 大阪府工賃向上目標",
      body: targetRate != null
        ? `月額 ${formatCount(OSAKA_WAGE_TARGET_YEN)} 円以上の事業所は ${formatPercent(targetRate)}。中央値は ${formatMaybeYen(wageStats.median)} で${wageStats.median && wageStats.median >= OSAKA_WAGE_TARGET_YEN ? "目標を上回っている" : "まだ目標に届いていない"}。`
        : "工賃データなし。",
    },
    {
      title: "💰 経営者向け: 利用率改善の収益効果",
      body: avgRevPerPoint != null
        ? `利用率70%未満の ${formatCount(lowUtilRecords.length)} 事業所の場合、利用率を1%改善すると平均で月 ${formatMaybeYen(avgRevPerPoint)} の増収。10%改善なら月 ${formatMaybeYen(avgRevPerPoint * 10)} の差になる。`
        : "利用率データが不足。",
    },
    {
      title: "👥 人員配置基準（R6新基準対応）",
      body: matched.length
        ? `6:1達成（最高報酬）: ${formatCount(tier6Count)}件、7.5:1達成: ${formatCount(tier75Count)}件。${tier6Count === 0 ? "6:1に到達している事業所がまだない。人員増強で報酬単価を引き上げる戦略を検討したい。" : ""}`
        : "人員詳細ありの事業所がない。",
    },
    {
      title: "🏆 高工賃の好事例（作業内容）",
      body: topPerformers.length
        ? topPerformers.map((p) => `${p.name}（${formatWageText(p.wage)}）: ${p.detail}`).join(" / ")
        : "工賃が高い事業所の作業内容が確認できない。",
    },
    {
      title: "🚐 送迎実施と工賃の関係",
      body: transportDelta == null
        ? "送迎あり/なしの比較に十分な件数がない。"
        : `送迎ありの平均工賃は ${formatMaybeYen(transportTrue)}、なしは ${formatMaybeYen(transportFalse)}。差分 ${formatSignedYen(transportDelta)}。${transportDelta > 0 ? "送迎は利用者確保につながり、結果として工賃向上に寄与している可能性が高い。" : "送迎の有無と工賃に明確な相関は見られない。"}`,
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
    ["growthList", "fixList", "highHighList"].forEach((id) => {
      document.getElementById(id).innerHTML = document.getElementById("emptyStateTemplate").innerHTML;
    });
    return;
  }

  const growthCandidates = records
    .filter(
      (record) =>
        isNumber(record.average_wage_yen) &&
        isNumber(record.daily_user_capacity_ratio) &&
        (record.wage_ratio_to_municipality_mean ?? record.wage_ratio_to_overall_mean ?? 0) >= 1.1 &&
        record.daily_user_capacity_ratio < 0.8
    )
    .sort((left, right) => growthScore(right) - growthScore(left))
    .slice(0, 6);

  const fixCandidates = records
    .filter(
      (record) =>
        isNumber(record.average_wage_yen) &&
        isNumber(record.daily_user_capacity_ratio) &&
        (record.wage_ratio_to_municipality_mean ?? record.wage_ratio_to_overall_mean ?? 1) < 0.9 &&
        record.daily_user_capacity_ratio < 0.75 &&
        (record.wam_staffing_efficiency_quadrant === "低工賃 × 厚い人員" ||
          record.wam_staffing_outlier_flag === "high")
    )
    .sort((left, right) => fixScore(right) - fixScore(left))
    .slice(0, 6);

  const highHighCandidates = records
    .filter(
      (record) =>
        isNumber(record.average_wage_yen) &&
        isNumber(record.daily_user_capacity_ratio) &&
        (record.wage_ratio_to_municipality_mean ?? record.wage_ratio_to_overall_mean ?? 0) >= 1.1 &&
        record.daily_user_capacity_ratio >= 0.85
    )
    .sort((left, right) => highHighScore(right) - highHighScore(left))
    .slice(0, 6);

  rootSummary.textContent = `高工賃・低利用率 ${formatCount(growthCandidates.length)} 件 / 低工賃・低利用率・人員過剰 ${formatCount(
    fixCandidates.length
  )} 件 / 高工賃・高利用率 ${formatCount(highHighCandidates.length)} 件`;

  renderStrategyList("growthList", growthCandidates, "高工賃・低利用率の事業所はまだない", buildGrowthReason);
  renderStrategyList("fixList", fixCandidates, "低工賃・低利用率・人員過剰の事業所はまだない", buildFixReason);
  renderStrategyList("highHighList", highHighCandidates, "高工賃・高利用率の事業所はまだない", buildHighHighReason);
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
  const sortedWages = [...wages].sort((a, b) => a - b);
  const wageStdDev = computeStdDev(wages);
  const gini = computeGini(wages);
  const p10 = percentile(sortedWages, 0.1);
  const p25 = percentile(sortedWages, 0.25);
  const p75 = percentile(sortedWages, 0.75);
  const welfareStaffFte = meanFor(matched, "wam_welfare_staff_fte_total");
  const keyStaffPerCapacity = meanFor(matched, "wam_key_staff_fte_per_capacity");
  const homeUseActiveCount = records.filter((record) => record.home_use_active === true).length;
  const homeUseActiveRate = ratioOf(records, (record) => record.home_use_active === true);
  const homeUseRateAmongActive = meanFor(
    records.filter((record) => record.home_use_active === true && isNumber(record.home_use_user_ratio_decimal)),
    "home_use_user_ratio_decimal"
  );
  const transportRate = ratioOf(matched, (record) => record.wam_transport_available === true);
  const managerMultiRate = ratioOf(matched, (record) => record.wam_manager_multi_post === true);
  const heavyLowCount = matched.filter((record) => record.wam_staffing_efficiency_quadrant === "低工賃 × 厚い人員").length;
  const workShortageCount = records.filter((record) => hasWorkShortageRisk(record)).length;

  // 新指標: 報酬算定区分
  const tierDist = wageTierDistribution(records);
  const modeTier = tierDist.reduce((best, item) => (item.value > (best?.value ?? 0) ? item : best), null);
  const nearRankUp = records.filter((r) => {
    const info = getWageTierUpInfo(r);
    return info && info.next && info.gapYen > 0 && info.gapYen <= 3000;
  }).length;
  const targetRate = osakaTargetRate(records);

  // 人員配置基準（R6年度 3段階）
  const staffingCritical = matched.filter((r) => {
    const c = getStaffingComplianceLevel(r);
    return c && (c.level === "critical" || c.level === "tier_10_1");
  }).length;
  const staffing6to1 = matched.filter((r) => {
    const c = getStaffingComplianceLevel(r);
    return c && c.level === "tier_6_1";
  }).length;

  // 利用者あたり職員
  const perUserValues = matched.map((r) => staffPerActualUser(r)).filter(isNumber);
  const avgPerUser = perUserValues.length ? perUserValues.reduce((s, v) => s + v, 0) / perUserValues.length : null;

  const cards = [
    { label: "表示件数", value: formatCount(records.length), hint: `全 ${formatCount(state.records.length)} 件中` },
    { label: "平均工賃", value: formatMaybeYen(wageStats.mean), hint: `標準偏差 ${formatMaybeYen(wageStdDev)}` },
    { label: "中央値", value: formatMaybeYen(wageStats.median), hint: `P25: ${formatMaybeYen(p25)} / P75: ${formatMaybeYen(p75)}` },
    { label: "上位10%の目安", value: formatMaybeYen(wageStats.p90), hint: `下位10%: ${formatMaybeYen(p10)}` },
    { label: "最多の報酬算定区分", value: modeTier ? `${modeTier.label}（${formatCount(modeTier.value)}件）` : "-", hint: "令和6年度の9区分で最も多い区分" },
    { label: "ランクアップ接近", value: formatCount(nearRankUp), hint: "あと3,000円以内で次の区分に上がれる" },
    { label: "大阪府目標達成率", value: formatPercent(targetRate), hint: `月額 ${formatCount(OSAKA_WAGE_TARGET_YEN)} 円以上の事業所` },
    { label: "工賃の格差（ジニ係数）", value: isNumber(gini) ? ratioFormatter.format(gini) : "-", hint: "0に近いほど均等、1に近いほど格差大" },
    { label: "人員詳細あり", value: formatCount(matched.length), hint: "WAM公開情報と紐づいた事業所" },
    { label: "平均職員数（常勤換算）", value: formatFte(welfareStaffFte), hint: "人員詳細ありの事業所のみ" },
    { label: "利用者1人あたり職員", value: isNumber(avgPerUser) ? `${ratioFormatter.format(avgPerUser)} 人` : "-", hint: "実利用者数ベース（定員ではない）" },
    { label: "人員基準ギリギリ", value: formatCount(staffingCritical), hint: "10:1最低基準がギリギリ or 未達" },
    { label: "6:1達成（R6新設）", value: formatCount(staffing6to1), hint: `人員詳細ありの ${formatPercent(matched.length ? staffing6to1 / matched.length : null)}` },
    { label: "支援職員 / 定員", value: formatPercent(keyStaffPerCapacity), hint: "主な支援職員の人数を定員で割った目安" },
    { label: "送迎実施率", value: formatPercent(transportRate), hint: "一致レコードのみ" },
    { label: "管理者兼務率", value: formatPercent(managerMultiRate), hint: "一致レコードのみ" },
    { label: "在宅利用あり", value: formatCount(homeUseActiveCount), hint: `表示中の ${formatPercent(homeUseActiveRate)}` },
    { label: "在宅利用ありの平均在宅率", value: formatPercent(homeUseRateAmongActive), hint: "在宅利用ありの事業所のみ" },
    { label: "低工賃・人員過剰", value: formatCount(heavyLowCount), hint: "低工賃・低利用率・人員過剰の候補" },
    { label: "仕事不足の可能性", value: formatCount(workShortageCount), hint: "工賃・利用率・人員のバランスで抽出" },
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
  // 報酬算定区分チャート（wageBandChart の中に追加表示）
  const tierChartEl = document.getElementById("tierChart");
  if (tierChartEl) {
    renderBarChart("tierChart", wageTierDistribution(records), formatCountSuffix("件"));
  }
  renderScatterChart("utilizationScatter", records, {
    xKey: "daily_user_capacity_ratio",
    yKey: "average_wage_yen",
    xLabel: "利用率",
    yLabel: "平均工賃",
    yFormatter: (value) => formatMaybeYen(value),
    xFormatter: (value) => formatPercent(value),
  });
  renderScatterChart("staffingScatter", matchedRecords(records), {
    xKey: "wam_key_staff_fte_per_capacity",
    yKey: "average_wage_yen",
    xLabel: "支援職員 / 定員",
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
          <span class="bar-label">${escapeHtml(labelForSelect(item.label))}</span>
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
    .filter((record) => record.wage_outlier_flag || record.wam_staffing_outlier_flag || (record.daily_user_capacity_ratio ?? 0) > 1.5)
    .sort((left, right) => anomalyScore(right) - anomalyScore(left))
    .slice(0, 12);

  summaryRoot.textContent = anomalies.length
    ? `${formatCount(anomalies.length)} 件を確認候補として表示`
    : "条件に一致する確認候補はない";

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
            ${record.wam_staffing_outlier_flag ? `<span class="metric-chip">${escapeHtml(`人員配置 ${staffingLevelLabel(record.wam_staffing_outlier_flag)}`)}</span>` : ""}
            ${(record.daily_user_capacity_ratio ?? 0) > 1.5 ? `<span class="metric-chip" style="background:rgba(220,38,38,0.1);color:#dc2626">利用率過大 ${formatPercent(record.daily_user_capacity_ratio)}</span>` : ""}
            <span class="metric-chip">${escapeHtml(`工賃 ${formatWageText(record.average_wage_yen)}`)}</span>
            <span class="metric-chip">${escapeHtml(`支援職員 / 定員 ${formatPercent(record.wam_key_staff_fte_per_capacity)}`)}</span>
          </div>
          <p>平均との差 ${escapeHtml(formatRatio(record.wage_ratio_to_overall_mean))} / 支援職員 / 定員 ${escapeHtml(formatPercent(record.wam_key_staff_fte_per_capacity))}</p>
          <p>${escapeHtml(labelForSelect(record.wam_staffing_efficiency_quadrant ?? "人員配置の分類なし"))} / ${escapeHtml(record.wam_primary_activity_type ?? "活動不明")}</p>
        </button>
      `
    )
    .join("");
}

function renderDetail(record) {
  const summaryRoot = document.getElementById("detailContent");
  const dialogRoot = document.getElementById("detailDialogContent");
  const openButton = document.getElementById("openSelectedDetailButton");
  const focusButton = document.getElementById("focusSelectedButton");

  if (!record) {
    if (summaryRoot) {
      summaryRoot.innerHTML = document.getElementById("emptyStateTemplate").innerHTML;
    }
    if (dialogRoot) {
      dialogRoot.innerHTML = document.getElementById("emptyStateTemplate").innerHTML;
    }
    if (openButton) openButton.disabled = true;
    if (focusButton) focusButton.disabled = true;
    return;
  }

  const actionNotes = buildActionNotes(record);
  const officeUrl = safeExternalUrl(record.wam_office_url);
  const homepageUrl = safeExternalUrl(record.homepage_url) ?? (officeUrl && !officeUrl.includes("instagram.com") ? officeUrl : null);
  const instagramUrl = safeExternalUrl(record.instagram_url) ?? (officeUrl?.includes("instagram.com") ? officeUrl : null);
  const websiteSearchUrl = buildWebsiteSearchUrl(record);
  const instagramSearchUrl = buildInstagramSearchUrl(record);
  const homepageSource = record.homepage_source ? linkSourceLabel(record.homepage_source) : null;
  const instagramSource = record.instagram_source ? linkSourceLabel(record.instagram_source) : null;
  const corporationName = String(record.corporation_name ?? "").trim();
  const corporationActionButton = corporationName
    ? `<button class="ghost-button" type="button" data-open-corporation="${escapeAttribute(corporationName)}">法人の事業所一覧</button>`
    : "";

  if (openButton) openButton.disabled = false;
  if (focusButton) focusButton.disabled = false;

  if (summaryRoot) {
    summaryRoot.innerHTML = `
      <article class="selected-office-card">
        <div class="selected-office-head">
          <div>
            <p class="section-kicker">${escapeHtml(record.municipality ?? "-")} / No.${escapeHtml(record.office_no ?? "-")}</p>
            <h3>${escapeHtml(record.office_name ?? "-")}</h3>
            <p class="detail-subtitle">${corporationLinkButton(record.corporation_name, "entity-link-button detail-entity-link")} / ${escapeHtml(record.corporation_type_label ?? "-")}</p>
          </div>
          <div class="selected-office-status">
            ${matchBadge(record.wam_match_status, record.wam_match_confidence)}
            ${record.wage_outlier_flag ? outlierBadge(record.wage_outlier_flag, record.wage_outlier_severity) : ""}
          </div>
        </div>
        <div class="selected-office-metrics">
          <span class="metric-chip">${escapeHtml(`平均工賃 ${formatWageText(record.average_wage_yen)}`)}</span>
          <span class="metric-chip">${escapeHtml(`利用率 ${formatPercent(record.daily_user_capacity_ratio)}`)}</span>
          <span class="metric-chip">${escapeHtml(`在宅率 ${formatPercent(record.home_use_user_ratio_decimal)}`)}</span>
          <span class="metric-chip">${escapeHtml(`支援職員 / 定員 ${formatPercent(record.wam_key_staff_fte_per_capacity)}`)}</span>
        </div>
        <p class="selected-office-note">${escapeHtml(actionNotes[0] ?? "詳しい内訳は詳細を開いて確認できる。")}</p>
        <div class="selected-office-actions">
          ${corporationActionButton}
          ${homepageUrl ? `<a class="ghost-button link-button" href="${escapeAttribute(homepageUrl)}" target="_blank" rel="noreferrer">ホームページ</a>` : `<a class="ghost-button link-button" href="${escapeAttribute(websiteSearchUrl)}" target="_blank" rel="noreferrer">Webで探す</a>`}
          ${instagramUrl ? `<a class="ghost-button link-button" href="${escapeAttribute(instagramUrl)}" target="_blank" rel="noreferrer">Instagram</a>` : `<a class="ghost-button link-button" href="${escapeAttribute(instagramSearchUrl)}" target="_blank" rel="noreferrer">Instagramを探す</a>`}
        </div>
      </article>
    `;
  }

  if (!dialogRoot) return;

  const comparisonContext = getDetailComparisonContext(record);
  const tierInfo = getWageTierUpInfo(record);
  const staffComp = getStaffingComplianceLevel(record);
  const peer = computePeerBenchmark(record, comparisonContext.records);
  const perUser = staffPerActualUser(record);
  const missedAddons = checkMissedAddons(record);
  const revPerPoint = revenuePerUtilizationPoint(record);
  const sortedWages = numericValues(comparisonContext.records, "average_wage_yen").sort((a, b) => a - b);
  const wagePercentile = computePercentileRank(record.average_wage_yen, sortedWages);

  dialogRoot.innerHTML = `
    <div class="detail-hero">
      <div>
        <p class="section-kicker">${escapeHtml(record.municipality ?? "-")} / No.${escapeHtml(record.office_no ?? "-")}</p>
        <h3>${escapeHtml(record.office_name ?? "-")}</h3>
        <p class="detail-subtitle">${corporationLinkButton(record.corporation_name, "entity-link-button detail-entity-link")} / ${escapeHtml(record.corporation_type_label ?? "-")}</p>
      </div>
      <div class="detail-cta">
        ${corporationActionButton}
        ${homepageUrl ? `<a class="solid-button link-button" href="${escapeAttribute(homepageUrl)}" target="_blank" rel="noreferrer">ホームページ</a>` : `<a class="solid-button link-button" href="${escapeAttribute(websiteSearchUrl)}" target="_blank" rel="noreferrer">Webで探す</a>`}
        ${instagramUrl ? `<a class="ghost-button link-button" href="${escapeAttribute(instagramUrl)}" target="_blank" rel="noreferrer">Instagram</a>` : `<a class="ghost-button link-button" href="${escapeAttribute(instagramSearchUrl)}" target="_blank" rel="noreferrer">Instagramを探す</a>`}
      </div>
    </div>
    <div class="detail-kpi-grid">
      ${detailKpi("平均工賃", formatWageText(record.average_wage_yen), `${wagePercentile != null ? `上位 ${100 - wagePercentile}%` : "-"} / 平均との差 ${formatRatio(record.wage_ratio_to_overall_mean)}`)}
      ${detailKpi(
        "報酬算定区分",
        tierInfo?.current ? `${tierInfo.current.label}（${formatCount(tierInfo.current.unitYen)}円/日）` : "-",
        tierInfo?.next ? `次の${tierInfo.next.label}まで あと ${formatCount(tierInfo.gapYen)} 円 → 月 ${formatMaybeYen(tierInfo.revenueImpact)} 増収` : "最上位区分"
      )}
      ${detailKpi("利用率", formatPercent(record.daily_user_capacity_ratio), `定員 ${formatNullable(record.capacity)} 名 / 平均利用 ${formatNumber(record.average_daily_users)} 名`)}
      ${detailKpi(
        "人員配置基準",
        staffComp ? staffComp.label : record.wam_match_status === "matched" ? "-" : "詳細なし",
        staffComp ? `${staffComp.qualifiedTier} / 職員FTE ${formatFte(staffComp.staffFte)}` : ""
      )}
      ${detailKpi(
        "利用者1人あたり職員",
        isNumber(perUser) ? `${ratioFormatter.format(perUser)} 人` : "-",
        record.wam_match_status === "matched" ? `支援職員 / 定員 ${formatPercent(record.wam_key_staff_fte_per_capacity)}` : "人員詳細なし"
      )}
      ${detailKpi("公開情報の月額工賃", formatMaybeYen(record.wam_average_wage_monthly_yen), `Excelとの差 ${formatSignedYen(record.wam_average_wage_gap_yen)}`)}
    </div>
    <div class="detail-grid">
      <article class="detail-card detail-card-wide detail-comparison-card">
        <div class="detail-comparison-head">
          <div>
            <h3>主要指標と平均との差・中央値差</h3>
            <p class="detail-card-note">比較対象: ${escapeHtml(comparisonContext.label)}</p>
          </div>
          <p class="detail-card-note">${escapeHtml(comparisonContext.note)}</p>
        </div>
        <div class="detail-comparison-legend" aria-hidden="true">
          <span><i class="detail-legend-line is-mean"></i>平均</span>
          <span><i class="detail-legend-line is-median"></i>中央値</span>
          <span><i class="detail-legend-dot"></i>この事業所</span>
        </div>
        <div class="detail-comparison-list">
          ${renderDetailComparisonRows(record, comparisonContext)}
        </div>
      </article>
      <article class="detail-card">
        <h3>ピアベンチマーク</h3>
        <ul class="detail-list">
          ${peer.activity ? `<li>【${escapeHtml(peer.activity.type)}】${formatCount(peer.activity.count)}事業所中 ${formatCount(peer.activity.rank)}位（中央値 ${formatMaybeYen(peer.activity.median)}）</li>` : "<li>主活動のピアデータなし</li>"}
          ${peer.capacity ? `<li>【${escapeHtml(peer.capacity.band)}】${formatCount(peer.capacity.count)}事業所中 ${formatCount(peer.capacity.rank)}位（中央値 ${formatMaybeYen(peer.capacity.median)}）</li>` : "<li>定員帯のピアデータなし</li>"}
          <li>市町村平均との差: ${escapeHtml(formatRatio(record.wage_ratio_to_municipality_mean))}</li>
          <li>法人種別平均との差: ${escapeHtml(formatRatio(record.wage_ratio_to_corporation_type_mean))}</li>
          <li>工賃と利用の位置: ${escapeHtml(labelForSelect(record.market_position_quadrant ?? "-"))}</li>
        </ul>
      </article>
      <article class="detail-card">
        <h3>経営者向けシミュレーション</h3>
        <ul class="detail-list">
          ${tierInfo?.next ? `<li>ランクアップに必要な工賃改善: あと ${formatCount(tierInfo.gapYen)} 円</li>` : "<li>最上位の報酬算定区分</li>"}
          ${tierInfo?.revenueImpact ? `<li>ランクアップ時の増収: 月 ${formatMaybeYen(tierInfo.revenueImpact)}</li>` : ""}
          ${revPerPoint ? `<li>利用率1%改善: 月 ${formatMaybeYen(revPerPoint)} の増収</li>` : ""}
          ${missedAddons.length ? missedAddons.map((m) => `<li>💡 ${escapeHtml(m.name)}: ${escapeHtml(m.hint)}</li>`).join("") : "<li>主要加算の取り漏れは確認されない</li>"}
        </ul>
      </article>
      <article class="detail-card">
        <h3>運営体制</h3>
        <ul class="detail-list">
          <li>人員詳細: ${escapeHtml(record.wam_match_status === "matched" ? `あり / ${matchConfidenceLabel(record.wam_match_confidence)}` : "なし")}</li>
          <li>送迎: ${escapeHtml(formatBool(record.wam_transport_available))}</li>
          <li>食事加算: ${escapeHtml(formatBool(record.wam_meal_support_addon))}</li>
          <li>管理者兼務: ${escapeHtml(formatBool(record.wam_manager_multi_post))}</li>
          <li>主活動: ${escapeHtml(record.wam_primary_activity_type ?? "-")} ${record.wam_primary_activity_detail ? `（${escapeHtml(record.wam_primary_activity_detail)}）` : ""}</li>
          <li>サービス管理責任者（常勤換算）: ${escapeHtml(formatFte(record.wam_service_manager_fte))}</li>
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
          <li>公開情報の定員: ${escapeHtml(formatNullable(record.wam_office_capacity))}</li>
          <li>在宅利用: ${escapeHtml(formatBool(record.home_use_active))} ${isNumber(record.home_use_user_ratio_decimal) ? `（在宅率 ${formatPercent(record.home_use_user_ratio_decimal)}）` : ""}</li>
          <li>新設: ${escapeHtml(record.is_new_office ? "あり" : "なし")}</li>
          <li>備考: ${escapeHtml(record.remarks ?? "-")}</li>
        </ul>
      </article>
      <article class="detail-card">
        <h3>外部リンク</h3>
        <ul class="detail-list">
          <li>ホームページ: ${homepageUrl ? `<a class="link-button" href="${escapeAttribute(homepageUrl)}" target="_blank" rel="noreferrer">開く</a>` : "未登録"}${homepageSource ? ` / ${escapeHtml(homepageSource)}` : ""}</li>
          <li>Instagram: ${instagramUrl ? `<a class="link-button" href="${escapeAttribute(instagramUrl)}" target="_blank" rel="noreferrer">開く</a>` : "未登録"}${instagramSource ? ` / ${escapeHtml(instagramSource)}` : ""}</li>
          <li>検索補助: <a class="link-button" href="${escapeAttribute(websiteSearchUrl)}" target="_blank" rel="noreferrer">Webで探す</a> / <a class="link-button" href="${escapeAttribute(instagramSearchUrl)}" target="_blank" rel="noreferrer">Instagramを探す</a></li>
        </ul>
      </article>
    </div>
  `;
}

function renderCorporationDialog(corporationName) {
  const dialogRoot = document.getElementById("corporationDialogContent");
  if (!dialogRoot) return;

  const records = sortCorporationRecords(corporationRecords(corporationName));
  state.selectedCorporationKey = getCorporationKey(corporationName);
  state.selectedCorporationLabel = corporationName;

  if (!records.length) {
    dialogRoot.innerHTML = document.getElementById("emptyStateTemplate").innerHTML;
    return;
  }

  const filteredOfficeNos = new Set(state.filteredRecords.map((record) => String(record.office_no)));
  const visibleCount = records.filter((record) => filteredOfficeNos.has(String(record.office_no))).length;
  const municipalityCount = new Set(records.map((record) => record.municipality).filter(Boolean)).size;
  const osakaCityCount = records.filter((record) => record.municipality === "大阪市").length;
  const matchedCount = records.filter((record) => record.wam_match_status === "matched").length;
  const answeredWages = records.map((record) => record.average_wage_yen).filter(isNumber);
  const averageWage = answeredWages.length
    ? answeredWages.reduce((sum, value) => sum + value, 0) / answeredWages.length
    : null;
  const corporationTypeCounts = new Map();
  records.forEach((record) => {
    const label = record.corporation_type_label;
    if (!label) return;
    corporationTypeCounts.set(label, (corporationTypeCounts.get(label) ?? 0) + 1);
  });
  const corporationTypeLabel =
    [...corporationTypeCounts.entries()].sort((left, right) => right[1] - left[1])[0]?.[0] ?? "-";
  const sortButtons = CORPORATION_SORT_OPTIONS.map(
    (option) => `
      <button
        class="preset-chip ${state.selectedCorporationSort === option.key ? "is-active" : ""}"
        type="button"
        data-corporation-sort="${escapeAttribute(option.key)}"
      >${escapeHtml(option.label)}</button>
    `
  ).join("");

  dialogRoot.innerHTML = `
    <div class="corporation-hero">
      <div>
        <p class="section-kicker">法人名からまとめて確認</p>
        <h3>${escapeHtml(corporationName)}</h3>
        <p class="detail-subtitle">${escapeHtml(corporationTypeLabel)} / 大阪府内で ${formatCount(records.length)} 事業所</p>
      </div>
      <div class="selected-office-actions">
        <button class="ghost-button" type="button" data-select-office="${escapeAttribute(records[0].office_no)}">先頭の事業所を開く</button>
      </div>
    </div>
    <div class="corporation-toolbar">
      <div class="corporation-sort-group">
        <span class="corporation-sort-label">並び替え</span>
        <div class="preset-row corporation-sort-row">${sortButtons}</div>
      </div>
    </div>
    <div class="corporation-summary-grid">
      <article class="corporation-summary-card">
        <span>運営事業所数</span>
        <strong>${formatCount(records.length)}</strong>
        <em>大阪府全体</em>
      </article>
      <article class="corporation-summary-card">
        <span>現在の一覧に出ている数</span>
        <strong>${formatCount(visibleCount)}</strong>
        <em>今の絞り込み条件に入る事業所</em>
      </article>
      <article class="corporation-summary-card">
        <span>大阪市内</span>
        <strong>${formatCount(osakaCityCount)}</strong>
        <em>大阪市の運営事業所</em>
      </article>
      <article class="corporation-summary-card">
        <span>人員詳細あり</span>
        <strong>${formatCount(matchedCount)}</strong>
        <em>WAM と突合済み</em>
      </article>
      <article class="corporation-summary-card">
        <span>平均工賃の平均</span>
        <strong>${escapeHtml(formatWageText(averageWage))}</strong>
        <em>${formatCount(municipalityCount)} 市区にまたがる</em>
      </article>
    </div>
    <article class="detail-card">
      <h3>運営している事業所一覧</h3>
      <ul class="detail-list corporation-note-list">
        <li>現在の絞り込みに入っている事業所には「一覧内」を付けている。</li>
        <li>事業所を押すと、そのまま詳細モーダルへ移動する。</li>
      </ul>
      <div class="corporation-office-list">
        ${records
          .map((record) => {
            const isVisible = filteredOfficeNos.has(String(record.office_no));
            const isCurrent = String(record.office_no) === String(state.selectedOfficeNo);
            return `
              <button class="corporation-office-item ${isCurrent ? "is-current" : ""}" type="button" data-select-office="${escapeAttribute(record.office_no)}">
                <div class="corporation-office-head">
                  <div>
                    <p class="section-kicker">${escapeHtml(getAreaLabel(record) || record.municipality || "-")} / No.${escapeHtml(record.office_no ?? "-")}</p>
                    <h3>${escapeHtml(record.office_name ?? "-")}</h3>
                    <p class="detail-subtitle">${escapeHtml(composeAddress(record))}</p>
                  </div>
                  <div class="selected-office-status">
                    ${isVisible ? `<span class="badge badge-answered">一覧内</span>` : `<span class="badge badge-annotated">一覧外</span>`}
                    ${matchBadge(record.wam_match_status, record.wam_match_confidence)}
                  </div>
                </div>
                <div class="corporation-office-meta">
                  <span class="metric-chip">${escapeHtml(`平均工賃 ${formatWageText(record.average_wage_yen)}`)}</span>
                  <span class="metric-chip">${escapeHtml(`利用率 ${formatPercent(record.daily_user_capacity_ratio)}`)}</span>
                  <span class="metric-chip">${escapeHtml(`在宅率 ${formatPercent(record.home_use_user_ratio_decimal)}`)}</span>
                  <span class="metric-chip">${escapeHtml(`支援職員 / 定員 ${formatPercent(record.wam_key_staff_fte_per_capacity)}`)}</span>
                  ${attentionBadges(record)}
                </div>
                <p class="corporation-office-note">${escapeHtml(record.wam_primary_activity_type ?? "主活動の記載なし")} / ${escapeHtml(record.remarks ?? "備考なし")}</p>
              </button>
            `;
          })
          .join("")}
      </div>
    </article>
  `;
}

function renderTable(records) {
  const tableBody = document.getElementById("recordsTableBody");
  const cardList = document.getElementById("recordsCardList");
  const pageCount = Math.max(1, Math.ceil(records.length / PAGE_SIZE));
  const startIndex = (state.currentPage - 1) * PAGE_SIZE;
  const pagedRecords = records.slice(startIndex, startIndex + PAGE_SIZE);

  document.getElementById("tableSummary").textContent = `${formatCount(records.length)} 件 / ${formatCount(pageCount)} ページ`;
  document.getElementById("pageSummary").textContent = `${formatCount(state.currentPage)} / ${formatCount(pageCount)} ページ`;
  document.getElementById("prevPageButton").disabled = state.currentPage <= 1;
  document.getElementById("nextPageButton").disabled = state.currentPage >= pageCount;

  if (!pagedRecords.length) {
    tableBody.innerHTML = `<tr><td colspan="13">${document.getElementById("emptyStateTemplate").innerHTML}</td></tr>`;
    if (cardList) {
      cardList.innerHTML = document.getElementById("emptyStateTemplate").innerHTML;
    }
    return;
  }

  tableBody.innerHTML = pagedRecords
    .map(
      (record) => `
        <tr class="${rowClass(record)}">
          <td class="numeric">${formatNullable(record.office_no)}</td>
          <td>${escapeHtml(record.office_name ?? "-")}</td>
          <td>${escapeHtml(getAreaLabel(record) || record.municipality || "-")}</td>
          <td>${corporationLinkButton(record.corporation_name, "entity-link-button table-entity-link")}</td>
          <td class="numeric">${formatWage(record.average_wage_yen, record.average_wage_error)}</td>
          <td class="numeric">${formatRatio(record.wage_ratio_to_overall_mean)}</td>
          <td class="numeric">${formatPercent(record.daily_user_capacity_ratio)}</td>
          <td class="numeric">${formatPercent(record.home_use_user_ratio_decimal)}</td>
          <td>${matchBadge(record.wam_match_status, record.wam_match_confidence)}</td>
          <td class="numeric">${formatPercent(record.wam_key_staff_fte_per_capacity)}</td>
          <td>${escapeHtml(record.wam_primary_activity_type ?? "-")}</td>
          <td><div class="attention-cell">${attentionBadges(record)}</div></td>
          <td><button class="table-link" data-select-office="${escapeHtml(record.office_no)}" type="button">詳細</button></td>
        </tr>
      `
    )
    .join("");

  if (cardList) {
    cardList.innerHTML = pagedRecords
      .map(
        (record) => `
          <article class="record-card ${rowClass(record)}">
            <div class="record-card-head">
              <div>
                <p class="section-kicker">${escapeHtml(record.municipality ?? "-")} / No.${escapeHtml(record.office_no ?? "-")}</p>
                <h3>${escapeHtml(record.office_name ?? "-")}</h3>
                <p class="detail-subtitle">${corporationLinkButton(record.corporation_name, "entity-link-button detail-entity-link")}</p>
              </div>
              <button class="table-link" data-select-office="${escapeHtml(record.office_no)}" type="button">詳細</button>
            </div>
            <div class="record-card-metrics">
              <span class="metric-chip">${escapeHtml(`工賃 ${formatWageText(record.average_wage_yen)}`)}</span>
              <span class="metric-chip">${escapeHtml(`利用率 ${formatPercent(record.daily_user_capacity_ratio)}`)}</span>
              <span class="metric-chip">${escapeHtml(`在宅率 ${formatPercent(record.home_use_user_ratio_decimal)}`)}</span>
              <span class="metric-chip">${escapeHtml(`支援職員 / 定員 ${formatPercent(record.wam_key_staff_fte_per_capacity)}`)}</span>
            </div>
            <div class="record-card-status">
              ${matchBadge(record.wam_match_status, record.wam_match_confidence)}
              ${attentionBadges(record)}
            </div>
            <p class="record-card-note">${escapeHtml(record.wam_primary_activity_type ?? "主活動の記載なし")}</p>
          </article>
        `
      )
      .join("");
  }
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

function attentionBadges(record) {
  const badges = [];
  if (record.wage_outlier_flag) {
    badges.push(outlierBadge(record.wage_outlier_flag, record.wage_outlier_severity));
  }
  if (record.wam_staffing_outlier_flag) {
    badges.push(staffingOutlierBadge(record.wam_staffing_outlier_flag, record.wam_staffing_outlier_severity));
  }
  if (record.response_status && record.response_status !== "answered") {
    badges.push(statusBadge(record.response_status));
  }
  if (record.is_new_office) {
    badges.push(`<span class="badge badge-annotated">新設</span>`);
  }
  if (hasWorkShortageRisk(record)) {
    badges.push(`<span class="badge badge-annotated">仕事不足の可能性</span>`);
  }
  if (!badges.length) {
    badges.push(`<span class="badge badge-answered">通常</span>`);
  }
  return badges.join("");
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
  if (!values.length) {
    return {
      count: 0,
      min: null,
      p10: null,
      mean: null,
      median: null,
      p90: null,
      max: null,
    };
  }
  const sorted = [...values].sort((left, right) => left - right);
  return {
    count: sorted.length,
    min: sorted[0],
    p10: percentile(sorted, 0.1),
    mean: sorted.reduce((sum, value) => sum + value, 0) / sorted.length,
    median: median(sorted),
    p90: percentile(sorted, 0.9),
    max: sorted[sorted.length - 1],
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
    { label: "人員詳細あり", value: matched },
    { label: "人員詳細なし", value: unmatched },
  ];
}

function combinedOutlierBreakdown(records) {
  return [
    { label: "工賃高め", value: records.filter((record) => record.wage_outlier_flag === "high").length },
    { label: "工賃低め", value: records.filter((record) => record.wage_outlier_flag === "low").length },
    { label: "人員多め", value: records.filter((record) => record.wam_staffing_outlier_flag === "high").length },
    { label: "人員少なめ", value: records.filter((record) => record.wam_staffing_outlier_flag === "low").length },
  ].filter((item) => item.value > 0);
}

function staffingRoleAverages(records) {
  const matched = matchedRecords(records);
  const rows = [
    { label: "サービス管理責任者", value: meanFor(matched, "wam_service_manager_fte") },
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
  const utilizationOver150 = (record.daily_user_capacity_ratio ?? 0) > 1.5 ? 35 : 0;
  return wageScore + staffingScore + utilizationOver150;
}

function growthScore(record) {
  return (
    (record.wage_ratio_to_municipality_mean ?? record.wage_ratio_to_overall_mean ?? 0) * 70 +
    (1 - (record.daily_user_capacity_ratio ?? 1)) * 25 +
    (record.wam_staffing_efficiency_quadrant === "高工賃 × 少ない人員" ? 12 : 0)
  );
}

function fixScore(record) {
  return (
    Math.max(1 - (record.wage_ratio_to_municipality_mean ?? record.wage_ratio_to_overall_mean ?? 1), 0) * 80 +
    ((record.wam_key_staff_fte_per_capacity ?? 0) * 100) +
    Math.max(0.7 - (record.daily_user_capacity_ratio ?? 0.7), 0) * 25
  );
}

function highHighScore(record) {
  return (
    (record.wage_ratio_to_municipality_mean ?? record.wage_ratio_to_overall_mean ?? 0) * 70 +
    (record.daily_user_capacity_ratio ?? 0) * 35
  );
}

function buildGrowthReason(record) {
  return `市町村平均との差 ${formatRatio(record.wage_ratio_to_municipality_mean)} / 利用率 ${formatPercent(
    record.daily_user_capacity_ratio
  )} / 支援職員 / 定員 ${formatPercent(record.wam_key_staff_fte_per_capacity)}`;
}

function buildFixReason(record) {
  const workShortageNote = hasWorkShortageRisk(record) ? " / 仕事不足の可能性" : "";
  return `市町村平均との差 ${formatRatio(record.wage_ratio_to_municipality_mean)} / 利用率 ${formatPercent(
    record.daily_user_capacity_ratio
  )} / 人員配置 ${labelForSelect(record.wam_staffing_efficiency_quadrant ?? "-")}${workShortageNote}`;
}

function buildHighHighReason(record) {
  return `市町村平均との差 ${formatRatio(record.wage_ratio_to_municipality_mean)} / 利用率 ${formatPercent(
    record.daily_user_capacity_ratio
  )} / 在宅率 ${formatPercent(record.home_use_user_ratio_decimal)}`;
}

function buildActionNotes(record) {
  const notes = [];

  // 1. 報酬算定区分とランクアップ
  const tierInfo = getWageTierUpInfo(record);
  if (tierInfo) {
    if (tierInfo.next && tierInfo.gapYen > 0) {
      notes.push(`【報酬算定】現在 ${tierInfo.current.label}（単価 ${formatCount(tierInfo.current.unitYen)} 円/日）。次の${tierInfo.next.label}まであと ${formatCount(tierInfo.gapYen)} 円。ランクアップで月 ${formatMaybeYen(tierInfo.revenueImpact)} の増収見込み。`);
    } else if (tierInfo.current) {
      notes.push(`【報酬算定】${tierInfo.current.label}（単価 ${formatCount(tierInfo.current.unitYen)} 円/日）。${tierInfo.current.tierNo >= 8 ? "上位区分を維持できている。工賃水準を落とさないよう作業受注の安定化が重要。" : ""}`);
    }
  }

  // 2. 人員配置基準（R6年度3段階）
  const staffing = getStaffingComplianceLevel(record);
  if (staffing) {
    notes.push(`【人員配置】${staffing.label}（${staffing.qualifiedTier}）。${staffing.advice}`);
  } else if (record.wam_match_status !== "matched") {
    notes.push("【人員配置】公開情報と紐づいていないため人員配置基準の評価ができない。");
  }

  // 3. 大阪府目標
  if (isNumber(record.average_wage_yen)) {
    if (record.average_wage_yen < OSAKA_WAGE_TARGET_YEN) {
      const gap = OSAKA_WAGE_TARGET_YEN - record.average_wage_yen;
      notes.push(`【工賃向上計画】大阪府目標（月額 ${formatCount(OSAKA_WAGE_TARGET_YEN)} 円）未達。あと ${formatCount(Math.round(gap))} 円の向上が必要。工賃向上計画の見直しを推奨。`);
    }
  }

  // 4. 加算取り漏れ
  const missed = checkMissedAddons(record);
  if (missed.length > 0) {
    notes.push(`【加算の検討】${missed.map((m) => m.name).join("、")}の導入余地あり。${missed[0].hint}`);
  }

  // 5. 利用率
  if (isNumber(record.daily_user_capacity_ratio) && record.daily_user_capacity_ratio < 0.6) {
    const revPerPoint = revenuePerUtilizationPoint(record);
    notes.push(`【利用率】${formatPercent(record.daily_user_capacity_ratio)} と低水準。${revPerPoint ? `利用率10%改善で 月 ${formatMaybeYen(revPerPoint * 10)} の増収が見込める。体験利用の導線強化と相談支援事業所への営業強化を検討。` : "体験利用・紹介元の強化を検討したい。"}`);
  }

  // 6. 仕事不足リスク
  if (hasWorkShortageRisk(record)) {
    notes.push("【仕事不足】工賃と利用率の両方が弱め。企業営業・新規作業種目の開拓・施設外就労の検討を。");
  }

  // 7. 高工賃の好事例
  if (record.wage_outlier_flag === "high") {
    notes.push("【好事例】高工賃事業所。主活動と取引構造を他事業所への横展開候補として注目。作業内容と品質管理体制の共有を推奨。");
  }

  // 8. 新設
  if (record.is_new_office) {
    notes.push("【新設】立ち上がり期のため単月値ではなく定員充足の推移を追う。3か月目以降の利用率と工賃のトレンドが重要。");
  }

  if (!notes.length) {
    notes.push("大きな警戒シグナルなし。報酬算定区分の維持と利用率の安定を定点観測したい。");
  }
  return notes.slice(0, 6);
}

function getDetailComparisonContext(record) {
  const filteredCount = state.filteredRecords.length;
  const recordInFiltered = state.filteredRecords.some((item) => String(item.office_no) === String(record.office_no));
  if (filteredCount && recordInFiltered) {
    const usesAllRecords = filteredCount === state.records.length;
    return {
      records: state.filteredRecords,
      label: usesAllRecords ? `大阪府全体 ${formatCount(filteredCount)}件` : `現在の絞り込み結果 ${formatCount(filteredCount)}件`,
      note: usesAllRecords
        ? "平均との差・中央値差は大阪府全体との比較で表示している。"
        : "平均との差・中央値差は現在の絞り込み結果を母集団にしている。絞り込みを変えると比較値も変わる。",
    };
  }

  return {
    records: state.records,
    label: `大阪府全体 ${formatCount(state.records.length)}件`,
    note: "この事業所は現在の一覧外のため、大阪府全体を母集団にして比較している。",
  };
}

function getDetailComparisonMetricConfigs() {
  return [
    {
      key: "average_wage_yen",
      label: "平均工賃",
      description: "1人あたりの平均月額工賃",
      emptyText: "未回答",
      formatValue: formatMaybeYen,
      formatDiff: formatSignedYen,
    },
    {
      key: "daily_user_capacity_ratio",
      label: "利用率",
      description: "定員に対する平均利用人数",
      emptyText: "未回答",
      formatValue: formatPercent,
      formatDiff: formatSignedPercentPoint,
    },
    {
      key: "average_daily_users",
      label: "平均利用人数",
      description: "1日あたりの平均利用人数",
      emptyText: "未回答",
      formatValue: (value) => formatDecimalUnit(value, "人"),
      formatDiff: (value) => formatSignedDecimalUnit(value, "人"),
    },
    {
      key: "capacity",
      label: "定員",
      description: "届出上の受入定員",
      emptyText: "未回答",
      formatValue: (value) => formatIntegerUnit(value, "名"),
      formatDiff: (value) => formatSignedIntegerUnit(value, "名"),
    },
    {
      key: "home_use_user_ratio_decimal",
      label: "在宅率",
      description: "在宅利用者の構成比",
      emptyText: "未回答",
      formatValue: formatPercent,
      formatDiff: formatSignedPercentPoint,
    },
    {
      key: "wam_key_staff_fte_per_capacity",
      label: "支援職員 / 定員",
      description: "定員1人あたりの主要支援職員",
      emptyText: "人員詳細なし",
      formatValue: formatPercent,
      formatDiff: formatSignedPercentPoint,
    },
    {
      key: "wam_welfare_staff_fte_total",
      label: "福祉職員（常勤換算）",
      description: "就労支援員・職業指導員・生活支援員などの合計",
      emptyText: "人員詳細なし",
      formatValue: formatFte,
      formatDiff: (value) => formatSignedDecimalUnit(value, "人分"),
    },
    {
      key: "wam_service_manager_fte",
      label: "サービス管理責任者（常勤換算）",
      description: "公開情報にあるサービス管理責任者の人数",
      emptyText: "人員詳細なし",
      formatValue: formatFte,
      formatDiff: (value) => formatSignedDecimalUnit(value, "人分"),
    },
    {
      key: "wam_average_wage_monthly_yen",
      label: "公開情報の月額工賃",
      description: "WAM公開情報に掲載された月額工賃",
      emptyText: "公開情報なし",
      formatValue: formatMaybeYen,
      formatDiff: formatSignedYen,
    },
  ];
}

function comparisonTrackPosition(value, min, max) {
  if (!isNumber(value) || !isNumber(min) || !isNumber(max)) return null;
  if (max <= min) return 50;
  const ratio = ((value - min) / (max - min)) * 100;
  return Math.max(0, Math.min(100, ratio));
}

function deltaClassName(value) {
  if (!isNumber(value)) return "is-neutral";
  if (value > 0) return "is-positive";
  if (value < 0) return "is-negative";
  return "is-neutral";
}

function renderDetailComparisonRows(record, comparisonContext) {
  return getDetailComparisonMetricConfigs()
    .map((metric) => {
      const values = numericValues(comparisonContext.records, metric.key);
      const stats = computeLocalStats(values);
      const currentValue = record[metric.key];
      const meanDiff = isNumber(currentValue) && isNumber(stats.mean) ? currentValue - stats.mean : null;
      const medianDiff = isNumber(currentValue) && isNumber(stats.median) ? currentValue - stats.median : null;
      const trackMin = isNumber(stats.p10) ? stats.p10 : stats.min;
      const trackMax = isNumber(stats.p90) ? stats.p90 : stats.max;
      const currentPosition = comparisonTrackPosition(currentValue, trackMin, trackMax);
      const meanPosition = comparisonTrackPosition(stats.mean, trackMin, trackMax);
      const medianPosition = comparisonTrackPosition(stats.median, trackMin, trackMax);
      const hasTrack =
        isNumber(currentPosition) &&
        isNumber(meanPosition) &&
        isNumber(medianPosition) &&
        isNumber(trackMin) &&
        isNumber(trackMax);

      return `
        <div class="detail-comparison-row">
          <div class="detail-comparison-cell detail-comparison-label">
            <strong>${escapeHtml(metric.label)}</strong>
            <small>${escapeHtml(metric.description)} / 比較件数 ${formatCount(stats.count ?? 0)}件</small>
          </div>
          <div class="detail-comparison-cell">
            <span class="detail-comparison-heading">この事業所</span>
            <strong>${escapeHtml(isNumber(currentValue) ? metric.formatValue(currentValue) : metric.emptyText)}</strong>
          </div>
          <div class="detail-comparison-cell">
            <span class="detail-comparison-heading">平均との差</span>
            <strong class="detail-delta ${deltaClassName(meanDiff)}">${escapeHtml(metric.formatDiff(meanDiff))}</strong>
            <small>平均 ${escapeHtml(metric.formatValue(stats.mean))}</small>
          </div>
          <div class="detail-comparison-cell">
            <span class="detail-comparison-heading">中央値差</span>
            <strong class="detail-delta ${deltaClassName(medianDiff)}">${escapeHtml(metric.formatDiff(medianDiff))}</strong>
            <small>中央値 ${escapeHtml(metric.formatValue(stats.median))}</small>
          </div>
          <div class="detail-comparison-cell detail-deviation-cell">
            <span class="detail-comparison-heading">位置グラフ</span>
            ${
              hasTrack
                ? `
                  <div class="detail-deviation-track" aria-hidden="true">
                    <span class="detail-deviation-band"></span>
                    <span class="detail-deviation-marker is-mean" style="left:${meanPosition}%;"></span>
                    <span class="detail-deviation-marker is-median" style="left:${medianPosition}%;"></span>
                    <span class="detail-deviation-point" style="left:${currentPosition}%;"></span>
                  </div>
                  <small>左右端は比較対象の下位10%〜上位10%</small>
                `
                : `<small>${escapeHtml(metric.emptyText)}のため位置グラフを出せない。</small>`
            }
          </div>
        </div>
      `;
    })
    .join("");
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

function getAreaLabel(record) {
  const city = String(record.wam_office_address_city ?? "");
  const line = String(record.wam_office_address_line ?? "");
  const addressText = `${city}${line}`;
  const osakaWardMatch = addressText.match(/大阪市([^\s0-9]+?区)/);
  if (osakaWardMatch) {
    return osakaWardMatch[1];
  }
  if (city) {
    const normalizedCity = city.replace(/^大阪府/, "").trim();
    return normalizedCity === "大阪市" ? "大阪市（区未取得）" : normalizedCity;
  }
  return record.municipality === "大阪市" ? "大阪市（区未取得）" : record.municipality ?? null;
}

function hasWorkShortageRisk(record) {
  const wageRatio = record.wage_ratio_to_overall_mean;
  const utilization = record.daily_user_capacity_ratio;
  const lowWage = isNumber(wageRatio) && wageRatio < 0.9;
  const lowUtilization = isNumber(utilization) && utilization < 0.7;
  const veryLowUtilization = isNumber(utilization) && utilization < 0.6;
  const staffingHeavy =
    record.wam_staffing_efficiency_quadrant === "低工賃 × 厚い人員" ||
    record.wam_staffing_outlier_flag === "high";

  return (
    (lowWage && lowUtilization) ||
    (lowWage && staffingHeavy) ||
    (veryLowUtilization && staffingHeavy)
  );
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
    ["wage_ratio_to_overall_mean", "平均との差"],
    ["wage_ratio_to_municipality_mean", "市町村平均との差"],
    ["wage_ratio_to_capacity_band_mean", "定員帯平均との差"],
    ["daily_user_capacity_ratio", "利用率"],
    ["home_use_active", "在宅利用あり"],
    ["home_use_user_ratio_pct", "在宅率"],
    ["wage_outlier_flag", "工賃水準"],
    ["wam_match_status", "人員詳細"],
    ["wam_welfare_staff_fte_total", "福祉職員（常勤換算）"],
    ["wam_key_staff_fte_per_capacity", "支援職員 / 定員"],
    ["wam_transport_available", "送迎"],
    ["wam_meal_support_addon", "食事加算"],
    ["wam_manager_multi_post", "管理者兼務"],
    ["wam_staffing_efficiency_quadrant", "工賃と人員配置の分類"],
    ["wam_staffing_outlier_flag", "人員配置"],
    ["derived_work_shortage_risk", "仕事不足の可能性"],
    ["wam_primary_activity_type", "主活動種別"],
    ["wam_office_number", "公開情報の事業所番号"],
    ["wam_office_url", "WAM掲載URL"],
    ["homepage_url", "ホームページ"],
    ["instagram_url", "Instagram"],
    ["derived_website_search_url", "ホームページ検索URL"],
    ["derived_instagram_search_url", "Instagram検索URL"],
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
      if (key === "derived_work_shortage_risk") {
        return escapeCsvCell(hasWorkShortageRisk(record) ? "あり" : "なし");
      }
      if (key === "derived_website_search_url") {
        return escapeCsvCell(buildWebsiteSearchUrl(record));
      }
      if (key === "derived_instagram_search_url") {
        return escapeCsvCell(buildInstagramSearchUrl(record));
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
  const label = `人員 ${staffingLevelLabel(flag)}${severity ? ` / ${severity}` : ""}`;
  return `<span class="badge badge-${variant}">${escapeHtml(label)}</span>`;
}

function matchBadge(status, confidence) {
  if (status !== "matched") return `<span class="badge badge-unanswered">人員詳細なし</span>`;
  return `<span class="badge badge-annotated">人員詳細あり${confidence ? ` / ${escapeHtml(matchConfidenceLabel(confidence))}` : ""}</span>`;
}

function staffingLevelLabel(flag) {
  if (flag === "high") return "多め";
  if (flag === "low") return "少なめ";
  return "通常";
}

function matchConfidenceLabel(value) {
  if (value === "high") return "一致しやすい";
  if (value === "medium") return "おおむね一致";
  if (value === "low") return "要確認";
  return value ?? "-";
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
  return isNumber(value) ? `${decimalFormatter.format(value)}人分` : "-";
}

function formatIntegerUnit(value, unit) {
  return isNumber(value) ? `${formatCount(Math.round(value))}${unit}` : "-";
}

function formatSignedIntegerUnit(value, unit) {
  if (!isNumber(value)) return "-";
  return `${value >= 0 ? "+" : ""}${formatCount(Math.round(value))}${unit}`;
}

function formatDecimalUnit(value, unit) {
  return isNumber(value) ? `${decimalFormatter.format(value)}${unit}` : "-";
}

function formatSignedDecimalUnit(value, unit) {
  if (!isNumber(value)) return "-";
  return `${value >= 0 ? "+" : ""}${decimalFormatter.format(value)}${unit}`;
}

function formatNumber(value) {
  return isNumber(value) ? ratioFormatter.format(value) : "-";
}

function formatRatio(value) {
  return isNumber(value) ? `${ratioFormatter.format(value)}倍` : "-";
}

function formatPercent(value) {
  return isNumber(value) ? `${percentFormatter.format(value * 100)}%` : "-";
}

function formatSignedPercentPoint(value) {
  if (!isNumber(value)) return "-";
  return `${value >= 0 ? "+" : ""}${percentFormatter.format(value * 100)}pt`;
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

function buildWebsiteSearchUrl(record) {
  return buildSearchUrl(
    [record.office_name, record.municipality, record.corporation_name, "就労継続支援B型"]
      .filter(Boolean)
      .join(" ")
  );
}

function buildInstagramSearchUrl(record) {
  return buildSearchUrl(
    [record.office_name, record.municipality, record.corporation_name, "site:instagram.com"]
      .filter(Boolean)
      .join(" ")
  );
}

function buildSearchUrl(query) {
  return `https://www.google.com/search?q=${encodeURIComponent(query)}`;
}

function linkSourceLabel(source) {
  if (source === "wam_url") return "WAM掲載URL";
  if (source === "osaka_city_pdf_r6") return "大阪市公式PDF";
  if (source === "homepage_crawl") return "ホームページ内リンク";
  if (source === "bing_rss_search") return "Instagram検索";
  if (source === "ddg_search") return "Web検索で補完";
  return source ?? "";
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
