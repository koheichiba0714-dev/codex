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
const CORPORATION_SORT_OPTIONS = [
  { key: "municipality", label: "市区順" },
  { key: "wage", label: "工賃順" },
  { key: "utilization", label: "利用率順" },
];
const LEGAL_ENTITY_PREFIXES = [
  "株式会社",
  "有限会社",
  "合同会社",
  "一般社団法人",
  "一般財団法人",
  "公益社団法人",
  "公益財団法人",
  "社会福祉法人",
  "医療法人",
  "学校法人",
  "宗教法人",
  "特定非営利活動法人",
];
const WORK_MODEL_RULES = [
  { label: "軽作業中心", pattern: /(内職|軽作業|袋詰|梱包|検品|シール|箱折|封入|セット|組立|仕分|ピッキング)/ },
  { label: "PC・クリエイティブ", pattern: /(データ入力|web|hp|sns|デザイン|動画|ライティング|ec|印刷|チラシ|ブログ|名刺|パンフレット|pc)/ },
  { label: "飲食・食品", pattern: /(弁当|飲食|カフェ|調理|菓子|パン|喫茶|盛付|仕込み|プリン|食品)/ },
  { label: "農業・園芸", pattern: /(農業|農作業|栽培|園芸|水耕|畑|野菜|多肉植物)/ },
  { label: "ものづくり", pattern: /(製造|加工|製作|アクセサリー|ハンドメイド|縫製|木工|クラフト|ものづくり)/ },
  { label: "サービス・役務", pattern: /(清掃|洗濯|ポスティング|役務|接客|販売|受託|施設外|請負)/ },
];

/* 工賃レンジごとの内部計算テーブル */
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

/* ─────────────────────────────────────────────
   解析関数群（データサイエンス＋制度知識）
   ───────────────────────────────────────────── */

/** 工賃に応じた内部レンジを返す */
function getWageTier(wageYen) {
  if (!isNumber(wageYen)) return null;
  return WAGE_TIER_TABLE.find((t) => wageYen >= t.min && wageYen < t.max) ?? WAGE_TIER_TABLE[0];
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

/** 利用率1%改善あたりの月間収入増加額（経営者向け） */
function revenuePerUtilizationPoint(record) {
  const tier = getWageTier(record.average_wage_yen);
  const capacity = record.capacity ?? record.wam_office_capacity;
  if (!tier || !isNumber(capacity) || capacity <= 0) return null;
  // 1%改善 = 定員 × 0.01 人分の利用者増 × 単価 × 22日
  return Math.round(capacity * 0.01 * tier.unitYen * 22);
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

function sumComputed(records, selector) {
  return records.reduce((sum, record) => {
    const value = selector(record);
    return sum + (isNumber(value) ? value : 0);
  }, 0);
}

function getPrimaryActivityLabel(record) {
  return compactInlineText(record?.wam_primary_activity_type) || "主活動未取得";
}

function deriveWorkModelLabel(record) {
  const text = `${record?.wam_primary_activity_type ?? ""} ${record?.wam_primary_activity_detail ?? ""}`
    .normalize("NFKC")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
  if (!text) return "主活動未取得";
  const matched = WORK_MODEL_RULES.find((rule) => rule.pattern.test(text));
  if (matched) return matched.label;
  if (text.includes("その他")) return "その他";
  return getPrimaryActivityLabel(record);
}

function isLightWorkModel(record) {
  return deriveWorkModelLabel(record) === "軽作業中心";
}

function availableCapacity(record) {
  const capacity = record.capacity ?? record.wam_office_capacity;
  if (!isNumber(capacity) || !isNumber(record.average_daily_users)) return null;
  return Math.max(capacity - record.average_daily_users, 0);
}

function buildGroupStats(records, getLabel, metricKey, minCount = 1) {
  const groups = new Map();
  records.forEach((record) => {
    const label = getLabel(record);
    if (!label || !isNumber(record[metricKey])) return;
    const bucket = groups.get(label) ?? [];
    bucket.push(record[metricKey]);
    groups.set(label, bucket);
  });
  return [...groups.entries()]
    .map(([label, values]) => {
      const stats = computeLocalStats(values);
      return {
        label,
        count: stats.count,
        mean: stats.mean,
        median: stats.median,
        min: stats.min,
        max: stats.max,
      };
    })
    .filter((item) => (item.count ?? 0) >= minCount);
}

function groupRecords(records, getKey) {
  const groups = new Map();
  records.forEach((record) => {
    const key = getKey(record);
    if (!key) return;
    const bucket = groups.get(key) ?? [];
    bucket.push(record);
    groups.set(key, bucket);
  });
  return groups;
}

function groupedCorporations(records) {
  const groups = new Map();
  records.forEach((record) => {
    const key = corporationIdentityKey(record) ?? `office:${record.office_no}`;
    const current =
      groups.get(key) ??
      {
        key,
        label: compactInlineText(record.corporation_name) || "-",
        records: [],
      };
    current.records.push(record);
    groups.set(key, current);
  });
  return [...groups.values()].map((group) => ({
    ...group,
    label: corporationLabelForKey(group.key) || group.label,
  }));
}

function competitionAreaRecords(record, records = state.records) {
  const areaLabel = getAreaLabel(record);
  if (!areaLabel) return [];
  return records.filter((item) => getAreaLabel(item) === areaLabel);
}

function computeCompetitionProfile(record, records = state.records) {
  const areaLabel = getAreaLabel(record);
  const areaRecords = competitionAreaRecords(record, records);
  const sameActivityRecords = areaRecords.filter(
    (item) =>
      String(item.office_no) !== String(record.office_no) &&
      getPrimaryActivityLabel(item) === getPrimaryActivityLabel(record)
  );
  const areaCounts = [...groupRecords(records, (item) => getAreaLabel(item)).values()]
    .map((items) => items.length)
    .sort((left, right) => right - left);
  const areaRank = areaLabel
    ? [...groupRecords(records, (item) => getAreaLabel(item)).entries()]
        .sort((left, right) => right[1].length - left[1].length)
        .findIndex(([label]) => label === areaLabel) + 1
    : null;
  return {
    areaLabel,
    areaCount: areaRecords.length,
    sameActivityAreaCount: sameActivityRecords.length,
    areaWageMedian: computeLocalStats(numericValues(areaRecords, "average_wage_yen")).median,
    areaUtilizationMean: meanFor(areaRecords, "daily_user_capacity_ratio"),
    areaDensityMedian: areaCounts.length ? median(areaCounts) : null,
    areaRank: areaRank || null,
  };
}

function computeCapacityProfile(record, records) {
  const capacity = record.capacity ?? record.wam_office_capacity;
  const peers = records.filter(
    (item) => item.capacity_band_label && item.capacity_band_label === record.capacity_band_label
  );
  const localPeers = peers.filter((item) => item.municipality === record.municipality);
  const benchmarkPeers = localPeers.length >= 5 ? localPeers : peers;
  return {
    capacity,
    availableSlots: availableCapacity(record),
    benchmarkCount: benchmarkPeers.length,
    peerUsersMedian: computeLocalStats(numericValues(benchmarkPeers, "average_daily_users")).median,
    peerUtilizationMedian: computeLocalStats(numericValues(benchmarkPeers, "daily_user_capacity_ratio")).median,
  };
}

function computeActivityProfile(record, records) {
  const activityLabel = getPrimaryActivityLabel(record);
  const peers = records.filter((item) => getPrimaryActivityLabel(item) === activityLabel);
  const wagePeers = peers.filter((item) => isNumber(item.average_wage_yen));
  const sortedWages = wagePeers.map((item) => item.average_wage_yen).sort((left, right) => left - right);
  const rank =
    wagePeers.length && isNumber(record.average_wage_yen)
      ? wagePeers.filter((item) => item.average_wage_yen > record.average_wage_yen).length + 1
      : null;
  return {
    activityLabel,
    workModelLabel: deriveWorkModelLabel(record),
    peerCount: peers.length,
    wageMedian: computeLocalStats(sortedWages).median,
    utilizationMean: meanFor(peers, "daily_user_capacity_ratio"),
    rank,
  };
}

function computeCorporationProfile(record) {
  const corporationKey = corporationIdentityKey(record);
  const recordsInCorporation = corporationKey ? corporationRecords(corporationKey) : [record];
  const wages = numericValues(recordsInCorporation, "average_wage_yen");
  return {
    officeCount: recordsInCorporation.length,
    municipalityCount: new Set(recordsInCorporation.map((item) => item.municipality).filter(Boolean)).size,
    averageWage: meanFor(recordsInCorporation, "average_wage_yen"),
    utilizationMean: meanFor(recordsInCorporation, "daily_user_capacity_ratio"),
    wageSpread:
      wages.length >= 2 ? Math.max(...wages) - Math.min(...wages) : null,
  };
}

function computeHomeUseProfile(record, records) {
  const homeUseRecords = records.filter((item) => item.home_use_active === true);
  const nonHomeUseRecords = records.filter((item) => item.home_use_active === false);
  const topWageThreshold = percentile(
    numericValues(records, "average_wage_yen").sort((left, right) => left - right),
    0.8
  );
  const topWageRecords = records.filter((item) => isNumber(item.average_wage_yen) && item.average_wage_yen >= topWageThreshold);
  return {
    homeUseAverageWage: meanFor(homeUseRecords, "average_wage_yen"),
    nonHomeUseAverageWage: meanFor(nonHomeUseRecords, "average_wage_yen"),
    homeUseAverageUtilization: meanFor(homeUseRecords, "daily_user_capacity_ratio"),
    nonHomeUseAverageUtilization: meanFor(nonHomeUseRecords, "daily_user_capacity_ratio"),
    topWageHomeUseRate: ratioOf(topWageRecords, (item) => item.home_use_active === true),
    currentGroupAverageWage: meanFor(
      record.home_use_active === true ? homeUseRecords : nonHomeUseRecords,
      "average_wage_yen"
    ),
  };
}

function buildWorkShortageSignals(record) {
  const signals = [];
  if (isNumber(record.wage_ratio_to_overall_mean) && record.wage_ratio_to_overall_mean < 0.9) {
    signals.push("工賃が平均より低め");
  }
  if (isNumber(record.daily_user_capacity_ratio) && record.daily_user_capacity_ratio < 0.7) {
    signals.push("利用率が低い");
  }
  if (isNumber(record.home_use_user_ratio_decimal) && record.home_use_user_ratio_decimal >= 0.3) {
    signals.push("在宅率が高め");
  }
  if (isLightWorkModel(record)) {
    signals.push("軽作業中心");
  }
  if (isNumber(availableCapacity(record)) && availableCapacity(record) >= 5) {
    signals.push("定員まで空きが多い");
  }
  return signals;
}

/* ─────────────────────────────────────────────
   加算シミュレーション（令和6年度 就労B型）
   ───────────────────────────────────────────── */

const OSAKA_UNIT_YEN = 10.96;
const KASAN_WORKING_DAYS = 22;

function computeKasanSimulation(record) {
  const results = [];
  const users = record.average_daily_users;
  const capacity = record.capacity ?? record.wam_office_capacity;
  const wage = record.average_wage_yen;
  const hasUsers = isNumber(users) && users > 0;
  const hasCapacity = isNumber(capacity) && capacity > 0;

  // 1. 送迎加算(I) — 21単位/片道/人
  if (record.wam_transport_available === false) {
    const unitPerUserMonth = 21 * 2 * OSAKA_UNIT_YEN * KASAN_WORKING_DAYS;
    results.push({
      name: "送迎加算(I)",
      status: "未取得",
      monthlyPerUser: Math.round(unitPerUserMonth),
      monthlyTotal: hasUsers ? Math.round(unitPerUserMonth * users) : null,
      confidence: "high",
      source: "WAM送迎「なし」",
      note: "片道21単位×往復。利用者の送迎手段を確保すれば取得可能。車両・ドライバーのコストとの比較が必要。",
      requirement: "平均10人以上の送迎、または定員の50%以上",
    });
  } else if (record.wam_transport_available === true) {
    results.push({
      name: "送迎加算(I)",
      status: "取得済み",
      monthlyPerUser: null,
      monthlyTotal: null,
      confidence: "high",
      source: "WAM送迎「あり」",
      note: null,
      requirement: null,
    });
  }

  // 2. 食事提供体制加算 — 30単位/日/人
  if (record.wam_meal_support_addon === false) {
    const unitPerUserMonth = 30 * OSAKA_UNIT_YEN * KASAN_WORKING_DAYS;
    results.push({
      name: "食事提供体制加算",
      status: "未取得",
      monthlyPerUser: Math.round(unitPerUserMonth),
      monthlyTotal: hasUsers ? Math.round(unitPerUserMonth * users) : null,
      confidence: "high",
      source: "WAM食事提供「なし」",
      note: "食事の提供体制を整えれば取得可能。外部委託でも可。利用者負担との兼ね合いを確認。",
      requirement: "食事の提供体制を整備し、栄養士等を配置または委託",
    });
  } else if (record.wam_meal_support_addon === true) {
    results.push({
      name: "食事提供体制加算",
      status: "取得済み",
      monthlyPerUser: null,
      monthlyTotal: null,
      confidence: "high",
      source: "WAM食事提供「あり」",
      note: null,
      requirement: null,
    });
  }

  // 3. 地域協働加算 — 30単位/日/人
  if (record.wam_regional_collaboration_addon === false) {
    const unitPerUserMonth = 30 * OSAKA_UNIT_YEN * KASAN_WORKING_DAYS;
    results.push({
      name: "地域協働加算",
      status: "未取得",
      monthlyPerUser: Math.round(unitPerUserMonth),
      monthlyTotal: hasUsers ? Math.round(unitPerUserMonth * users) : null,
      confidence: "high",
      source: "WAM地域協働「なし」",
      note: "地域の企業・団体と連携した取り組みで取得可能。地域への貢献活動、販路開拓等。",
      requirement: "地域の企業等と共同受注や販路開拓の取組を実施",
    });
  } else if (record.wam_regional_collaboration_addon === true) {
    results.push({
      name: "地域協働加算",
      status: "取得済み",
      monthlyPerUser: null,
      monthlyTotal: null,
      confidence: "high",
      source: "WAM地域協働「あり」",
      note: null,
      requirement: null,
    });
  }

  // 4. 目標工賃達成指導員配置加算 — 定員20以下: 84単位/日, 21以上: 53単位/日
  if (hasCapacity && isNumber(wage)) {
    const staffRatio = record.wam_key_staff_fte_per_capacity;
    const hasGoodStaffing = isNumber(staffRatio) && staffRatio >= 0.15;
    const kasanUnits = capacity <= 20 ? 84 : 53;
    const unitPerUserMonth = kasanUnits * OSAKA_UNIT_YEN * KASAN_WORKING_DAYS;
    if (hasGoodStaffing && wage >= 10000) {
      results.push({
        name: "目標工賃達成指導員配置加算",
        status: "確認推奨",
        monthlyPerUser: Math.round(unitPerUserMonth),
        monthlyTotal: hasUsers ? Math.round(unitPerUserMonth * users) : null,
        confidence: "medium",
        source: "人員配置比率と工賃水準から推定",
        note: `人員配置 ${formatNumber(staffRatio * capacity)}人 / 定員${capacity}名（比率 ${(staffRatio * 100).toFixed(1)}%）。工賃向上計画の策定と専任指導員の配置が条件。`,
        requirement: "工賃向上計画の策定・専任の目標工賃達成指導員を1名以上配置",
      });
    }
  }

  // 5. 福祉専門職員配置等加算 — 加算(I): 15単位/日, (II): 10単位/日, (III): 6単位/日
  if (isNumber(record.wam_welfare_staff_fte_total) && hasCapacity) {
    const staffTotal = record.wam_welfare_staff_fte_total;
    const ratio = record.wam_key_staff_fte_per_capacity;
    if (isNumber(ratio)) {
      let kasanLevel = null;
      let kasanUnits = 0;
      if (ratio >= 0.35) {
        kasanLevel = "I";
        kasanUnits = 15;
      } else if (ratio >= 0.25) {
        kasanLevel = "II";
        kasanUnits = 10;
      } else if (ratio >= 0.15) {
        kasanLevel = "III";
        kasanUnits = 6;
      }
      if (kasanLevel) {
        const unitPerUserMonth = kasanUnits * OSAKA_UNIT_YEN * KASAN_WORKING_DAYS;
        results.push({
          name: `福祉専門職員配置等加算(${kasanLevel})`,
          status: "確認推奨",
          monthlyPerUser: Math.round(unitPerUserMonth),
          monthlyTotal: hasUsers ? Math.round(unitPerUserMonth * users) : null,
          confidence: "medium",
          source: `職員FTE ${formatNumber(staffTotal)}人 / 配置比率 ${(ratio * 100).toFixed(1)}%`,
          note: "社会福祉士・介護福祉士等の有資格者配置率から推定。実際の有資格者数の確認が必要。",
          requirement: kasanLevel === "I" ? "有資格者が常勤換算で職員の35%以上" : kasanLevel === "II" ? "有資格者が常勤換算で職員の25%以上" : "常勤率75%以上かつ勤続3年以上30%以上",
        });
      }
    }
  }

  // 6. 在宅時生活支援サービス加算 — 300単位/日
  if (record.home_use_active === true && isNumber(record.home_use_user_ratio_decimal) && record.home_use_user_ratio_decimal > 0) {
    const homeUsers = hasUsers ? Math.round(users * record.home_use_user_ratio_decimal) : null;
    results.push({
      name: "在宅時生活支援サービス加算",
      status: "確認推奨",
      monthlyPerUser: Math.round(300 * OSAKA_UNIT_YEN * KASAN_WORKING_DAYS),
      monthlyTotal: homeUsers ? Math.round(300 * OSAKA_UNIT_YEN * KASAN_WORKING_DAYS * homeUsers) : null,
      confidence: "medium",
      source: `在宅利用あり（在宅率 ${(record.home_use_user_ratio_decimal * 100).toFixed(1)}%）`,
      note: `在宅利用者 推定${homeUsers ?? "-"}名。在宅利用者への生活支援を実施していれば取得可能。`,
      requirement: "在宅利用者に対する生活支援（電話・訪問等）を月2回以上実施",
    });
  }

  // 集計
  const opportunities = results.filter((r) => r.status !== "取得済み");
  const totalMonthlyOpportunity = opportunities.reduce((sum, r) => sum + (r.monthlyTotal ?? 0), 0);
  const acquired = results.filter((r) => r.status === "取得済み");

  return {
    items: results,
    opportunities,
    acquired,
    totalMonthlyOpportunity,
    totalAnnualOpportunity: totalMonthlyOpportunity * 12,
  };
}

/* ─────────────────────────────────────────────
   利用者の流れ分析
   ───────────────────────────────────────────── */

function computeUserFlowAnalysis(record, records) {
  const capacity = record.capacity ?? record.wam_office_capacity;
  const users = record.average_daily_users;
  const utilization = record.daily_user_capacity_ratio;
  const tier = getWageTier(record.average_wage_yen);
  const areaLabel = getAreaLabel(record);

  // 機会損失の算出
  let opportunityCost = null;
  if (isNumber(capacity) && isNumber(users) && tier && users < capacity) {
    const emptySlots = capacity - users;
    const dailyRevPerUser = tier.unitYen;
    opportunityCost = {
      emptySlots: Math.round(emptySlots * 10) / 10,
      monthlyLoss: Math.round(emptySlots * dailyRevPerUser * KASAN_WORKING_DAYS),
      annualLoss: Math.round(emptySlots * dailyRevPerUser * KASAN_WORKING_DAYS * 12),
      dailyRevPerUser,
    };
  }

  // 同エリア競合分析
  const areaRecords = areaLabel ? records.filter((r) => getAreaLabel(r) === areaLabel) : [];
  const areaWithUtilization = areaRecords.filter((r) => isNumber(r.daily_user_capacity_ratio) && isNumber(r.average_wage_yen));
  const areaHighUtil = areaWithUtilization
    .filter((r) => r.daily_user_capacity_ratio >= 0.85 && String(r.office_no) !== String(record.office_no))
    .sort((a, b) => b.daily_user_capacity_ratio - a.daily_user_capacity_ratio);

  // 高利用率事業所の共通パターン
  const highUtilTop = areaHighUtil.slice(0, 10);
  const patterns = {
    total: highUtilTop.length,
    withHomepage: highUtilTop.filter((r) => r.homepage_url || r.wam_office_url).length,
    withTransport: highUtilTop.filter((r) => r.wam_transport_available === true).length,
    withHomeUse: highUtilTop.filter((r) => r.home_use_active === true).length,
    withMeal: highUtilTop.filter((r) => r.wam_meal_support_addon === true).length,
    avgWage: meanFor(highUtilTop, "average_wage_yen"),
    avgCapacity: meanFor(highUtilTop, "capacity"),
  };

  // この事業所にないが、高利用率事業所が持っている特徴
  const advantages = [];
  if (!record.homepage_url && !record.wam_office_url && patterns.total > 0 && patterns.withHomepage / patterns.total >= 0.5) {
    advantages.push({ feature: "ホームページ", detail: `高利用率事業所の${Math.round(patterns.withHomepage / patterns.total * 100)}%が保有`, impact: "high" });
  }
  if (record.wam_transport_available !== true && patterns.total > 0 && patterns.withTransport / patterns.total >= 0.5) {
    advantages.push({ feature: "送迎", detail: `高利用率事業所の${Math.round(patterns.withTransport / patterns.total * 100)}%が実施`, impact: "high" });
  }
  if (record.home_use_active !== true && patterns.total > 0 && patterns.withHomeUse / patterns.total >= 0.5) {
    advantages.push({ feature: "在宅利用", detail: `高利用率事業所の${Math.round(patterns.withHomeUse / patterns.total * 100)}%が対応`, impact: "medium" });
  }
  if (record.wam_meal_support_addon !== true && patterns.total > 0 && patterns.withMeal / patterns.total >= 0.5) {
    advantages.push({ feature: "食事提供", detail: `高利用率事業所の${Math.round(patterns.withMeal / patterns.total * 100)}%が実施`, impact: "medium" });
  }

  // ベンチマーク先（同エリアの上位3事業所）
  const benchmarks = areaHighUtil.slice(0, 3).map((r) => ({
    officeName: r.office_name,
    officeNo: r.office_no,
    wage: r.average_wage_yen,
    utilization: r.daily_user_capacity_ratio,
    hasHomepage: !!(r.homepage_url || r.wam_office_url),
    hasTransport: r.wam_transport_available === true,
    hasHomeUse: r.home_use_active === true,
    activity: r.wam_primary_activity_type,
  }));

  // 利用率改善のインパクト
  const utilizationGap = isNumber(utilization) ? Math.max(0, 1.0 - utilization) : null;
  const potentialUsers = isNumber(utilizationGap) && isNumber(capacity) ? capacity * utilizationGap : null;

  return {
    opportunityCost,
    areaLabel,
    areaTotal: areaRecords.length,
    areaHighUtilCount: areaHighUtil.length,
    patterns,
    advantages,
    benchmarks,
    utilizationGap,
    potentialUsers: potentialUsers ? Math.round(potentialUsers * 10) / 10 : null,
  };
}

function buildMarketCards(records) {
  const areaGroups = [...groupRecords(records, (record) => getAreaLabel(record)).entries()]
    .map(([label, items]) => ({
      label,
      count: items.length,
      medianWage: computeLocalStats(numericValues(items, "average_wage_yen")).median,
      utilizationMean: meanFor(items, "daily_user_capacity_ratio"),
    }))
    .filter((item) => item.label)
    .sort((left, right) => right.count - left.count);
  const areaCountMedian = areaGroups.length ? median(areaGroups.map((item) => item.count).sort((a, b) => a - b)) : null;
  const overallWageMedian = computeLocalStats(numericValues(records, "average_wage_yen")).median;
  const overallUtilizationMean = meanFor(records, "daily_user_capacity_ratio");
  const topAreas = areaGroups.slice(0, 3);
  const opportunityArea = areaGroups
    .filter(
      (item) =>
        item.count >= 3 &&
        isNumber(item.medianWage) &&
        isNumber(item.utilizationMean) &&
        isNumber(areaCountMedian) &&
        item.count <= areaCountMedian &&
        item.medianWage >= overallWageMedian &&
        item.utilizationMean >= overallUtilizationMean
    )
    .sort((left, right) => right.medianWage + right.utilizationMean * 10000 - (left.medianWage + left.utilizationMean * 10000))[0];

  const activityStats = buildGroupStats(records, (record) => deriveWorkModelLabel(record), "average_wage_yen", 5)
    .sort((left, right) => (right.median ?? 0) - (left.median ?? 0));

  const shortageRecords = records.filter((record) => hasWorkShortageRisk(record));
  const shortageLightWorkCount = shortageRecords.filter((record) => isLightWorkModel(record)).length;
  const shortageHighHomeUseCount = shortageRecords.filter(
    (record) => isNumber(record.home_use_user_ratio_decimal) && record.home_use_user_ratio_decimal >= 0.3
  ).length;
  const shortageOpenSlots = sumComputed(shortageRecords, (record) => availableCapacity(record));

  const homeUseRecords = records.filter((record) => record.home_use_active === true);
  const nonHomeUseRecords = records.filter((record) => record.home_use_active === false);
  const highWageThreshold = percentile(
    numericValues(records, "average_wage_yen").sort((left, right) => left - right),
    0.8
  );
  const topWageRecords = records.filter((record) => isNumber(record.average_wage_yen) && record.average_wage_yen >= highWageThreshold);

  const corporationGroups = groupedCorporations(records);
  const multiOfficeGroups = corporationGroups.filter((group) => group.records.length >= 2);
  const largestCorporation = [...corporationGroups].sort((left, right) => right.records.length - left.records.length)[0];
  const averageCorporationSpread = computeLocalStats(
    multiOfficeGroups
      .map((group) => {
        const wages = numericValues(group.records, "average_wage_yen");
        return wages.length >= 2 ? Math.max(...wages) - Math.min(...wages) : null;
      })
      .filter(isNumber)
  ).mean;

  const averageOpenSlots = computeLocalStats(records.map((record) => availableCapacity(record)).filter(isNumber)).mean;

  return [
    {
      title: "競合密度",
      summary: topAreas.length
        ? `${topAreas.map((item) => `${item.label} ${formatCount(item.count)}件`).join(" / ")} が表示中の密集エリア。`
        : "エリア比較に必要な住所情報が不足している。",
      bullets: [
        opportunityArea
          ? `${opportunityArea.label} は ${formatCount(opportunityArea.count)}件と競合が薄めだが、中央値工賃 ${formatMaybeYen(opportunityArea.medianWage)}、平均利用率 ${formatPercent(opportunityArea.utilizationMean)}。`
          : "競合が薄くて強いエリアは、今の絞り込みでははっきり出ていない。",
        "詳細を開くと、各事業所ごとに同じエリア・同じ主活動の件数を確認できる。",
      ],
    },
    {
      title: "定員充足のしやすさ",
      summary: isNumber(averageOpenSlots)
        ? `表示中では定員まで平均 ${formatDecimalUnit(averageOpenSlots, "人")} の空きがある。`
        : "定員充足の比較に必要な定員または平均利用人数が不足している。",
      bullets: [
        `${formatCount(records.filter((record) => isNumber(availableCapacity(record)) && availableCapacity(record) >= 5).length)}件で、定員まで平均5人以上の空きが残る。`,
        "詳細では、同じ定員帯の中央値と比べて何人分の差があるかを表示する。",
      ],
    },
    {
      title: "主活動の勝ち筋",
      summary: activityStats.length
        ? `${activityStats
            .slice(0, 3)
            .map((item) => `${item.label} ${formatMaybeYen(item.median)}`)
            .join(" / ")} が高工賃側。`
        : "主活動の比較に必要な公開情報が不足している。",
      bullets: [
        activityStats[0]
          ? `${activityStats[0].label} は ${formatCount(activityStats[0].count)}件で中央値 ${formatMaybeYen(activityStats[0].median)}。`
          : "十分な件数がある主活動がまだ少ない。",
        "詳細では、主活動ごとの順位と中央値差を確認できる。",
      ],
    },
    {
      title: "仕事不足の実態",
      summary: `${formatCount(shortageRecords.length)}件を仕事不足の確認候補として抽出。`,
      bullets: [
        `そのうち ${formatCount(shortageLightWorkCount)}件が軽作業中心、${formatCount(shortageHighHomeUseCount)}件が在宅率30%以上。`,
        isNumber(shortageOpenSlots)
          ? `候補全体で定員まで合計 ${formatCount(Math.round(shortageOpenSlots))}人分の空きがある。`
          : "空き人数の比較に必要な値が不足している。",
      ],
    },
    {
      title: "在宅の使い方",
      summary: homeUseRecords.length && nonHomeUseRecords.length
        ? `在宅ありの平均工賃 ${formatMaybeYen(meanFor(homeUseRecords, "average_wage_yen"))}、なしは ${formatMaybeYen(meanFor(nonHomeUseRecords, "average_wage_yen"))}。`
        : "在宅あり/なしの比較に十分な件数がない。",
      bullets: [
        `在宅ありの平均利用率 ${formatPercent(meanFor(homeUseRecords, "daily_user_capacity_ratio"))} / なし ${formatPercent(meanFor(nonHomeUseRecords, "daily_user_capacity_ratio"))}。`,
        `工賃上位20%では ${formatPercent(ratioOf(topWageRecords, (record) => record.home_use_active === true))} が在宅あり。`,
      ],
    },
    {
      title: "法人展開の傾向",
      summary: `${formatCount(multiOfficeGroups.length)}法人が2拠点以上で展開。`,
      bullets: [
        largestCorporation
          ? `最大は ${largestCorporation.label} の ${formatCount(largestCorporation.records.length)}事業所。`
          : "複数拠点の法人はまだない。",
        isNumber(averageCorporationSpread)
          ? `複数拠点法人の工賃差は平均 ${formatMaybeYen(averageCorporationSpread)}。法人内でも差が出ている。`
          : "法人内の工賃差を比べるだけの件数が足りない。",
      ],
    },
  ];
}

/** 工賃の標準偏差を計算 */
function computeStdDev(values) {
  if (values.length < 2) return null;
  const mean = values.reduce((s, v) => s + v, 0) / values.length;
  const variance = values.reduce((s, v) => s + (v - mean) ** 2, 0) / (values.length - 1);
  return Math.sqrt(variance);
}

const FILTER_PRESETS = {
  all: {},
  "high-wage": {
    outlierFlag: "high",
  },
  "high-wage-low-util": {
    quadrant: "高工賃 × 低稼働",
  },
  "fix-priority": {
    quadrant: "低工賃 × 低稼働",
    workShortageRisk: "likely",
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
const INITIAL_PRESET = "all";
const KPI_PERIOD_LABEL = "令和6年度実績";
const KPI_BASELINE_LABEL = "大阪市全体";

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
    workShortageRisk: "all",
    primaryActivity: "all",
    workModel: "all",
    newOnly: false,
    homeUseOnly: false,
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
  currentCompareView: "table",
  filters: createPresetFilters(INITIAL_PRESET),
  draftFilters: createPresetFilters(INITIAL_PRESET),
  selectedOfficeNo: "339",
  selectedCorporationKey: null,
  selectedCorporationLabel: null,
  selectedCorporationSort: "municipality",
  selectedRepresentativeKey: null,
  selectedRepresentativeLabel: null,
  activePreset: INITIAL_PRESET,
  corporationGroupIndex: null,
  /* User view state */
  currentView: "user",
  userSearchQuery: "",
  userAreaFilter: "all",
  userSort: "wage-desc",
  userPage: 1,
  userPageSize: 24,
};

/* ===== User View (利用者向けビュー) ===== */

function switchView(viewName) {
  state.currentView = viewName;
  document.body.classList.remove("view-user", "view-operator");
  document.body.classList.add(`view-${viewName}`);

  document.querySelectorAll(".view-toggle-tab").forEach((tab) => {
    const isActive = tab.dataset.view === viewName;
    tab.classList.toggle("is-active", isActive);
    tab.setAttribute("aria-selected", String(isActive));
  });

  if (viewName === "user") {
    renderUserView();
  }
}

function getUserViewRecords() {
  let records = state.records.filter((r) => r.response_status !== "unanswered");

  if (state.userSearchQuery) {
    const q = state.userSearchQuery.toLowerCase();
    records = records.filter((r) => {
      const haystack = [
        r.office_name, r.corporation_name, r.municipality,
        getAreaLabel(r), r.wam_office_address_city, r.wam_office_address_line,
        composeAddress(r), r.wam_primary_activity_type, r.wam_primary_activity_detail,
      ].filter(Boolean).join(" ").toLowerCase();
      return haystack.includes(q);
    });
  }

  if (state.userAreaFilter !== "all") {
    records = records.filter((r) => getAreaLabel(r) === state.userAreaFilter);
  }

  switch (state.userSort) {
    case "wage-desc":
      records.sort((a, b) => (b.average_wage_yen ?? -1) - (a.average_wage_yen ?? -1));
      break;
    case "wage-asc":
      records.sort((a, b) => (a.average_wage_yen ?? Infinity) - (b.average_wage_yen ?? Infinity));
      break;
    case "utilization-desc":
      records.sort((a, b) => (b.daily_user_capacity_ratio ?? -1) - (a.daily_user_capacity_ratio ?? -1));
      break;
    case "name-asc":
      records.sort((a, b) => (a.office_name ?? "").localeCompare(b.office_name ?? "", "ja"));
      break;
  }
  return records;
}

function renderUserView() {
  if (state.currentView !== "user") return;
  const records = getUserViewRecords();
  const totalPages = Math.max(1, Math.ceil(records.length / state.userPageSize));
  state.userPage = Math.min(state.userPage, totalPages);
  const start = (state.userPage - 1) * state.userPageSize;
  const page = records.slice(start, start + state.userPageSize);

  const summaryEl = document.getElementById("userViewSummary");
  if (summaryEl) {
    summaryEl.textContent = `${formatCount(records.length)}件の事業所が見つかりました` +
      (state.userAreaFilter !== "all" ? `（${state.userAreaFilter}）` : "");
  }

  const container = document.getElementById("userCardList");
  if (!container) return;

  container.innerHTML = page.map((r, idx) => {
    const globalRank = start + idx + 1;
    const wage = r.average_wage_yen;
    const hasWage = isNumber(wage);
    const area = getAreaLabel(r) ?? r.municipality ?? "";
    const activity = r.wam_primary_activity_type ?? "";
    const detail = r.wam_primary_activity_detail ?? "";
    const hasTransport = r.wam_transport_available;
    const hasHome = r.home_use_active;
    const hasHp = Boolean(safeExternalUrl(r.homepage_url));
    const utilRatio = r.daily_user_capacity_ratio;
    const capacity = r.capacity;
    const address = composeAddress(r);

    let rankClass = "";
    if (state.userSort === "wage-desc" || state.userSort === "wage-asc") {
      if (globalRank === 1) rankClass = "user-card-rank-gold";
      else if (globalRank === 2) rankClass = "user-card-rank-silver";
      else if (globalRank === 3) rankClass = "user-card-rank-bronze";
    }

    const tags = [];
    if (area) tags.push(`<span class="user-card-tag user-card-tag-area">${escapeHtml(area)}</span>`);
    if (activity) tags.push(`<span class="user-card-tag user-card-tag-activity">${escapeHtml(activity)}</span>`);
    if (hasTransport) tags.push(`<span class="user-card-tag user-card-tag-transport">送迎あり</span>`);
    if (hasHome) tags.push(`<span class="user-card-tag user-card-tag-home">在宅利用可</span>`);

    const stats = [];
    if (isNumber(utilRatio)) stats.push(`<span class="user-card-stat">利用率 <strong>${percentFormatter.format(utilRatio * 100)}%</strong></span>`);
    if (isNumber(capacity)) stats.push(`<span class="user-card-stat">定員 <strong>${capacity}名</strong></span>`);

    const links = [];
    if (hasHp) links.push(`<a href="${escapeAttribute(safeExternalUrl(r.homepage_url))}" target="_blank" rel="noopener" class="user-card-link" onclick="event.stopPropagation()">ホームページ</a>`);

    return `
      <article class="user-facility-card" data-office-no="${escapeAttribute(r.office_no)}">
        <div class="user-card-header">
          <div style="display:flex;align-items:center;gap:8px;min-width:0;">
            ${rankClass ? `<span class="user-card-rank ${rankClass}">${globalRank}</span>` : ""}
            <h3 class="user-card-name">${escapeHtml(r.office_name ?? "名称不明")}</h3>
          </div>
          <div class="user-card-wage">
            ${hasWage
              ? `<div class="user-card-wage-amount">${formatCount(Math.round(wage))}<span style="font-size:13px;font-weight:600;">円</span></div>
                 <div class="user-card-wage-label">月額平均工賃</div>`
              : `<div class="user-card-no-wage">工賃非公開</div>`}
          </div>
        </div>
        ${tags.length ? `<div class="user-card-meta">${tags.join("")}</div>` : ""}
        ${detail ? `<p class="user-card-description">${escapeHtml(detail)}</p>` : ""}
        ${address ? `<p class="user-card-description" style="font-size:12px;">${escapeHtml(address)}</p>` : ""}
        ${stats.length ? `<div class="user-card-stats">${stats.join("")}</div>` : ""}
        ${links.length ? `<div class="user-card-links">${links.join("")}</div>` : ""}
      </article>`;
  }).join("");

  /* Card click → open detail dialog */
  container.querySelectorAll(".user-facility-card").forEach((card) => {
    card.addEventListener("click", () => {
      const officeNo = card.dataset.officeNo;
      const record = state.records.find((r) => String(r.office_no) === String(officeNo));
      if (record) {
        state.selectedOfficeNo = record.office_no;
        renderDetail(record);
        openDialogElement(document.getElementById("detailDialog"));
      }
    });
  });

  /* Pagination */
  const pageSummary = document.getElementById("userPageSummary");
  if (pageSummary) pageSummary.textContent = `${state.userPage} / ${totalPages}`;
  const prevBtn = document.getElementById("userPrevPage");
  const nextBtn = document.getElementById("userNextPage");
  if (prevBtn) prevBtn.disabled = state.userPage <= 1;
  if (nextBtn) nextBtn.disabled = state.userPage >= totalPages;
}

function populateUserAreaFilter() {
  const select = document.getElementById("userAreaFilter");
  if (!select) return;
  const areas = new Map();
  state.records.forEach((r) => {
    const area = getAreaLabel(r);
    if (area) areas.set(area, (areas.get(area) ?? 0) + 1);
  });
  const sorted = [...areas.entries()].sort((a, b) => b[1] - a[1]);
  select.innerHTML = `<option value="all">すべてのエリア（${formatCount(state.records.length)}件）</option>` +
    sorted.map(([area, count]) => `<option value="${escapeAttribute(area)}">${escapeHtml(area)}（${count}件）</option>`).join("");
}

function bindUserView() {
  /* View toggle tabs */
  document.querySelectorAll(".view-toggle-tab").forEach((tab) => {
    tab.addEventListener("click", () => switchView(tab.dataset.view));
  });

  /* Search */
  const searchInput = document.getElementById("userSearchInput");
  if (searchInput) {
    let debounceTimer;
    searchInput.addEventListener("input", () => {
      clearTimeout(debounceTimer);
      debounceTimer = setTimeout(() => {
        state.userSearchQuery = searchInput.value.trim();
        state.userPage = 1;
        renderUserView();
      }, 250);
    });
  }

  /* Area filter */
  const areaFilter = document.getElementById("userAreaFilter");
  if (areaFilter) {
    areaFilter.addEventListener("change", () => {
      state.userAreaFilter = areaFilter.value;
      state.userPage = 1;
      renderUserView();
    });
  }

  /* Sort */
  const sortSelect = document.getElementById("userSortSelect");
  if (sortSelect) {
    sortSelect.addEventListener("change", () => {
      state.userSort = sortSelect.value;
      state.userPage = 1;
      renderUserView();
    });
  }

  /* Pagination */
  document.getElementById("userPrevPage")?.addEventListener("click", () => {
    if (state.userPage > 1) { state.userPage--; renderUserView(); }
  });
  document.getElementById("userNextPage")?.addEventListener("click", () => {
    const totalPages = Math.max(1, Math.ceil(getUserViewRecords().length / state.userPageSize));
    if (state.userPage < totalPages) { state.userPage++; renderUserView(); }
  });
}

function switchCompareView(viewName) {
  state.currentCompareView = viewName === "priority" ? "priority" : "table";

  document.querySelectorAll("[data-compare-view]").forEach((button) => {
    const isActive = button.getAttribute("data-compare-view") === state.currentCompareView;
    button.classList.toggle("is-active", isActive);
    button.setAttribute("aria-selected", String(isActive));
  });

  document.querySelectorAll("[data-compare-panel]").forEach((panel) => {
    const isActive = panel.getAttribute("data-compare-panel") === state.currentCompareView;
    panel.classList.toggle("is-active", isActive);
    panel.hidden = !isActive;
  });

  renderCompareSummary(state.filteredRecords);
}

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
  bindRepresentativeDialog();
  bindPanelToggles();
  bindMobileSidebar();
  bindUserView();
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
  populateUserAreaFilter();
  bindEvents();
  syncPresetButtons();
  applyFilters();
  switchCompareView(state.currentCompareView);
  switchView("user");
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

function bindRepresentativeDialog() {
  const dialog = document.getElementById("representativeDialog");
  if (!dialog) return;

  document.getElementById("closeRepresentativeButton")?.addEventListener("click", () => {
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

function openCorporationDialog(officeNo) {
  const record = findRecordByOffice(officeNo);
  if (!record) return;
  const corporationKey = corporationIdentityKey(record);
  if (!corporationKey) return;
  renderCorporationDialog(corporationKey);
  closeDialogElement(document.getElementById("detailDialog"));
  closeDialogElement(document.getElementById("representativeDialog"));
  openDialogElement(document.getElementById("corporationDialog"));
}

function openRepresentativeDialog(representativeName) {
  const label = String(representativeName ?? "").trim();
  if (!label) return;
  renderRepresentativeDialog(label);
  closeDialogElement(document.getElementById("detailDialog"));
  closeDialogElement(document.getElementById("corporationDialog"));
  openDialogElement(document.getElementById("representativeDialog"));
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
  bindInput("workShortageRiskSelect", "change", (filters, value) => {
    filters.workShortageRisk = value;
  });
  bindInput("primaryActivitySelect", "change", (filters, value) => {
    filters.primaryActivity = value;
  });
  bindInput("workModelSelect", "change", (filters, value) => {
    filters.workModel = value;
  });
  bindCheckbox("newOnlyCheckbox", "newOnly");
  bindCheckbox("homeUseOnlyCheckbox", "homeUseOnly");

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

  document.querySelectorAll("[data-compare-view]").forEach((button) => {
    button.addEventListener("click", () => {
      const viewName = button.getAttribute("data-compare-view");
      if (!viewName) return;
      switchCompareView(viewName);
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
      if (sortKey && state.selectedCorporationKey) {
        state.selectedCorporationSort = sortKey;
        renderCorporationDialog(state.selectedCorporationKey);
      }
      return;
    }

    const corporationTrigger = event.target.closest("[data-open-corporation-office]");
    if (corporationTrigger) {
      const officeNo = corporationTrigger.getAttribute("data-open-corporation-office");
      if (!officeNo) return;
      openCorporationDialog(officeNo);
      return;
    }

    const representativeTrigger = event.target.closest("[data-open-representative]");
    if (representativeTrigger) {
      const representativeName = representativeTrigger.getAttribute("data-open-representative");
      if (!representativeName) return;
      openRepresentativeDialog(representativeName);
      return;
    }

    const trigger = event.target.closest("[data-select-office]");
    if (!trigger) return;
    const officeNo = trigger.getAttribute("data-select-office");
    if (!officeNo) return;
    if (trigger.closest("#corporationDialog") || trigger.closest("#representativeDialog")) {
      closeDialogElement(document.getElementById("corporationDialog"));
      closeDialogElement(document.getElementById("representativeDialog"));
    }
    selectRecord(officeNo, { openDetail: true });
  });
}

function bindInput(id, eventName, setter) {
  const element = document.getElementById(id);
  if (!element) return;
  element.addEventListener(eventName, (event) => {
    const draftFilters = state.draftFilters ?? state.filters;
    setter(draftFilters, event.target.value);
  });
}

function bindCheckbox(id, key) {
  const element = document.getElementById(id);
  if (!element) return;
  element.addEventListener("change", (event) => {
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
  const setValue = (id, value) => {
    const element = document.getElementById(id);
    if (element) element.value = value;
  };
  const setChecked = (id, value) => {
    const element = document.getElementById(id);
    if (element) element.checked = value;
  };

  setValue("searchInput", filters.search);
  setValue("municipalitySelect", filters.municipality);
  setValue("areaSelect", filters.area);
  setValue("corporationTypeSelect", filters.corporationType);
  setValue("responseStatusSelect", filters.responseStatus);
  setValue("outlierFlagSelect", filters.outlierFlag);
  setValue("capacityBandSelect", filters.capacityBand);
  setValue("quadrantSelect", filters.quadrant);
  setValue("workShortageRiskSelect", filters.workShortageRisk);
  setValue("primaryActivitySelect", filters.primaryActivity);
  setValue("workModelSelect", filters.workModel);
  setChecked("newOnlyCheckbox", filters.newOnly);
  setChecked("homeUseOnlyCheckbox", filters.homeUseOnly);
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

function findRecordByOffice(officeNo) {
  if (officeNo == null) return null;
  return (
    state.filteredRecords.find((record) => String(record.office_no) === String(officeNo)) ??
    state.records.find((record) => String(record.office_no) === String(officeNo)) ??
    null
  );
}

function compactInlineText(value) {
  return String(value ?? "").replace(/\s+/g, " ").trim();
}

function normalizeCorporationName(value) {
  return String(value ?? "")
    .normalize("NFKC")
    .trim()
    .toLowerCase()
    .replace(/^(㈱|\(株\)|（株）)/, "株式会社")
    .replace(/^(㈲|\(有\)|（有）)/, "有限会社")
    .replace(/^npo法人/, "特定非営利活動法人")
    .replace(/^\(一社\)|^（一社）/, "一般社団法人")
    .replace(/^\(社福\)|^（社福）/, "社会福祉法人")
    .replace(/[‐‑‒–—―ｰ－-]/g, "ー")
    .replace(/[・･·]/g, "")
    .replace(/\s+/g, "");
}

function hasLegalEntityPrefix(value) {
  const key = normalizeCorporationName(value);
  return LEGAL_ENTITY_PREFIXES.some((prefix) => key.startsWith(prefix));
}

function stripLegalEntityPrefix(value) {
  const key = typeof value === "string" ? value : normalizeCorporationName(value);
  const prefix = LEGAL_ENTITY_PREFIXES.find((item) => key.startsWith(item));
  return prefix ? key.slice(prefix.length) : key;
}

function normalizeRepresentativeName(value) {
  let text = String(value ?? "")
    .normalize("NFKC")
    .trim()
    .toLowerCase()
    .replace(/\s*[（(][^()（）]*(職場情報総合サイト|geps)[^()（）]*[)）]\s*$/g, "");
  const prefixes = [
    "代表取締役社長",
    "代表取締役会長",
    "代表取締役",
    "取締役社長",
    "代表社員",
    "代表理事",
    "理事長",
    "院長",
    "会長",
    "社長",
    "代表",
    "理事",
  ];
  while (text) {
    text = text.replace(/^[・･\s]+/g, "");
    const prefix = prefixes.find((item) => text.startsWith(item));
    if (!prefix) break;
    text = text.slice(prefix.length).trim();
  }
  return text.replace(/[・･·]/g, "").replace(/\s+/g, "");
}

function getCorporationKey(value) {
  const key = normalizeCorporationName(value);
  return key || null;
}

function getCorporationBaseKey(value) {
  const key = getCorporationKey(value);
  if (!key) return null;
  return stripLegalEntityPrefix(key) || key;
}

function getRepresentativeKey(value) {
  const key = normalizeRepresentativeName(value);
  return key || null;
}

function preferredCorporationRecord(records) {
  if (!records.length) return null;
  return [...records].sort((left, right) => {
    const leftLabel = compactInlineText(left.corporation_name);
    const rightLabel = compactInlineText(right.corporation_name);
    const leftScore =
      (hasLegalEntityPrefix(leftLabel) ? 100 : 0) +
      (String(left.corporation_number ?? "").trim() ? 20 : 0) +
      leftLabel.length;
    const rightScore =
      (hasLegalEntityPrefix(rightLabel) ? 100 : 0) +
      (String(right.corporation_number ?? "").trim() ? 20 : 0) +
      rightLabel.length;
    if (rightScore !== leftScore) return rightScore - leftScore;
    return rightLabel.localeCompare(leftLabel, "ja");
  })[0];
}

function buildCorporationGroupIndex() {
  if (state.corporationGroupIndex?.recordsRef === state.records) {
    return state.corporationGroupIndex;
  }

  const representativeBaseGroups = new Map();
  state.records.forEach((record) => {
    const representativeKey = getRepresentativeKey(record.representative_name);
    const baseKey = getCorporationBaseKey(record.corporation_name);
    const corporationNumber = compactInlineText(record.corporation_number);
    if (!representativeKey || !baseKey) return;
    const clusterKey = `${representativeKey}::${baseKey}`;
    const cluster =
      representativeBaseGroups.get(clusterKey) ??
      {
        numbers: new Set(),
      };
    if (corporationNumber) {
      cluster.numbers.add(corporationNumber);
    }
    representativeBaseGroups.set(clusterKey, cluster);
  });

  const corporationNameGroups = new Map();
  state.records.forEach((record) => {
    const corporationKey = getCorporationKey(record.corporation_name);
    const corporationNumber = compactInlineText(record.corporation_number);
    if (!corporationKey) return;
    const representativeKey = getRepresentativeKey(record.representative_name);
    const baseKey = getCorporationBaseKey(record.corporation_name);
    const clusterKey = representativeKey && baseKey ? `${representativeKey}::${baseKey}` : null;
    const inferredByRepresentative =
      !corporationNumber && clusterKey && representativeBaseGroups.get(clusterKey)?.numbers.size === 1
        ? [...representativeBaseGroups.get(clusterKey).numbers][0]
        : null;
    const cluster =
      corporationNameGroups.get(corporationKey) ??
      {
        numbers: new Set(),
      };
    if (corporationNumber || inferredByRepresentative) {
      cluster.numbers.add(corporationNumber || inferredByRepresentative);
    }
    corporationNameGroups.set(corporationKey, cluster);
  });

  const keyByOffice = new Map();
  const recordsByKey = new Map();
  const labelByKey = new Map();

  state.records.forEach((record) => {
    const officeKey = String(record.office_no ?? "").trim();
    const corporationNumber = compactInlineText(record.corporation_number);
    const representativeKey = getRepresentativeKey(record.representative_name);
    const baseKey = getCorporationBaseKey(record.corporation_name);
    const corporationNameKey = getCorporationKey(record.corporation_name);
    const clusterKey = representativeKey && baseKey ? `${representativeKey}::${baseKey}` : null;
    const inferredByExactName =
      !corporationNumber && corporationNameKey && corporationNameGroups.get(corporationNameKey)?.numbers.size === 1
        ? [...corporationNameGroups.get(corporationNameKey).numbers][0]
        : null;
    const inferredCorporationNumber =
      inferredByExactName ||
      (!corporationNumber && clusterKey && representativeBaseGroups.get(clusterKey)?.numbers.size === 1
        ? [...representativeBaseGroups.get(clusterKey).numbers][0]
        : null);
    const corporationKey =
      (corporationNumber && `corp:${corporationNumber}`) ||
      (inferredCorporationNumber && `corp:${inferredCorporationNumber}`) ||
      (getCorporationKey(record.corporation_name) && `name:${getCorporationKey(record.corporation_name)}`) ||
      `office:${officeKey}`;

    keyByOffice.set(officeKey, corporationKey);
    const list = recordsByKey.get(corporationKey) ?? [];
    list.push(record);
    recordsByKey.set(corporationKey, list);
  });

  recordsByKey.forEach((records, key) => {
    labelByKey.set(key, compactInlineText(preferredCorporationRecord(records)?.corporation_name) || "-");
  });

  state.corporationGroupIndex = {
    recordsRef: state.records,
    keyByOffice,
    recordsByKey,
    labelByKey,
  };
  return state.corporationGroupIndex;
}

function corporationRecords(corporationKey) {
  if (!corporationKey) return [];
  const index = buildCorporationGroupIndex();
  return index.recordsByKey.get(corporationKey) ?? [];
}

function corporationIdentityKey(record) {
  const officeKey = String(record?.office_no ?? "").trim();
  if (!officeKey) return null;
  const index = buildCorporationGroupIndex();
  return index.keyByOffice.get(officeKey) ?? null;
}

function corporationLabelForKey(corporationKey) {
  const index = buildCorporationGroupIndex();
  return index.labelByKey.get(corporationKey) ?? "-";
}

function representativeRecords(representativeName) {
  const key = getRepresentativeKey(representativeName);
  if (!key) return [];
  return state.records.filter((record) => getRepresentativeKey(record.representative_name) === key);
}

function representativeCorporationGroups(representativeName) {
  const groups = new Map();
  representativeRecords(representativeName).forEach((record) => {
    const key = corporationIdentityKey(record);
    if (!key) return;
    const current =
      groups.get(key) ??
      {
        key,
        corporation_name: record.corporation_name,
        corporation_number: record.corporation_number,
        representative_name: record.representative_name,
        representative_role: record.representative_role,
        representative_raw: record.representative_raw,
        records: [],
      };
    current.records.push(record);
    groups.set(key, current);
  });
  return [...groups.values()]
    .map((group) => {
      const preferred = preferredCorporationRecord(group.records);
      return {
        ...group,
        corporation_name: compactInlineText(preferred?.corporation_name ?? group.corporation_name),
        corporation_number: compactInlineText(preferred?.corporation_number ?? group.corporation_number),
        representative_name: preferred?.representative_name ?? group.representative_name,
        representative_role: preferred?.representative_role ?? group.representative_role,
        representative_raw: preferred?.representative_raw ?? group.representative_raw,
      };
    })
    .sort((left, right) => {
    if (right.records.length !== left.records.length) return right.records.length - left.records.length;
    const wageDiff = meanFor(right.records, "average_wage_yen") - meanFor(left.records, "average_wage_yen");
    if (Number.isFinite(wageDiff) && wageDiff !== 0) return wageDiff;
    return String(left.corporation_name ?? "").localeCompare(String(right.corporation_name ?? ""), "ja");
  });
}

function representativeDisplayText(record) {
  return String(record?.representative_raw || record?.representative_name || "").trim();
}

function representativeLinkMarkup(record, className = "entity-link-button") {
  const label = representativeDisplayText(record);
  if (!label) return `<span class="flag-muted">未取得</span>`;
  const groupCount = representativeCorporationGroups(label).length;
  if (groupCount <= 1) {
    return escapeHtml(label);
  }
  return `<button class="${escapeAttribute(className)}" type="button" data-open-representative="${escapeAttribute(label)}">${escapeHtml(label)}</button>`;
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

function corporationLinkButton(recordOrName, className = "entity-link-button", officeNo = null) {
  const label = compactInlineText(
    typeof recordOrName === "object" && recordOrName != null ? recordOrName.corporation_name : recordOrName
  );
  const resolvedOfficeNo =
    officeNo ??
    (typeof recordOrName === "object" && recordOrName != null ? String(recordOrName.office_no ?? "").trim() : "");
  if (!label) return escapeHtml("-");
  if (!resolvedOfficeNo) return escapeHtml(label);
  return `<button class="${escapeAttribute(className)}" type="button" data-open-corporation-office="${escapeAttribute(resolvedOfficeNo)}">${escapeHtml(label)}</button>`;
}

function selectRecord(officeNo, options = {}) {
  const record = findRecordByOffice(officeNo);
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
  renderMarket(filtered);
  renderStrategy(filtered);
  renderStats(filtered);
  renderCharts(filtered);
  renderAnomalies(filtered);
  renderDetail(getSelectedRecord());
  renderTable(filtered);
  renderCompareSummary(filtered);
}

function renderCompareSummary(records) {
  const root = document.getElementById("compareSummary");
  if (!root) return;

  if (state.currentCompareView === "priority") {
    const summary = document.getElementById("strategySummary")?.textContent?.trim();
    root.textContent = summary || "重点候補を表示";
    return;
  }

  const pageCount = Math.max(1, Math.ceil(records.length / PAGE_SIZE));
  root.textContent = `${formatCount(records.length)} 件 / ${formatCount(pageCount)} ページ`;
}

function matchesFilters(record, filters) {
  if (filters.search) {
    const haystack = [
      record.office_no,
      record.municipality,
      getAreaLabel(record),
      record.corporation_name,
      record.representative_name,
      record.representative_raw,
      record.office_name,
      record.corporation_type_label,
      record.remarks,
      record.wam_primary_activity_type,
      record.wam_primary_activity_detail,
      deriveWorkModelLabel(record),
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
  if (filters.workShortageRisk === "likely" && !hasWorkShortageRisk(record)) {
    return false;
  }
  if (filters.primaryActivity !== "all" && (record.wam_primary_activity_type ?? "unknown") !== filters.primaryActivity) {
    return false;
  }
  if (filters.workModel !== "all" && deriveWorkModelLabel(record) !== filters.workModel) {
    return false;
  }
  if (filters.newOnly && !record.is_new_office) return false;
  if (filters.homeUseOnly && !record.home_use_active) return false;
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
  populateSelect("workShortageRiskSelect", "all", "すべての仕事状況", ["likely"]);
  populateSelect("primaryActivitySelect", "all", "すべての主活動", uniqueValues(records, "wam_primary_activity_type"));
  populateSelect("workModelSelect", "all", "すべての作業モデル", uniqueWorkModelValues(records));
}

function populateSelect(id, defaultValue, defaultLabel, values) {
  const root = document.getElementById(id);
  if (!root) return;
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

function uniqueWorkModelValues(records) {
  return [...new Set(records.map((record) => deriveWorkModelLabel(record)).filter(Boolean))].sort((left, right) =>
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
  document.getElementById("compareSummary").textContent = "データを読み込み中...";
  const filterTags = document.getElementById("activeFilterTags");
  if (filterTags) {
    filterTags.innerHTML = "";
  }
  document.getElementById("insightList").innerHTML = loadingCard;
  const marketList = document.getElementById("marketList");
  if (marketList) {
    marketList.innerHTML = loadingCard;
  }
  document.getElementById("growthList").innerHTML = loadingCard;
  document.getElementById("fixList").innerHTML = loadingCard;
  document.getElementById("highHighList").innerHTML = loadingCard;
  document.getElementById("statsGrid").innerHTML = loadingCard;
  document.getElementById("recordsCardList").innerHTML = loadingCard;
  [
    "municipalityChart",
    "wageChart",
    "corporationChart",
    "areaChart",
    "activityMedianChart",
    "homeUseComparisonChart",
    "corporationScaleChart",
    "capacityChart",
    "quadrantChart",
    "outlierChart",
    "wageBandChart",
    "utilizationScatter",
    "anomalyList",
    "detailContent",
    "detailDialogContent",
  ].forEach((id) => {
    const element = document.getElementById(id);
    if (element) {
      element.innerHTML = loadingCard;
    }
  });
  document.getElementById("recordsTableBody").innerHTML = `<tr><td colspan="11">${loadingCard}</td></tr>`;
}

function renderMeta(dashboard) {
  const updatedAt = dashboard.meta?.generated_at ? new Date(dashboard.meta.generated_at) : null;
  const homepageCount = state.records.filter((record) => Boolean(safeExternalUrl(record.homepage_url))).length;
  const instagramCount = state.records.filter((record) => Boolean(safeExternalUrl(record.instagram_url))).length;
  document.getElementById("updatedAt").textContent =
    updatedAt && !Number.isNaN(updatedAt.valueOf()) ? updatedAt.toLocaleString("ja-JP") : "-";
  document.getElementById("totalRecords").textContent = formatCount(state.records.length);
  document.getElementById("homepageCount").textContent = formatCount(homepageCount);
  document.getElementById("instagramCount").textContent = formatCount(instagramCount);
}

function renderQuality(dashboard, records = state.records) {
  const issuesRoot = document.getElementById("issuesList");
  const notesRoot = document.getElementById("notesList");
  const issues = dashboard.issues ?? [];
  const notes = dashboard.notes ?? [];
  const wageStats = computeLocalStats(numericValues(records, "average_wage_yen"));

  const summaryCards = [
    `<article class="note-card"><strong>表示対象</strong><p>大阪市内の就労継続支援B型 ${formatCount(
      records.length
    )} 件を表示している。</p></article>`,
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
  if (state.filters.workShortageRisk !== "all") filters.push(`仕事状況: ${labelForSelect(state.filters.workShortageRisk)}`);
  if (state.filters.primaryActivity !== "all") filters.push(`主活動: ${state.filters.primaryActivity}`);
  if (state.filters.workModel !== "all") filters.push(`作業モデル: ${state.filters.workModel}`);
  if (state.filters.newOnly) filters.push("新設のみ");
  if (state.filters.homeUseOnly) filters.push("在宅利用あり");

  document.getElementById("activeFilterSummary").textContent = filters.length
    ? `${formatCount(records.length)} 件を表示中。`
    : `${formatCount(records.length)} 件を表示中。追加条件なし。`;
  const tagsRoot = document.getElementById("activeFilterTags");
  if (tagsRoot) {
    const tags = [
      `期間: ${KPI_PERIOD_LABEL}`,
      `比較: ${KPI_BASELINE_LABEL}`,
      ...(state.filters.municipality === "all" ? ["対象: 大阪市全域"] : []),
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

  const wages = numericValues(records, "average_wage_yen");
  const wageStats = computeLocalStats(wages);

  // 好事例
  const topPerformers = topPerformerActivities(records, 3);
  const activityStats = buildGroupStats(records, (record) => deriveWorkModelLabel(record), "average_wage_yen", 5).sort(
    (left, right) => (right.median ?? 0) - (left.median ?? 0)
  );
  const homeUseRecords = records.filter((record) => record.home_use_active === true);
  const nonHomeUseRecords = records.filter((record) => record.home_use_active === false);

  // 利用率1%改善のインパクト
  const lowUtilRecords = records.filter((r) => isNumber(r.daily_user_capacity_ratio) && r.daily_user_capacity_ratio < 0.7);
  const avgRevPerPoint = lowUtilRecords.length
    ? Math.round(lowUtilRecords.reduce((s, r) => s + (revenuePerUtilizationPoint(r) ?? 0), 0) / lowUtilRecords.length)
    : null;

  const insights = [
    {
      title: "工賃の中心",
      body: wageStats.median != null
        ? `表示中の工賃の中心は ${formatMaybeYen(wageStats.median)} で、平均は ${formatMaybeYen(wageStats.mean)}。中心値から大きく離れた事業所を先に確認しやすい。`
        : "工賃データなし。",
    },
    {
      title: "利用率の改善余地",
      body: avgRevPerPoint != null
        ? `利用率70%未満の ${formatCount(lowUtilRecords.length)} 事業所の場合、利用率を1%改善すると平均で月 ${formatMaybeYen(avgRevPerPoint)} の増収。10%改善なら月 ${formatMaybeYen(avgRevPerPoint * 10)} の差になる。`
        : "利用率データが不足。",
    },
    {
      title: "高工賃の好事例",
      body: topPerformers.length
        ? topPerformers.map((p) => `${p.name}（${formatWageText(p.wage)}）: ${p.detail}`).join(" / ")
        : "工賃が高い事業所の作業内容が確認できない。",
    },
    {
      title: "主活動の勝ち筋",
      body: activityStats.length
        ? `${activityStats
            .slice(0, 2)
            .map((item) => `${item.label} は中央値 ${formatMaybeYen(item.median)}`)
            .join(" / ")}`
        : "主活動の公開情報が少なく、勝ち筋を読み切れない。",
    },
    {
      title: "在宅の使い方",
      body:
        homeUseRecords.length && nonHomeUseRecords.length
          ? `在宅ありの平均工賃は ${formatMaybeYen(meanFor(homeUseRecords, "average_wage_yen"))}、在宅なしは ${formatMaybeYen(meanFor(nonHomeUseRecords, "average_wage_yen"))}。在宅の有無で工賃差が出るかを確認しやすい。`
          : "在宅あり/なしの比較に十分な件数がない。",
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

function renderMarket(records) {
  const root = document.getElementById("marketList");
  if (!root) return;
  if (!records.length) {
    root.innerHTML = document.getElementById("emptyStateTemplate").innerHTML;
    return;
  }

  const cards = buildMarketCards(records);
  root.innerHTML = cards
    .map(
      (card) => `
        <article class="market-card">
          <h3>${escapeHtml(card.title)}</h3>
          <p class="market-summary">${escapeHtml(card.summary)}</p>
          <ul class="market-points">
            ${card.bullets.map((bullet) => `<li>${escapeHtml(bullet)}</li>`).join("")}
          </ul>
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
        hasWorkShortageRisk(record)
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
  const wages = numericValues(records, "average_wage_yen");
  const wageStats = computeLocalStats(wages);
  const utilizationMean = meanFor(records, "daily_user_capacity_ratio");
  const overallWageMean = meanFor(state.records, "average_wage_yen");
  const overallUtilizationStats = computeLocalStats(numericValues(state.records, "daily_user_capacity_ratio"));
  const utilizationMedian = overallUtilizationStats.median;
  const homeUseActiveCount = records.filter((record) => record.home_use_active === true).length;
  const homeUseActiveRate = ratioOf(records, (record) => record.home_use_active === true);
  const workShortageCount = records.filter((record) => hasWorkShortageRisk(record)).length;
  const recordShare = state.records.length ? records.length / state.records.length : null;
  const wageDiffFromOverall = isNumber(wageStats.mean) && isNumber(overallWageMean) ? wageStats.mean - overallWageMean : null;
  const utilizationDiffFromMedian =
    isNumber(utilizationMean) && isNumber(utilizationMedian) ? utilizationMean - utilizationMedian : null;
  const workShortageRatio = records.length ? workShortageCount / records.length : null;
  const homeUseRateDiffFromOverall =
    isNumber(homeUseActiveRate) && state.records.length
      ? homeUseActiveRate - ratioOf(state.records, (record) => record.home_use_active === true)
      : null;

  const cards = [
    {
      label: "対象事業所数",
      value: formatCount(records.length),
      hint: `${KPI_PERIOD_LABEL} / ${KPI_BASELINE_LABEL} ${formatCount(state.records.length)}件中 ${formatPercent(recordShare)}`,
    },
    {
      label: "平均工賃",
      value: formatMaybeYen(wageStats.mean),
      hint: `単位: 月額円 / ${KPI_BASELINE_LABEL}平均との差 ${formatSignedYen(wageDiffFromOverall)}`,
    },
    {
      label: "平均利用率",
      value: formatPercent(utilizationMean),
      hint: `単位: 定員比 / ${KPI_BASELINE_LABEL}中央値差 ${formatSignedPercentPoint(utilizationDiffFromMedian)}`,
    },
    {
      label: "重点確認件数",
      value: formatCount(workShortageCount),
      hint: `低工賃・低利用率など / 表示中の ${formatPercent(workShortageRatio)} / 在宅あり率 ${formatPercent(homeUseActiveRate)} (${formatSignedPercentPoint(homeUseRateDiffFromOverall)})`,
    },
  ];

  const contextRoot = document.getElementById("statsContextNote");
  if (contextRoot) {
    contextRoot.textContent = `期間: ${KPI_PERIOD_LABEL} / 比較対象: ${KPI_BASELINE_LABEL} ${formatCount(
      state.records.length
    )}件 / 単位: 工賃は月額円、利用率は定員比`;
  }

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
  renderBarChart("wageChart", topAverageWageByArea(records, 12), formatCountSuffix("円", true));
  renderBarChart("corporationChart", topCounts(records, "corporation_type_label", 6), formatCountSuffix("件"));
  renderBarChart("areaChart", topCounts(records.map((record) => ({ ...record, derived_area_label: getAreaLabel(record) })), "derived_area_label", 10), formatCountSuffix("件"));
  renderBarChart(
    "activityMedianChart",
    buildGroupStats(records, (record) => deriveWorkModelLabel(record), "average_wage_yen", 5)
      .sort((left, right) => (right.median ?? 0) - (left.median ?? 0))
      .slice(0, 8)
      .map((item) => ({ label: item.label, value: item.median })),
    formatCountSuffix("円", true)
  );
  renderBarChart(
    "homeUseComparisonChart",
    [
      { label: "在宅あり", value: meanFor(records.filter((record) => record.home_use_active === true), "average_wage_yen") },
      { label: "在宅なし", value: meanFor(records.filter((record) => record.home_use_active === false), "average_wage_yen") },
    ].filter((item) => isNumber(item.value)),
    formatCountSuffix("円", true)
  );
  renderBarChart(
    "corporationScaleChart",
    (() => {
      const groups = groupedCorporations(records);
      return [
        { label: "1拠点", value: groups.filter((group) => group.records.length === 1).length },
        { label: "2-3拠点", value: groups.filter((group) => group.records.length >= 2 && group.records.length <= 3).length },
        { label: "4拠点以上", value: groups.filter((group) => group.records.length >= 4).length },
      ].filter((item) => item.value > 0);
    })(),
    formatCountSuffix("件")
  );
  renderBarChart(
    "capacityChart",
    averageByOrderedGroup(records, "capacity_band_label", "average_wage_yen", CAPACITY_BAND_ORDER),
    formatCountSuffix("円", true)
  );
  renderBarChart("quadrantChart", orderedCounts(records, "market_position_quadrant", QUADRANT_ORDER), formatCountSuffix("件"));
  renderBarChart("outlierChart", combinedOutlierBreakdown(records), formatCountSuffix("件"));
  renderBarChart("wageBandChart", orderedCounts(records, "wage_band_label", WAGE_BAND_ORDER), formatCountSuffix("件"));
  renderScatterChart("utilizationScatter", records, {
    xKey: "daily_user_capacity_ratio",
    yKey: "average_wage_yen",
    xLabel: "利用率",
    yLabel: "平均工賃",
    yFormatter: (value) => formatMaybeYen(value),
    xFormatter: (value) => formatPercent(value),
  });
}

function renderBarChart(id, items, valueFormatter) {
  const root = document.getElementById(id);
  if (!root) return;
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
  if (!root) return;
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
        record.wage_outlier_flag === "high"
          ? "#c86f43"
          : isSelected
            ? "#233033"
            : "#0f7c79";
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
    .filter((record) => record.wage_outlier_flag || (record.daily_user_capacity_ratio ?? 0) > 1.5)
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
          data-flag="${escapeHtml(record.wage_outlier_flag ?? "none")}"
          type="button"
        >
          <div>
            <p class="section-kicker">${escapeHtml(record.municipality ?? "-")} / No.${escapeHtml(record.office_no ?? "-")}</p>
            <h3>${escapeHtml(record.office_name ?? "-")}</h3>
            <p>${escapeHtml(record.corporation_name ?? "-")}</p>
          </div>
          <div class="anomaly-meta">
            ${record.wage_outlier_flag ? `<span class="metric-chip">${escapeHtml(`工賃 ${labelForSelect(record.wage_outlier_flag)}`)}</span>` : ""}
            ${(record.daily_user_capacity_ratio ?? 0) > 1.5 ? `<span class="metric-chip" style="background:rgba(220,38,38,0.1);color:#dc2626">利用率過大 ${formatPercent(record.daily_user_capacity_ratio)}</span>` : ""}
            <span class="metric-chip">${escapeHtml(`工賃 ${formatWageText(record.average_wage_yen)}`)}</span>
          </div>
          <p>平均との差 ${escapeHtml(formatRatio(record.wage_ratio_to_overall_mean))} / 在宅率 ${escapeHtml(formatPercent(record.home_use_user_ratio_decimal))}</p>
          <p>${escapeHtml(record.wam_primary_activity_type ?? "活動不明")} / ${escapeHtml(record.remarks ?? "備考なし")}</p>
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
  const corporationName = compactInlineText(record.corporation_name);
  const corporationActionButton = corporationName
    ? `<button class="ghost-button" type="button" data-open-corporation-office="${escapeAttribute(record.office_no)}">法人の事業所一覧</button>`
    : "";
  const representativeLabel = representativeDisplayText(record);
  const representativeGroups = representativeCorporationGroups(representativeLabel);
  const representativeActionButton =
    representativeGroups.length > 1
      ? `<button class="ghost-button" type="button" data-open-representative="${escapeAttribute(representativeLabel)}">同じ代表者名の法人候補</button>`
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
            <p class="detail-subtitle">${corporationLinkButton(record, "entity-link-button detail-entity-link")} / ${escapeHtml(record.corporation_type_label ?? "-")}</p>
          </div>
        <div class="selected-office-status">
            ${record.wage_outlier_flag ? outlierBadge(record.wage_outlier_flag, record.wage_outlier_severity) : ""}
          </div>
        </div>
        <div class="selected-office-metrics">
          <span class="metric-chip">${escapeHtml(`平均工賃 ${formatWageText(record.average_wage_yen)}`)}</span>
          <span class="metric-chip">${escapeHtml(`利用率 ${formatPercent(record.daily_user_capacity_ratio)}`)}</span>
          <span class="metric-chip">${escapeHtml(`在宅率 ${formatPercent(record.home_use_user_ratio_decimal)}`)}</span>
        </div>
        <p class="selected-office-note">${escapeHtml(actionNotes[0] ?? "詳しい内訳は詳細を開いて確認できる。")}</p>
        <div class="selected-office-actions">
          ${corporationActionButton}
          ${representativeActionButton}
          ${homepageUrl ? `<a class="ghost-button link-button" href="${escapeAttribute(homepageUrl)}" target="_blank" rel="noreferrer">ホームページ</a>` : `<a class="ghost-button link-button" href="${escapeAttribute(websiteSearchUrl)}" target="_blank" rel="noreferrer">Webで探す</a>`}
          ${instagramUrl ? `<a class="ghost-button link-button" href="${escapeAttribute(instagramUrl)}" target="_blank" rel="noreferrer">Instagram</a>` : `<a class="ghost-button link-button" href="${escapeAttribute(instagramSearchUrl)}" target="_blank" rel="noreferrer">Instagramを探す</a>`}
        </div>
      </article>
    `;
  }

  if (!dialogRoot) return;

  const comparisonContext = getDetailComparisonContext(record);
  const peer = computePeerBenchmark(record, comparisonContext.records);
  const revPerPoint = revenuePerUtilizationPoint(record);
  const sortedWages = numericValues(comparisonContext.records, "average_wage_yen").sort((a, b) => a - b);
  const wagePercentile = computePercentileRank(record.average_wage_yen, sortedWages);
  const competitionProfile = computeCompetitionProfile(record, state.records);
  const capacityProfile = computeCapacityProfile(record, comparisonContext.records);
  const activityProfile = computeActivityProfile(record, comparisonContext.records);
  const corporationProfile = computeCorporationProfile(record);
  const homeUseProfile = computeHomeUseProfile(record, comparisonContext.records);
  const workShortageSignals = buildWorkShortageSignals(record);
  const kasanSim = computeKasanSimulation(record);
  const userFlow = computeUserFlowAnalysis(record, state.records);

  dialogRoot.innerHTML = `
    <div class="detail-hero">
      <div>
        <p class="section-kicker">${escapeHtml(record.municipality ?? "-")} / No.${escapeHtml(record.office_no ?? "-")}</p>
        <h3>${escapeHtml(record.office_name ?? "-")}</h3>
        <p class="detail-subtitle">${corporationLinkButton(record, "entity-link-button detail-entity-link")} / ${escapeHtml(record.corporation_type_label ?? "-")}</p>
      </div>
      <div class="detail-cta">
        ${corporationActionButton}
        ${representativeActionButton}
        ${homepageUrl ? `<a class="solid-button link-button" href="${escapeAttribute(homepageUrl)}" target="_blank" rel="noreferrer">ホームページ</a>` : `<a class="solid-button link-button" href="${escapeAttribute(websiteSearchUrl)}" target="_blank" rel="noreferrer">Webで探す</a>`}
        ${instagramUrl ? `<a class="ghost-button link-button" href="${escapeAttribute(instagramUrl)}" target="_blank" rel="noreferrer">Instagram</a>` : `<a class="ghost-button link-button" href="${escapeAttribute(instagramSearchUrl)}" target="_blank" rel="noreferrer">Instagramを探す</a>`}
      </div>
    </div>
    <div class="detail-kpi-grid">
      ${detailKpi("平均工賃", formatWageText(record.average_wage_yen), `${wagePercentile != null ? `上位 ${100 - wagePercentile}%` : "-"} / 平均との差 ${formatRatio(record.wage_ratio_to_overall_mean)}`)}
      ${detailKpi("利用率", formatPercent(record.daily_user_capacity_ratio), `定員 ${formatNullable(record.capacity)} 名 / 平均利用 ${formatNumber(record.average_daily_users)} 名`)}
      ${detailKpi("平均利用人数", formatDecimalUnit(record.average_daily_users, "人"), `定員 ${formatNullable(record.capacity)} 名に対して平均で使われている人数`)}
      ${detailKpi("定員", formatIntegerUnit(record.capacity, "名"), `参考定員 ${formatNullable(record.wam_office_capacity)} 名`)}
      ${detailKpi("在宅率", formatPercent(record.home_use_user_ratio_decimal), formatBool(record.home_use_active) === "あり" ? "在宅利用あり" : "在宅利用なし")}
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
          ${revPerPoint ? `<li>利用率1%改善: 月 ${formatMaybeYen(revPerPoint)} の増収</li>` : "<li>利用率改善の試算に必要な値が不足している</li>"}
          <li>主活動: ${escapeHtml(record.wam_primary_activity_type ?? "-")} ${record.wam_primary_activity_detail ? `（${escapeHtml(record.wam_primary_activity_detail)}）` : ""}</li>
        </ul>
      </article>
      <article class="detail-card detail-card-highlight">
        <h3>加算シミュレーション</h3>
        ${kasanSim.opportunities.length > 0 ? `
        <p class="kasan-impact-summary">未取得の加算を取得した場合の推定増収: <strong>月 ${formatMaybeYen(kasanSim.totalMonthlyOpportunity)}</strong>（年 ${formatMaybeYen(kasanSim.totalAnnualOpportunity)}）</p>
        ` : ""}
        <ul class="detail-list">
          ${kasanSim.items.map((item) => `<li>
            <span class="kasan-status kasan-status-${item.status === "取得済み" ? "acquired" : item.status === "確認推奨" ? "check" : "opportunity"}">${escapeHtml(item.status)}</span>
            <strong>${escapeHtml(item.name)}</strong>
            ${item.monthlyPerUser ? `— 1人あたり月 ${formatMaybeYen(item.monthlyPerUser)}${item.monthlyTotal ? ` / 全体で月 ${formatMaybeYen(item.monthlyTotal)}` : ""}` : ""}
            ${item.note ? `<br><small>${escapeHtml(item.note)}</small>` : ""}
            ${item.requirement ? `<br><small class="kasan-requirement">要件: ${escapeHtml(item.requirement)}</small>` : ""}
          </li>`).join("")}
        </ul>
        <p class="kasan-disclaimer"><small>※ 推定値です。実際の取得可否は行政書士・社労士にご確認ください。大阪市の地域単価 ${OSAKA_UNIT_YEN}円/単位、月${KASAN_WORKING_DAYS}日で試算。</small></p>
      </article>
      <article class="detail-card detail-card-highlight">
        <h3>利用者の流れ分析</h3>
        <ul class="detail-list">
          ${userFlow.opportunityCost ? `<li><strong>定員空きの機会損失:</strong> ${formatNumber(userFlow.opportunityCost.emptySlots)}人分 × 月${formatMaybeYen(userFlow.opportunityCost.dailyRevPerUser * KASAN_WORKING_DAYS)} = <strong>月 ${formatMaybeYen(userFlow.opportunityCost.monthlyLoss)}</strong>（年 ${formatMaybeYen(userFlow.opportunityCost.annualLoss)}）</li>` : "<li>定員充足済み — 空き枠なし</li>"}
          <li>同エリア（${escapeHtml(userFlow.areaLabel ?? "-")}）: ${formatCount(userFlow.areaTotal)}事業所中、利用率85%以上は ${formatCount(userFlow.areaHighUtilCount)}事業所</li>
          ${userFlow.advantages.length > 0 ? `<li><strong>高利用率事業所が持つ特徴で、この事業所にないもの:</strong><ul>${userFlow.advantages.map((a) => `<li>${escapeHtml(a.feature)} — ${escapeHtml(a.detail)}</li>`).join("")}</ul></li>` : "<li>高利用率事業所と比較して、主要な特徴差は検出されなかった。</li>"}
          ${userFlow.benchmarks.length > 0 ? `<li><strong>同エリアのベンチマーク先（利用率上位）:</strong><ul>${userFlow.benchmarks.map((b) => `<li>${escapeHtml(b.officeName)}（No.${b.officeNo}）— 工賃 ${formatMaybeYen(b.wage)} / 利用率 ${formatPercent(b.utilization)}${b.hasHomepage ? " / HP有" : ""}${b.hasTransport ? " / 送迎有" : ""}${b.hasHomeUse ? " / 在宅有" : ""}</li>`).join("")}</ul></li>` : ""}
          ${userFlow.patterns.total > 0 ? `<li>高利用率事業所の平均工賃: ${formatMaybeYen(userFlow.patterns.avgWage)} / 平均定員: ${formatNumber(userFlow.patterns.avgCapacity)}名</li>` : ""}
        </ul>
      </article>
      <article class="detail-card">
        <h3>エリア競合</h3>
        <ul class="detail-list">
          <li>同じエリア: ${escapeHtml(competitionProfile.areaLabel ?? "-")} に ${formatCount(competitionProfile.areaCount)}事業所</li>
          <li>同じエリアで同じ主活動: ${formatCount(competitionProfile.sameActivityAreaCount)}事業所</li>
          <li>エリア中央値工賃: ${escapeHtml(formatMaybeYen(competitionProfile.areaWageMedian))} / 平均利用率 ${escapeHtml(formatPercent(competitionProfile.areaUtilizationMean))}</li>
          <li>エリア密度順位: ${competitionProfile.areaRank ? `${formatCount(competitionProfile.areaRank)}位` : "-"} ${isNumber(competitionProfile.areaDensityMedian) ? ` / エリア件数の中央値 ${formatCount(Math.round(competitionProfile.areaDensityMedian))}件` : ""}</li>
        </ul>
      </article>
      <article class="detail-card">
        <h3>定員充足のしやすさ</h3>
        <ul class="detail-list">
          <li>定員まで平均あと: ${escapeHtml(formatDecimalUnit(capacityProfile.availableSlots, "人"))}</li>
          <li>同じ定員帯の中央値との差: ${escapeHtml(formatSignedDecimalUnit(isNumber(record.average_daily_users) && isNumber(capacityProfile.peerUsersMedian) ? record.average_daily_users - capacityProfile.peerUsersMedian : null, "人"))}</li>
          <li>同じ定員帯の利用率中央値との差: ${escapeHtml(formatSignedPercentPoint(isNumber(record.daily_user_capacity_ratio) && isNumber(capacityProfile.peerUtilizationMedian) ? record.daily_user_capacity_ratio - capacityProfile.peerUtilizationMedian : null))}</li>
          <li>比較件数: ${formatCount(capacityProfile.benchmarkCount)}件</li>
        </ul>
      </article>
      <article class="detail-card">
        <h3>主活動の勝ち筋</h3>
        <ul class="detail-list">
          <li>主活動: ${escapeHtml(activityProfile.activityLabel)}</li>
          <li>作業モデル: ${escapeHtml(activityProfile.workModelLabel)}</li>
          <li>同じ主活動の中央値工賃: ${escapeHtml(formatMaybeYen(activityProfile.wageMedian))} / 平均利用率 ${escapeHtml(formatPercent(activityProfile.utilizationMean))}</li>
          <li>${activityProfile.rank ? `同じ主活動 ${formatCount(activityProfile.peerCount)}件中 ${formatCount(activityProfile.rank)}位` : `同じ主活動 ${formatCount(activityProfile.peerCount)}件`}</li>
        </ul>
      </article>
      <article class="detail-card">
        <h3>法人内の横比較</h3>
        <ul class="detail-list">
          <li>同法人の運営事業所: ${formatCount(corporationProfile.officeCount)}件 / ${formatCount(corporationProfile.municipalityCount)}市区</li>
          <li>法人内平均工賃: ${escapeHtml(formatMaybeYen(corporationProfile.averageWage))}</li>
          <li>法人内平均利用率: ${escapeHtml(formatPercent(corporationProfile.utilizationMean))}</li>
          <li>法人内の工賃差: ${escapeHtml(formatMaybeYen(corporationProfile.wageSpread))}</li>
        </ul>
      </article>
      <article class="detail-card">
        <h3>在宅の使い方</h3>
        <ul class="detail-list">
          <li>この事業所の在宅利用: ${escapeHtml(formatBool(record.home_use_active))} ${isNumber(record.home_use_user_ratio_decimal) ? `（在宅率 ${escapeHtml(formatPercent(record.home_use_user_ratio_decimal))}）` : ""}</li>
          <li>在宅ありの平均工賃: ${escapeHtml(formatMaybeYen(homeUseProfile.homeUseAverageWage))} / なし ${escapeHtml(formatMaybeYen(homeUseProfile.nonHomeUseAverageWage))}</li>
          <li>在宅ありの平均利用率: ${escapeHtml(formatPercent(homeUseProfile.homeUseAverageUtilization))} / なし ${escapeHtml(formatPercent(homeUseProfile.nonHomeUseAverageUtilization))}</li>
          <li>この事業所が属する群の平均工賃: ${escapeHtml(formatMaybeYen(homeUseProfile.currentGroupAverageWage))} / 上位20%で在宅あり ${escapeHtml(formatPercent(homeUseProfile.topWageHomeUseRate))}</li>
        </ul>
      </article>
      <article class="detail-card">
        <h3>注視ポイント</h3>
        <ul class="detail-list">
          ${actionNotes.map((note) => `<li>${escapeHtml(note)}</li>`).join("")}
        </ul>
      </article>
      <article class="detail-card">
        <h3>仕事状況の見立て</h3>
        <ul class="detail-list">
          <li>仕事不足の可能性: ${escapeHtml(hasWorkShortageRisk(record) ? "あり" : "今のところ強くない")}</li>
          ${workShortageSignals.length ? workShortageSignals.map((signal) => `<li>${escapeHtml(signal)}</li>`).join("") : "<li>公開データ上では強い不足シグナルは少ない。</li>"}
        </ul>
      </article>
      ${
        representativeLabel
          ? `<article class="detail-card">
        <h3>代表者名の候補グループ</h3>
        <ul class="detail-list">
          <li>代表者: ${representativeLinkMarkup(record, "entity-link-button detail-entity-link")}</li>
          <li>候補法人数: ${formatCount(representativeGroups.length)} 法人</li>
          <li>候補事業所数: ${formatCount(representativeRecords(representativeLabel).length)} 事業所</li>
          <li>注意: 同姓同名の別人を含む可能性があるため、候補として扱う。</li>
        </ul>
      </article>`
          : ""
      }
      <article class="detail-card">
        <h3>基本情報</h3>
        <ul class="detail-list">
          <li>住所: ${escapeHtml(composeAddress(record))}</li>
          <li>電話: ${escapeHtml(record.wam_office_phone ?? "-")}</li>
          <li>事業所番号: ${escapeHtml(record.wam_office_number ?? "-")}</li>
          <li>代表者: ${representativeLinkMarkup(record, "entity-link-button detail-entity-link")}</li>
          <li>参考定員: ${escapeHtml(formatNullable(record.wam_office_capacity))}</li>
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

function renderCorporationDialog(corporationKey) {
  const dialogRoot = document.getElementById("corporationDialogContent");
  if (!dialogRoot) return;

  const records = sortCorporationRecords(corporationRecords(corporationKey));
  state.selectedCorporationKey = corporationKey;
  state.selectedCorporationLabel = corporationLabelForKey(corporationKey);

  if (!records.length) {
    dialogRoot.innerHTML = document.getElementById("emptyStateTemplate").innerHTML;
    return;
  }

  const filteredOfficeNos = new Set(state.filteredRecords.map((record) => String(record.office_no)));
  const visibleCount = records.filter((record) => filteredOfficeNos.has(String(record.office_no))).length;
  const municipalityCount = new Set(records.map((record) => record.municipality).filter(Boolean)).size;
  const osakaCityCount = records.filter((record) => record.municipality === "大阪市").length;
  const answeredWages = records.map((record) => record.average_wage_yen).filter(isNumber);
  const averageWage = answeredWages.length
    ? answeredWages.reduce((sum, value) => sum + value, 0) / answeredWages.length
    : null;
  const averageUtilization = meanFor(records, "daily_user_capacity_ratio");
  const representativeLabel = representativeDisplayText(records[0]);
  const representativeGroups = representativeCorporationGroups(representativeLabel);
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
        <h3>${escapeHtml(state.selectedCorporationLabel)}</h3>
        <p class="detail-subtitle">${escapeHtml(corporationTypeLabel)} / 大阪市内で ${formatCount(records.length)} 事業所${representativeLabel ? ` / 代表者 ${escapeHtml(representativeLabel)}` : ""}</p>
      </div>
      <div class="selected-office-actions">
        ${representativeGroups.length > 1 ? `<button class="ghost-button" type="button" data-open-representative="${escapeAttribute(representativeLabel)}">同じ代表者名の法人候補</button>` : ""}
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
        <em>大阪市全体</em>
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
        <span>平均利用率</span>
        <strong>${formatPercent(averageUtilization)}</strong>
        <em>法人内の平均稼働</em>
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
                  </div>
                </div>
                <div class="corporation-office-meta">
                  <span class="metric-chip">${escapeHtml(`平均工賃 ${formatWageText(record.average_wage_yen)}`)}</span>
                  <span class="metric-chip">${escapeHtml(`利用率 ${formatPercent(record.daily_user_capacity_ratio)}`)}</span>
                  <span class="metric-chip">${escapeHtml(`在宅率 ${formatPercent(record.home_use_user_ratio_decimal)}`)}</span>
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

function renderRepresentativeDialog(representativeName) {
  const dialogRoot = document.getElementById("representativeDialogContent");
  if (!dialogRoot) return;

  const records = representativeRecords(representativeName);
  const groups = representativeCorporationGroups(representativeName);
  state.selectedRepresentativeKey = getRepresentativeKey(representativeName);
  state.selectedRepresentativeLabel = representativeName;

  if (!records.length) {
    dialogRoot.innerHTML = document.getElementById("emptyStateTemplate").innerHTML;
    return;
  }

  const municipalityCount = new Set(records.map((record) => record.municipality).filter(Boolean)).size;
  const officeCount = records.length;
  const corporationCount = groups.length;

  dialogRoot.innerHTML = `
    <div class="corporation-hero">
      <div>
        <p class="section-kicker">同じ代表者名の法人候補</p>
        <h3>${escapeHtml(representativeName)}</h3>
        <p class="detail-subtitle">法人 ${formatCount(corporationCount)} 件 / 事業所 ${formatCount(officeCount)} 件 / ${formatCount(municipalityCount)} 市区</p>
      </div>
    </div>
    <div class="corporation-summary-grid">
      <article class="corporation-summary-card">
        <span>候補法人数</span>
        <strong>${formatCount(corporationCount)}</strong>
        <em>同じ代表者名でまとまる法人</em>
      </article>
      <article class="corporation-summary-card">
        <span>候補事業所数</span>
        <strong>${formatCount(officeCount)}</strong>
        <em>大阪市内の対象事業所</em>
      </article>
      <article class="corporation-summary-card">
        <span>平均工賃の平均</span>
        <strong>${escapeHtml(formatWageText(meanFor(records, "average_wage_yen")))}</strong>
        <em>候補全体の平均</em>
      </article>
      <article class="corporation-summary-card">
        <span>平均利用率</span>
        <strong>${formatPercent(meanFor(records, "daily_user_capacity_ratio"))}</strong>
        <em>候補全体の平均</em>
      </article>
    </div>
    <article class="detail-card">
      <h3>候補法人一覧</h3>
      <ul class="detail-list corporation-note-list">
        <li>同じ代表者名で自動集約した候補であり、同姓同名の別人を含む可能性がある。</li>
        <li>法人名を押すと法人別の事業所一覧へ移動できる。</li>
      </ul>
      <div class="corporation-office-list">
        ${groups
          .map((group) => {
            const groupAverageWage = meanFor(group.records, "average_wage_yen");
            const groupAverageUtilization = meanFor(group.records, "daily_user_capacity_ratio");
            const officeButtons = group.records
              .sort((left, right) => String(left.office_name ?? "").localeCompare(String(right.office_name ?? ""), "ja"))
              .map(
                (record) => `
                  <button class="table-link" type="button" data-select-office="${escapeAttribute(record.office_no)}">
                    ${escapeHtml(record.office_name ?? "-")}
                  </button>
                `
              )
              .join(" / ");

            return `
              <article class="corporation-office-item">
                <div class="corporation-office-head">
                  <div>
                    <p class="section-kicker">${escapeHtml(group.records[0]?.municipality ?? "-")} / 法人番号 ${escapeHtml(group.corporation_number ?? "-")}</p>
                    <h3>${corporationLinkButton(group.records[0], "entity-link-button detail-entity-link")}</h3>
                    <p class="detail-subtitle">${escapeHtml(group.representative_raw ?? representativeName)}</p>
                  </div>
                </div>
                <div class="corporation-office-meta">
                  <span class="metric-chip">${escapeHtml(`事業所 ${formatCount(group.records.length)} 件`)}</span>
                  <span class="metric-chip">${escapeHtml(`平均工賃 ${formatWageText(groupAverageWage)}`)}</span>
                  <span class="metric-chip">${escapeHtml(`平均利用率 ${formatPercent(groupAverageUtilization)}`)}</span>
                </div>
                <p class="corporation-office-note">${officeButtons}</p>
              </article>
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
    tableBody.innerHTML = `<tr><td colspan="11">${document.getElementById("emptyStateTemplate").innerHTML}</td></tr>`;
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
          <td>${corporationLinkButton(record, "entity-link-button table-entity-link")}</td>
          <td class="numeric">${formatWage(record.average_wage_yen, record.average_wage_error)}</td>
          <td class="numeric">${formatRatio(record.wage_ratio_to_overall_mean)}</td>
          <td class="numeric">${formatPercent(record.daily_user_capacity_ratio)}</td>
          <td class="numeric">${formatPercent(record.home_use_user_ratio_decimal)}</td>
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
                <p class="detail-subtitle">${corporationLinkButton(record, "entity-link-button detail-entity-link")}</p>
              </div>
              <button class="table-link" data-select-office="${escapeHtml(record.office_no)}" type="button">詳細</button>
            </div>
            <div class="record-card-metrics">
              <span class="metric-chip">${escapeHtml(`工賃 ${formatWageText(record.average_wage_yen)}`)}</span>
              <span class="metric-chip">${escapeHtml(`利用率 ${formatPercent(record.daily_user_capacity_ratio)}`)}</span>
              <span class="metric-chip">${escapeHtml(`在宅率 ${formatPercent(record.home_use_user_ratio_decimal)}`)}</span>
            </div>
            <div class="record-card-status">
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
  if (record.wage_outlier_flag === "high") classes.push("row-high-outlier");
  if (record.wage_outlier_flag === "low") classes.push("row-low-outlier");
  return classes.join(" ");
}

function attentionBadges(record) {
  const badges = [];
  if (record.wage_outlier_flag) {
    badges.push(outlierBadge(record.wage_outlier_flag, record.wage_outlier_severity));
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

function topAverageWageByArea(records, limit) {
  const buckets = new Map();
  records.forEach((record) => {
    const area = getAreaLabel(record);
    if (!area || !isNumber(record.average_wage_yen)) return;
    const bucket = buckets.get(area) ?? [];
    bucket.push(record.average_wage_yen);
    buckets.set(area, bucket);
  });
  return [...buckets.entries()]
    .filter(([, values]) => values.length >= 5)
    .map(([label, values]) => ({ label, value: values.reduce((sum, value) => sum + value, 0) / values.length }))
    .sort((left, right) => right.value - left.value)
    .slice(0, limit);
}

function combinedOutlierBreakdown(records) {
  return [
    { label: "工賃高め", value: records.filter((record) => record.wage_outlier_flag === "high").length },
    { label: "工賃低め", value: records.filter((record) => record.wage_outlier_flag === "low").length },
  ].filter((item) => item.value > 0);
}

function anomalyScore(record) {
  const wageScore = Math.abs(record.wage_z_score ?? 0) * 5 + (record.wage_ratio_to_overall_mean ?? 0) * 2;
  const utilizationOver150 = (record.daily_user_capacity_ratio ?? 0) > 1.5 ? 35 : 0;
  return wageScore + utilizationOver150;
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
    Math.max(0.7 - (record.daily_user_capacity_ratio ?? 0.7), 0) * 35 +
    Math.max((availableCapacity(record) ?? 0) - 3, 0) * 2 +
    (isLightWorkModel(record) ? 10 : 0) +
    ((record.home_use_user_ratio_decimal ?? 0) >= 0.3 ? 8 : 0)
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
  )} / 在宅率 ${formatPercent(record.home_use_user_ratio_decimal)}`;
}

function buildFixReason(record) {
  const workShortageNote = hasWorkShortageRisk(record) ? " / 仕事不足の可能性" : "";
  return `市町村平均との差 ${formatRatio(record.wage_ratio_to_municipality_mean)} / 利用率 ${formatPercent(
    record.daily_user_capacity_ratio
  )} / 在宅率 ${formatPercent(record.home_use_user_ratio_decimal)}${workShortageNote}`;
}

function buildHighHighReason(record) {
  return `市町村平均との差 ${formatRatio(record.wage_ratio_to_municipality_mean)} / 利用率 ${formatPercent(
    record.daily_user_capacity_ratio
  )} / 在宅率 ${formatPercent(record.home_use_user_ratio_decimal)}`;
}

function buildActionNotes(record) {
  const notes = [];

  // 1. 利用率
  if (isNumber(record.daily_user_capacity_ratio) && record.daily_user_capacity_ratio < 0.6) {
    const revPerPoint = revenuePerUtilizationPoint(record);
    notes.push(`【利用率】${formatPercent(record.daily_user_capacity_ratio)} と低水準。${revPerPoint ? `利用率10%改善で 月 ${formatMaybeYen(revPerPoint * 10)} の増収が見込める。体験利用の導線強化と相談支援事業所への営業強化を検討。` : "体験利用・紹介元の強化を検討したい。"}`);
  }

  // 2. 仕事不足リスク
  if (hasWorkShortageRisk(record)) {
    notes.push("【仕事不足】工賃と利用率の両方が弱め。企業営業・新規作業種目の開拓・施設外就労の検討を。");
  }

  // 3. 高工賃の好事例
  if (record.wage_outlier_flag === "high") {
    notes.push("【好事例】高工賃事業所。主活動と取引構造を他事業所への横展開候補として注目。作業内容と品質管理体制の共有を推奨。");
  }

  // 4. 新設
  if (record.is_new_office) {
    notes.push("【新設】立ち上がり期のため単月値ではなく定員充足の推移を追う。3か月目以降の利用率と工賃のトレンドが重要。");
  }

  if (!notes.length) {
    notes.push("大きな警戒シグナルなし。工賃と利用率の推移を定点観測したい。");
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
      label: usesAllRecords ? `大阪市全体 ${formatCount(filteredCount)}件` : `現在の絞り込み結果 ${formatCount(filteredCount)}件`,
      note: usesAllRecords
        ? "平均との差・中央値差と比較率は大阪市全体との比較で表示している。"
        : "平均との差・中央値差と比較率は現在の絞り込み結果を母集団にしている。絞り込みを変えると比較値も変わる。",
    };
  }

  return {
    records: state.records,
    label: `大阪市全体 ${formatCount(state.records.length)}件`,
    note: "この事業所は現在の一覧外のため、大阪市全体を母集団にして比較している。",
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
  ];
}

function compareToBaseline(value, baseline) {
  if (!isNumber(value) || !isNumber(baseline) || baseline === 0) return null;
  return value / baseline;
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
      const meanRatio = compareToBaseline(currentValue, stats.mean);
      const medianRatio = compareToBaseline(currentValue, stats.median);

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
            <span class="detail-comparison-heading">平均比</span>
            <strong class="detail-delta ${deltaClassName(isNumber(meanRatio) ? meanRatio - 1 : null)}">${escapeHtml(formatRelativeRatio(meanRatio))}</strong>
            <small>${escapeHtml(describeRelativeRatio(meanRatio, "平均"))}</small>
          </div>
          <div class="detail-comparison-cell">
            <span class="detail-comparison-heading">中央値差</span>
            <strong class="detail-delta ${deltaClassName(medianDiff)}">${escapeHtml(metric.formatDiff(medianDiff))}</strong>
            <small>中央値 ${escapeHtml(metric.formatValue(stats.median))}</small>
          </div>
          <div class="detail-comparison-cell">
            <span class="detail-comparison-heading">中央値比</span>
            <strong class="detail-delta ${deltaClassName(isNumber(medianRatio) ? medianRatio - 1 : null)}">${escapeHtml(formatRelativeRatio(medianRatio))}</strong>
            <small>${escapeHtml(describeRelativeRatio(medianRatio, "中央値"))}</small>
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
  if (record.osaka_city_ward) {
    return record.osaka_city_ward;
  }
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
  const highHomeUse = isNumber(record.home_use_user_ratio_decimal) && record.home_use_user_ratio_decimal >= 0.3;
  const lightWork = isLightWorkModel(record);
  const openSlots = isNumber(availableCapacity(record)) && availableCapacity(record) >= 5;

  return (
    (lowWage && lowUtilization) ||
    (lowWage && highHomeUse && lightWork) ||
    (lowWage && openSlots && (lowUtilization || lightWork)) ||
    (veryLowUtilization && openSlots)
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
    ["corporation_number", "法人番号"],
    ["corporation_name", "法人名"],
    ["representative_name", "代表者名"],
    ["representative_role", "代表者役職"],
    ["office_name", "事業所名"],
    ["derived_area_label", "エリア"],
    ["average_wage_yen", "平均工賃"],
    ["wage_ratio_to_overall_mean", "平均との差"],
    ["wage_ratio_to_municipality_mean", "市町村平均との差"],
    ["wage_ratio_to_capacity_band_mean", "定員帯平均との差"],
    ["daily_user_capacity_ratio", "利用率"],
    ["derived_available_capacity", "定員まで平均あと人数"],
    ["home_use_active", "在宅利用あり"],
    ["home_use_user_ratio_pct", "在宅率"],
    ["wage_outlier_flag", "工賃水準"],
    ["derived_work_shortage_risk", "仕事不足の可能性"],
    ["derived_work_model", "作業モデル"],
    ["derived_same_area_office_count", "同じエリアの事業所数"],
    ["derived_same_activity_area_count", "同じエリアで同じ主活動の事業所数"],
    ["derived_corporation_office_count", "同法人の事業所数"],
    ["wam_primary_activity_type", "主活動種別"],
    ["wam_office_number", "事業所番号"],
    ["wam_office_url", "掲載ページ"],
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
      if (key === "wage_outlier_flag") {
        return escapeCsvCell(labelForSelect(value ?? "none"));
      }
      if (key === "derived_work_shortage_risk") {
        return escapeCsvCell(hasWorkShortageRisk(record) ? "あり" : "なし");
      }
      if (key === "derived_area_label") {
        return escapeCsvCell(getAreaLabel(record));
      }
      if (key === "derived_available_capacity") {
        const value = availableCapacity(record);
        return escapeCsvCell(isNumber(value) ? decimalFormatter.format(value) : "");
      }
      if (key === "derived_work_model") {
        return escapeCsvCell(deriveWorkModelLabel(record));
      }
      if (key === "derived_same_area_office_count") {
        return escapeCsvCell(String(computeCompetitionProfile(record, state.records).areaCount ?? ""));
      }
      if (key === "derived_same_activity_area_count") {
        return escapeCsvCell(String(computeCompetitionProfile(record, state.records).sameActivityAreaCount ?? ""));
      }
      if (key === "derived_corporation_office_count") {
        return escapeCsvCell(String(computeCorporationProfile(record).officeCount ?? ""));
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

function formatRelativeRatio(value) {
  return isNumber(value) ? `${ratioFormatter.format(value)}倍` : "-";
}

function describeRelativeRatio(value, label) {
  if (!isNumber(value)) return `${label}との比較なし`;
  const gapPercent = (value - 1) * 100;
  if (Math.abs(gapPercent) < 0.05) return `${label}とほぼ同水準`;
  return `${label}を${percentFormatter.format(Math.abs(gapPercent))}% ${gapPercent > 0 ? "上回る" : "下回る"}`;
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
  if (source === "wam_url") return "掲載ページ";
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
