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

/* 比較用の月額ライン。報酬算定区分の境目として使う。 */
const REFERENCE_WAGE_LINE_YEN = 15000;

/* 北海道公式の就労継続支援B型 概要版から直近3年分を手元表示用に保持。
   令和5年度から平均工賃月額の算定方法が変更されている。 */
const HOKKAIDO_B_TYPE_HISTORY = {
  source_page_url: "https://www.pref.hokkaido.lg.jp/hf/shf/ko-chin.html",
  caution:
    "令和5年度から平均工賃月額の算定方法が変わっているため、令和4年度との伸びは制度上の計算変更も含む。",
  years: [
    {
      fiscal_year_label: "令和4年度",
      average_wage_monthly_yen: 19931.7,
      facility_count: 1021,
      average_wage_hourly_yen: 275.3,
      calculation_method_label: "旧算定方式",
      source_url:
        "https://www.pref.hokkaido.lg.jp/fs/1/2/9/0/1/1/4/7/_/%E4%BB%A4%E5%92%8C4%E5%B9%B4%E5%BA%A6%20%E6%A6%82%E8%A6%81%E7%89%88.pdf",
    },
    {
      fiscal_year_label: "令和5年度",
      average_wage_monthly_yen: 26675.4,
      facility_count: 1024,
      average_wage_hourly_yen: 261.7,
      calculation_method_label: "新算定方式へ変更",
      source_url:
        "https://www.pref.hokkaido.lg.jp/fs/1/2/9/0/1/1/3/6/_/R5%E5%B7%A5%E8%B3%83%E5%AE%9F%E7%B8%BE%E7%8A%B6%E6%B3%81%28%E6%A6%82%E8%A6%81%29.pdf",
    },
    {
      fiscal_year_label: "令和6年度",
      average_wage_monthly_yen: 27361,
      facility_count: 1235,
      average_wage_hourly_yen: 259.6,
      calculation_method_label: "現行算定方式",
      source_url:
        "https://www.pref.hokkaido.lg.jp/fs/1/2/9/0/1/1/2/3/_/R6%E5%B7%A5%E8%B3%83%E5%AE%9F%E7%B8%BE%E7%8A%B6%E6%B3%81%28%E6%A6%82%E8%A6%81%29.pdf",
    },
  ],
};

const HOS_CORPORATION_NAME = "合同会社 HOS";
const HOS_PRIMARY_OFFICE_NO = "518";
const HOS_OFFICIAL_SITE_URL = "https://hos-hokkaido.com/";
const HOS_OFFICIAL_COMPANY_URL = "https://hos-hokkaido.com/company.html";
const HOS_OFFICIAL_WORKPLACES = [
  {
    name: "ラインズ麻生",
    service: "就労継続支援B型",
    isDashboardTarget: true,
    officialCity: "札幌市",
    sourceUrl: HOS_OFFICIAL_COMPANY_URL,
    auditTone: "info",
    auditNote: "軽作業と清掃で所在地が分かれる。WAM収録は新琴似側、公式会社概要では麻生側も併記。",
  },
  {
    name: "ラインズ手稲",
    service: "就労継続支援B型",
    isDashboardTarget: true,
    officialCity: "札幌市",
    sourceUrl: HOS_OFFICIAL_COMPANY_URL,
  },
  {
    name: "ラインズ白石",
    service: "就労継続支援B型",
    isDashboardTarget: true,
    officialCity: "札幌市",
    sourceUrl: HOS_OFFICIAL_COMPANY_URL,
  },
  {
    name: "サニーズ東区役所前",
    service: "就労継続支援B型",
    isDashboardTarget: true,
    officialCity: "札幌市",
    sourceUrl: HOS_OFFICIAL_COMPANY_URL,
  },
  {
    name: "レガリス大通",
    service: "就労継続支援B型",
    isDashboardTarget: true,
    officialCity: "札幌市",
    sourceUrl: HOS_OFFICIAL_COMPANY_URL,
  },
  {
    name: "ローズ白石",
    service: "就労継続支援B型",
    isDashboardTarget: true,
    officialCity: "札幌市",
    sourceUrl: "https://hos-hokkaido.com/place/rose-shiroishi.html",
    auditTone: "alert",
    auditNote: "公式内で住所表記が揺れている。会社概要・WAMは本通4丁目北1-24、トップ/詳細は中央3条5丁目。",
  },
  {
    name: "ボルト澄川",
    service: "就労継続支援B型",
    isDashboardTarget: true,
    officialCity: "札幌市",
    sourceUrl: HOS_OFFICIAL_COMPANY_URL,
  },
  {
    name: "アルバ小樽",
    service: "就労継続支援B型",
    isDashboardTarget: true,
    officialCity: "小樽市",
    sourceUrl: "https://hos-hokkaido.com/place/otaru.html",
    auditTone: "alert",
    auditNote: "公式内で住所表記が揺れている。会社概要・WAMは稲穂4丁目5-16、トップ/詳細は入船2丁目5-1。",
  },
  {
    name: "ファニー函館",
    service: "就労継続支援B型",
    isDashboardTarget: true,
    officialCity: "函館市",
    sourceUrl: HOS_OFFICIAL_COMPANY_URL,
  },
  {
    name: "ファニー湯川",
    service: "就労移行支援・就労継続支援B型",
    isDashboardTarget: true,
    officialCity: "函館市",
    sourceUrl: "https://hos-hokkaido.com/place/yunokawa.html",
    auditTone: "info",
    auditNote: "就労移行支援との併設。ダッシュボードではB型実績として収録。",
  },
  {
    name: "DEKIRU～できる～",
    service: "就労選択支援",
    isDashboardTarget: false,
    officialCity: "函館市",
    sourceUrl: "https://hos-hokkaido.com/place/dekiru.html",
    auditTone: "neutral",
    auditNote: "B型ではないため、工賃実績ダッシュボードの集計対象外。",
  },
  {
    name: "だいじょうぶ",
    service: "指定特定相談支援・指定障害児相談支援",
    isDashboardTarget: false,
    officialCity: "函館市",
    sourceUrl: "https://hos-hokkaido.com/place/daijoubu.html",
    auditTone: "neutral",
    auditNote: "相談支援事業所のため、B型工賃実績ダッシュボードの集計対象外。",
  },
];

/* 人口順ソート用。北海道は令和7年1月1日、札幌市の区は令和8年1月1日の住民基本台帳人口を使用。 */
const MUNICIPALITY_POPULATION = {
  "札幌市": 1955678,
  "旭川市": 316183,
  "函館市": 236515,
  "苫小牧市": 165590,
  "帯広市": 160810,
  "釧路市": 154271,
  "江別市": 118055,
  "北見市": 110046,
  "小樽市": 104432,
  "千歳市": 97355,
  "室蘭市": 74855,
  "岩見沢市": 74204,
  "恵庭市": 70446,
  "石狩市": 57143,
  "北広島市": 56495,
  "登別市": 43615,
  "北斗市": 42810,
  "音更町": 42683,
  "滝川市": 36515,
  "網走市": 32199,
  "伊達市": 31208,
  "稚内市": 30336,
  "七飯町": 27139,
  "幕別町": 25269,
  "名寄市": 24742,
  "根室市": 22468,
  "中標津町": 22257,
  "紋別市": 19760,
  "富良野市": 19624,
  "美唄市": 18427,
  "釧路町": 18380,
  "深川市": 18329,
  "留萌市": 18169,
  "芽室町": 17773,
  "遠軽町": 17646,
  "美幌町": 17329,
  "倶知安町": 17120,
  "余市": 16954,
  "士別市": 16440,
  "砂川市": 15231,
  "当別町": 15113,
  "白老町": 15095,
  "八雲町": 14514,
  "別海町": 13964,
  "森町": 13582,
  "芦別市": 11243,
  "浦河町": 11231,
  "岩内町": 10913,
  "日高町": 10904,
  "栗山町": 10653,
  "斜里町": 10513,
  "長沼町": 9937,
  "東神楽町": 9775,
  "美瑛町": 9283,
  "清水町": 8755,
  "東川町": 8673,
  "赤平市": 8464,
  "厚岸町": 8195,
  "南幌町": 7933,
  "洞爺湖町": 7906,
  "三笠市": 7268,
  "むかわ町": 7221,
  "枝幸町": 7199,
  "標茶町": 6836,
  "せたな町": 6730,
  "江差町": 6607,
  "鷹栖町": 6463,
  "弟子屈町": 6432,
  "新十津川町": 6229,
  "夕張市": 6107,
  "当麻町": 6054,
  "本別町": 6046,
  "足寄町": 5952,
  "羽幌町": 5945,
  "池田町": 5906,
  "広尾町": 5880,
  "士幌町": 5690,
  "新得町": 5503,
  "新冠町": 5011,
  "鹿追町": 4928,
  "標津町": 4798,
  "上士幌町": 4758,
  "奈井江町": 4704,
  "訓子府町": 4448,
  "今金町": 4444,
  "平取町": 4435,
  "小清水町": 4342,
  "羅臼町": 4267,
  "厚真町": 4227,
  "津別町": 3970,
  "中札内村": 3823,
  "知内町": 3783,
  "美深町": 3681,
  "豊富町": 3515,
  "豊浦町": 3491,
  "比布町": 3425,
  "乙部町": 3107,
  "更別村": 3084,
  "仁木町": 3051,
  "下川町": 2836,
  "月形町": 2732,
  "新篠津村": 2730,
  "剣淵町": 2728,
  "小平町": 2664,
  "寿都町": 2629,
  "古平町": 2589,
  "妹背牛町": 2565,
  "黒松内町": 2446,
  "鶴居村": 2432,
  "愛別町": 2418,
  "上砂川町": 2360,
  "壮瞥町": 2313,
  "陸別町": 2120,
  "雨竜町": 2044,
  "幌延町": 2042,
  "喜茂別町": 1928,
  "中頓別町": 1472,
  "初山別村": 1005,
  "西興部村": 956,
};

const AREA_POPULATION = {
  "旭川市": 316183,
  "北区": 284333,
  "東区": 260418,
  "中央区": 245916,
  "函館市": 236515,
  "豊平区": 227092,
  "西区": 218835,
  "白石区": 213531,
  "苫小牧市": 165590,
  "帯広市": 160810,
  "釧路市": 154271,
  "手稲区": 139810,
  "南区": 133337,
  "厚別区": 123293,
  "江別市": 118055,
  "北見市": 110046,
  "清田区": 109113,
  "小樽市": 104432,
  "千歳市": 97355,
  "室蘭市": 74855,
  "岩見沢市": 74204,
  "恵庭市": 70446,
  "石狩市": 57143,
  "北広島市": 56495,
  "登別市": 43615,
  "北斗市": 42810,
  "音更町": 42683,
  "滝川市": 36515,
  "網走市": 32199,
  "伊達市": 31208,
  "稚内市": 30336,
  "七飯町": 27139,
  "幕別町": 25269,
  "名寄市": 24742,
  "根室市": 22468,
  "中標津町": 22257,
  "紋別市": 19760,
  "富良野市": 19624,
  "美唄市": 18427,
  "釧路町": 18380,
  "深川市": 18329,
  "留萌市": 18169,
  "芽室町": 17773,
  "遠軽町": 17646,
  "美幌町": 17329,
  "倶知安町": 17120,
  "余市": 16954,
  "士別市": 16440,
  "砂川市": 15231,
  "当別町": 15113,
  "白老町": 15095,
  "八雲町": 14514,
  "別海町": 13964,
  "森町": 13582,
  "芦別市": 11243,
  "浦河町": 11231,
  "岩内町": 10913,
  "日高町": 10904,
  "栗山町": 10653,
  "斜里町": 10513,
  "長沼町": 9937,
  "東神楽町": 9775,
  "美瑛町": 9283,
  "清水町": 8755,
  "東川町": 8673,
  "赤平市": 8464,
  "厚岸町": 8195,
  "南幌町": 7933,
  "洞爺湖町": 7906,
  "三笠市": 7268,
  "むかわ町": 7221,
  "枝幸町": 7199,
  "標茶町": 6836,
  "せたな町": 6730,
  "江差町": 6607,
  "鷹栖町": 6463,
  "弟子屈町": 6432,
  "新十津川町": 6229,
  "夕張市": 6107,
  "当麻町": 6054,
  "本別町": 6046,
  "足寄町": 5952,
  "羽幌町": 5945,
  "池田町": 5906,
  "広尾町": 5880,
  "士幌町": 5690,
  "新得町": 5503,
  "新冠町": 5011,
  "鹿追町": 4928,
  "標津町": 4798,
  "上士幌町": 4758,
  "奈井江町": 4704,
  "訓子府町": 4448,
  "今金町": 4444,
  "平取町": 4435,
  "小清水町": 4342,
  "羅臼町": 4267,
  "厚真町": 4227,
  "津別町": 3970,
  "中札内村": 3823,
  "知内町": 3783,
  "美深町": 3681,
  "豊富町": 3515,
  "豊浦町": 3491,
  "比布町": 3425,
  "乙部町": 3107,
  "更別村": 3084,
  "仁木町": 3051,
  "下川町": 2836,
  "月形町": 2732,
  "新篠津村": 2730,
  "剣淵町": 2728,
  "小平町": 2664,
  "寿都町": 2629,
  "古平町": 2589,
  "妹背牛町": 2565,
  "黒松内町": 2446,
  "鶴居村": 2432,
  "愛別町": 2418,
  "上砂川町": 2360,
  "壮瞥町": 2313,
  "陸別町": 2120,
  "雨竜町": 2044,
  "幌延町": 2042,
  "喜茂別町": 1928,
  "中頓別町": 1472,
  "初山別村": 1005,
  "西興部村": 956,
};

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
    level = "tier_6_1";
    label = "最も手厚い配置";
    qualifiedTier = "報酬上は最上位（6:1）";
    advice = "定員6人に対して職員1人の水準まで届いている。配置は厚く、報酬上も最も有利な立ち位置。";
  } else if (staffFte >= requiredStd) {
    level = "tier_7_5_1";
    label = "標準より手厚い配置";
    qualifiedTier = "報酬上は標準より上（7.5:1）";
    const gap = requiredHigh - staffFte;
    advice = `標準の7.5:1はクリア。あと常勤換算 ${ratioFormatter.format(Math.max(gap, 0))} 人で、最も手厚い6:1水準に届く。`;
  } else if (staffFte >= requiredMin) {
    level = "tier_10_1";
    label = "最低ラインで運営";
    qualifiedTier = "最低基準はクリア（10:1）";
    const gap = requiredStd - staffFte;
    advice = `法定の最低基準は満たしているが余裕は薄い。あと常勤換算 ${ratioFormatter.format(Math.max(gap, 0))} 人で、標準の7.5:1に届く。`;
  } else {
    level = "critical";
    label = "最低基準を下回る恐れ";
    qualifiedTier = "至急確認";
    advice = "法定の10:1を下回る可能性がある。配置表と請求区分を至急確認し、人員補充を検討したい。";
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

/** 月額1.5万円ライン到達率 */
function benchmarkWageRate(records) {
  const withWage = records.filter((r) => isNumber(r.average_wage_yen));
  if (!withWage.length) return null;
  return withWage.filter((r) => r.average_wage_yen >= REFERENCE_WAGE_LINE_YEN).length / withWage.length;
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
  "hos-office": {
    search: "HOS",
  },
  "sapporo-city": {
    municipality: "札幌市",
  },
  "sapporo-otaru-wam": {
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
const INITIAL_PRESET = "sapporo-city";

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
  selectedOfficeNo: HOS_PRIMARY_OFFICE_NO,
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
  renderHosManagement();
  renderHistoricalTrend(dashboard);
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

function normalizeSearchText(value) {
  return String(value ?? "")
    .normalize("NFKC")
    .toLowerCase()
    .replace(/[\s\u3000]+/g, " ")
    .trim();
}

function normalizeCorporationName(value) {
  return normalizeSearchText(value).replace(/\s+/g, "");
}

function normalizeOfficeName(value) {
  return normalizeSearchText(value).replace(/\s+/g, "");
}

function getHosRecords(records = state.records) {
  const targetName = normalizeCorporationName(HOS_CORPORATION_NAME);
  return records.filter((record) => normalizeCorporationName(record.corporation_name) === targetName);
}

function getHosOfficialDashboardTargets() {
  return HOS_OFFICIAL_WORKPLACES.filter((office) => office.isDashboardTarget);
}

function getHosOfficialNonDashboardTargets() {
  return HOS_OFFICIAL_WORKPLACES.filter((office) => !office.isDashboardTarget);
}

function getHosOfficialWorkplace(record) {
  const normalizedOfficeName = normalizeOfficeName(record?.office_name);
  if (!normalizedOfficeName) return null;
  return HOS_OFFICIAL_WORKPLACES.find((office) => normalizeOfficeName(office.name) === normalizedOfficeName) ?? null;
}

function getHosPrimaryOffice(records = state.records) {
  return (
    records.find((record) => String(record.office_no) === HOS_PRIMARY_OFFICE_NO) ??
    getHosRecords(records)[0] ??
    null
  );
}

function renderCorporationTrigger(name, className = "corporation-link") {
  const displayName = String(name ?? "").trim();
  const normalizedName = normalizeCorporationName(name);
  if (!normalizedName) return escapeHtml(displayName || "-");
  return `<button class="${escapeAttribute(className)}" data-open-corporation="${escapeAttribute(normalizedName)}" type="button">${escapeHtml(displayName)}</button>`;
}

function renderCorporationSubtitle(record, includeType = true) {
  const parts = [renderCorporationTrigger(record.corporation_name)];
  if (includeType && record.corporation_type_label) {
    parts.push(escapeHtml(record.corporation_type_label));
  }
  return `<p class="detail-subtitle">${parts.join(" / ")}</p>`;
}

function compareCorporationOffice(left, right) {
  const municipalityDiff = (MUNICIPALITY_POPULATION[right.municipality] ?? -1) - (MUNICIPALITY_POPULATION[left.municipality] ?? -1);
  if (municipalityDiff !== 0) return municipalityDiff;

  const leftArea = getAreaLabel(left);
  const rightArea = getAreaLabel(right);
  const areaDiff = (AREA_POPULATION[rightArea] ?? -1) - (AREA_POPULATION[leftArea] ?? -1);
  if (areaDiff !== 0) return areaDiff;

  return String(left.office_name ?? "").localeCompare(String(right.office_name ?? ""), "ja")
    || String(left.office_no ?? "").localeCompare(String(right.office_no ?? ""), "ja");
}

function openCorporationDialog(corporationName) {
  const dialog = document.getElementById("corporationDialog");
  const root = document.getElementById("corporationDialogContent");
  const heading = document.getElementById("corporationDialogHeading");
  const normalizedName = normalizeCorporationName(corporationName);
  if (!dialog || !root || !normalizedName) return;

  const offices = state.records
    .filter((record) => normalizeCorporationName(record.corporation_name) === normalizedName)
    .sort(compareCorporationOffice);

  const displayName = offices[0]?.corporation_name ?? normalizedName;
  const municipalityList = [...new Set(offices.map((record) => record.municipality).filter(Boolean))];
  const municipalitySummary = municipalityList.length
    ? `${municipalityList.slice(0, 4).join(" / ")}${municipalityList.length > 4 ? ` ほか ${formatCount(municipalityList.length - 4)} 市町村` : ""}`
    : "市町村情報なし";
  const corporationType = offices.find((record) => record.corporation_type_label)?.corporation_type_label ?? "法人種別未登録";

  if (heading) {
    heading.textContent = `${displayName} が運営する事業所`;
  }

  if (!offices.length) {
    root.innerHTML = document.getElementById("emptyStateTemplate").innerHTML;
    openDialogElement(dialog);
    return;
  }

  root.innerHTML = `
    <section class="corporation-summary">
      <article class="corporation-summary-card">
        <p class="section-kicker">法人サマリー</p>
        <h3>${escapeHtml(displayName)}</h3>
        <p>${escapeHtml(corporationType)} / 北海道内 ${formatCount(offices.length)} 事業所 / ${escapeHtml(municipalitySummary)}</p>
      </article>
    </section>
    <section class="corporation-office-list" aria-label="${escapeAttribute(displayName)} が運営する事業所一覧">
      ${offices.map((record) => {
        const staffing = getStaffingComplianceLevel(record);
        return `
          <article class="corporation-office-item">
            <div class="corporation-office-head">
              <div>
                <p class="section-kicker">${escapeHtml(record.municipality ?? "-")} / ${escapeHtml(getAreaLabel(record) ?? "-")} / No.${escapeHtml(record.office_no ?? "-")}</p>
                <h3>${escapeHtml(record.office_name ?? "-")}</h3>
                <p class="detail-subtitle">${escapeHtml(composeAddress(record))}</p>
              </div>
              <button class="table-link" data-open-office-detail="${escapeAttribute(record.office_no ?? "")}" type="button">詳細</button>
            </div>
            <div class="corporation-office-meta">
              <span class="metric-chip">${escapeHtml(`平均工賃 ${formatWageText(record.average_wage_yen)}`)}</span>
              <span class="metric-chip">${escapeHtml(`利用率 ${formatPercent(record.daily_user_capacity_ratio)}`)}</span>
              ${record.wam_match_status === "matched" ? `<span class="metric-chip">${escapeHtml(`定員に対する支援職員 ${formatPercent(record.wam_key_staff_fte_per_capacity)}`)}</span>` : ""}
              ${staffing ? `<span class="metric-chip">${escapeHtml(`職員配置 ${staffing.label}`)}</span>` : ""}
            </div>
          </article>
        `;
      }).join("")}
    </section>
  `;

  openDialogElement(dialog);
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
    syncAreaSelectOptions(filters);
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
    const statsTrigger = event.target.closest("[data-stat-action]");
    if (statsTrigger) {
      const action = statsTrigger.getAttribute("data-stat-action");
      if (!action) return;
      applyStatsAction(action);
      return;
    }
    const corporationTrigger = event.target.closest("[data-open-corporation]");
    if (corporationTrigger) {
      const corporationName = corporationTrigger.getAttribute("data-open-corporation");
      if (!corporationName) return;
      openCorporationDialog(corporationName);
      return;
    }
    const detailTrigger = event.target.closest("[data-open-office-detail]");
    if (detailTrigger) {
      const officeNo = detailTrigger.getAttribute("data-open-office-detail");
      if (!officeNo) return;
      closeDialogElement(document.getElementById("corporationDialog"));
      openRecordDetailByOfficeNo(officeNo);
      return;
    }
    const trigger = event.target.closest("[data-select-office]");
    if (!trigger) return;
    const officeNo = trigger.getAttribute("data-select-office");
    if (!officeNo) return;
    closeDialogElement(document.getElementById("corporationDialog"));
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
  syncAreaSelectOptions(filters);
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
  document.getElementById("noufukuOnlyCheckbox").checked = filters.noufukuOnly;
  document.getElementById("wamOnlyCheckbox").checked = filters.wamOnly;
}

function syncPresetButtons() {
  document.querySelectorAll(".preset-chip[data-preset]").forEach((button) => {
    button.classList.toggle("is-active", button.dataset.preset === state.activePreset);
  });
}

const MY_OFFICE_NO = HOS_PRIMARY_OFFICE_NO;

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

function selectRecord(officeNo, options = {}) {
  const record = state.filteredRecords.find((item) => String(item.office_no) === String(officeNo));
  if (!record) return;
  state.selectedOfficeNo = String(officeNo);
  renderCharts(state.filteredRecords);
  renderDetail(record);
  renderTable(state.filteredRecords);
  if (options.openDetail) {
    openDialogElement(document.getElementById("detailDialog"));
  }
}

function openRecordDetailByOfficeNo(officeNo) {
  const record = state.records.find((item) => String(item.office_no) === String(officeNo));
  if (!record) return;
  state.selectedOfficeNo = String(officeNo);
  renderDetail(record);
  openDialogElement(document.getElementById("detailDialog"));
}

function applyStatsAction(action) {
  if (action === "history") {
    document.getElementById("historyHeading")?.scrollIntoView({ behavior: "smooth", block: "start" });
    return;
  }
  if (action === "hos-corporation") {
    openCorporationDialog(HOS_CORPORATION_NAME);
    return;
  }
  if (action === "hos-detail") {
    openRecordDetailByOfficeNo(HOS_PRIMARY_OFFICE_NO);
    return;
  }
  if (action === "hos-audit") {
    document.getElementById("hosAuditList")?.scrollIntoView({ behavior: "smooth", block: "center" });
    return;
  }

  const nextFilters = cloneFilters(state.filters);
  if (action === "work-shortage") {
    nextFilters.workShortageRisk = "likely";
    nextFilters.quadrant = "all";
  } else if (action === "benchmark") {
    nextFilters.workShortageRisk = "all";
    nextFilters.quadrant = "高工賃 × 高稼働";
  } else {
    return;
  }

  state.filters = nextFilters;
  state.draftFilters = cloneFilters(nextFilters);
  state.activePreset = detectActivePreset(state.filters);
  syncFilterControls(state.draftFilters);
  syncPresetButtons();
  state.currentPage = 1;
  applyFilters();
  document.getElementById("tableHeading")?.scrollIntoView({ behavior: "smooth", block: "start" });
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
    const haystack = normalizeSearchText(
      [
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
    );
    if (!haystack.includes(normalizeSearchText(filters.search))) {
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
  populateSelect("municipalitySelect", "all", "すべての市町村", uniqueMunicipalityValues(records));
  syncAreaSelectOptions(state.draftFilters ?? state.filters, records);
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

function uniqueMunicipalityValues(records) {
  return sortValuesByPopulation(records.map((record) => record.municipality), MUNICIPALITY_POPULATION);
}

function uniqueAreaValues(records) {
  return sortValuesByPopulation(records.map((record) => getAreaLabel(record)), AREA_POPULATION);
}

function sortValuesByPopulation(values, populationLookup) {
  return [...new Set(values.filter(Boolean))].sort((left, right) => {
    const leftPopulation = populationLookup[left] ?? -1;
    const rightPopulation = populationLookup[right] ?? -1;
    if (leftPopulation !== rightPopulation) {
      return rightPopulation - leftPopulation;
    }
    return String(left).localeCompare(String(right), "ja");
  });
}

function areaSourceRecords(records, municipality = "all") {
  if (municipality === "all") return records;
  return records.filter((record) => record.municipality === municipality);
}

function syncAreaSelectOptions(filters = state.filters, records = state.records) {
  const areaValues = uniqueAreaValues(areaSourceRecords(records, filters?.municipality ?? "all"));
  const validAreaValues = new Set(["all", ...areaValues]);
  if (!validAreaValues.has(filters.area)) {
    filters.area = "all";
  }
  populateSelect("areaSelect", "all", "すべてのエリア", areaValues);
  document.getElementById("areaSelect").value = filters.area;
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
  document.getElementById("workShortageList").innerHTML = loadingCard;
  document.getElementById("fixList").innerHTML = loadingCard;
  document.getElementById("highHighList").innerHTML = loadingCard;
  document.getElementById("statsGrid").innerHTML = loadingCard;
  document.getElementById("hosStatsGrid").innerHTML = loadingCard;
  document.getElementById("hosPeerList").innerHTML = loadingCard;
  document.getElementById("hosBenchmarkList").innerHTML = loadingCard;
  document.getElementById("hosAuditList").innerHTML = loadingCard;
  document.getElementById("hosSummary").textContent = "HOS向け要点を読み込み中...";
  document.getElementById("historyTrendGrid").innerHTML = loadingCard;
  document.getElementById("historyTrendFoot").innerHTML = "";
  document.getElementById("historyTrendSummary").textContent = "年度推移を読み込み中...";
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

function rankOfficeByWage(targetOfficeNo, records) {
  const comparableRecords = records.filter((record) => isNumber(record.average_wage_yen));
  const sorted = comparableRecords
    .slice()
    .sort((left, right) => right.average_wage_yen - left.average_wage_yen);
  const rank = sorted.findIndex((record) => String(record.office_no) === String(targetOfficeNo));
  if (rank === -1) return null;
  return { rank: rank + 1, total: sorted.length };
}

function hosAttentionScore(record, hosMean) {
  const wageGapScore =
    isNumber(record.average_wage_yen) && isNumber(hosMean) && hosMean > 0
      ? Math.max((hosMean - record.average_wage_yen) / hosMean, 0) * 90
      : 0;
  const utilizationScore = isNumber(record.daily_user_capacity_ratio)
    ? Math.max(0.75 - record.daily_user_capacity_ratio, 0) * 120
    : 0;
  const shortageScore = hasWorkShortageRisk(record) ? 80 : 0;
  const lowQuadrantScore =
    record.market_position_quadrant === "低工賃 × 低稼働"
      ? 45
      : record.market_position_quadrant === "低工賃 × 高稼働"
        ? 25
        : 0;
  const staffingScore = record.wam_staffing_efficiency_quadrant === "高工賃 × 少ない人員" ? 16 : 0;
  const auditScore = getHosOfficialWorkplace(record)?.auditTone === "alert" ? 12 : 0;
  return shortageScore + wageGapScore + utilizationScore + lowQuadrantScore + staffingScore + auditScore;
}

function buildHosAttentionCandidates(hosRecords) {
  const hosMean = meanFor(hosRecords, "average_wage_yen");
  return hosRecords
    .slice()
    .sort((left, right) => hosAttentionScore(right, hosMean) - hosAttentionScore(left, hosMean))
    .slice(0, 4);
}

function buildHosOfficialAuditRows(hosRecords) {
  const officialTargets = getHosOfficialDashboardTargets();
  const officialNonTargets = getHosOfficialNonDashboardTargets();
  const recordedOfficeNames = new Set(hosRecords.map((record) => normalizeOfficeName(record.office_name)));
  const missingTargets = officialTargets.filter((office) => !recordedOfficeNames.has(normalizeOfficeName(office.name)));
  const matchedTargetCount = officialTargets.length - missingTargets.length;
  const rawCorporationNames = [...new Set(hosRecords.map((record) => String(record.corporation_name ?? "").trim()).filter(Boolean))];
  const rows = [
    {
      tone: missingTargets.length ? "alert" : "good",
      title: `公式B型 ${formatCount(matchedTargetCount)} / ${formatCount(officialTargets.length)} 収録`,
      body: missingTargets.length
        ? `未収録: ${missingTargets.map((office) => office.name).join("、")}`
        : "公式会社概要に掲載されているB型10拠点をすべて収録。",
      chips: ["公式照合", "B型"],
      sourceUrl: HOS_OFFICIAL_COMPANY_URL,
    },
  ];

  if (rawCorporationNames.length > 1) {
    rows.push({
      tone: "info",
      title: "法人名の表記ゆれ",
      body: "合同会社HOSの法人名に全角/半角・空白ありの表記が混在。集計では同一法人として統合済み。",
      chips: ["集計補正", `${formatCount(rawCorporationNames.length)}表記`],
    });
  }

  if (officialNonTargets.length) {
    rows.push({
      tone: "neutral",
      title: "B型外の公式掲載拠点",
      body: `${officialNonTargets.map((office) => `${office.name}（${office.service}）`).join("、")} は工賃実績の集計対象外。`,
      chips: ["対象外", `${formatCount(officialNonTargets.length)}件`],
      sourceUrl: HOS_OFFICIAL_SITE_URL,
    });
  }

  officialTargets
    .filter((office) => office.auditNote)
    .forEach((office) => {
      rows.push({
        tone: office.auditTone ?? "info",
        title: office.name,
        body: office.auditNote,
        chips: [office.service, office.officialCity].filter(Boolean),
        sourceUrl: office.sourceUrl,
      });
    });

  return rows;
}

function buildHosBenchmarkCandidates(records, hosOffice) {
  const hosRecords = getHosRecords(records);
  const focusMunicipalities = new Set(hosRecords.map((record) => record.municipality).filter(Boolean));
  if (!focusMunicipalities.size) {
    focusMunicipalities.add("札幌市");
    if (hosOffice.municipality) focusMunicipalities.add(hosOffice.municipality);
  }
  const hosOfficeNos = new Set(hosRecords.map((record) => String(record.office_no)));
  return records
    .filter(
      (record) =>
        !hosOfficeNos.has(String(record.office_no)) &&
        focusMunicipalities.has(record.municipality) &&
        record.market_position_quadrant === "高工賃 × 高稼働" &&
        isNumber(record.average_wage_yen)
    )
    .sort((left, right) => {
      const leftSameBand = left.capacity_band_label === hosOffice.capacity_band_label ? 1 : 0;
      const rightSameBand = right.capacity_band_label === hosOffice.capacity_band_label ? 1 : 0;
      if (leftSameBand !== rightSameBand) return rightSameBand - leftSameBand;

      const leftMatched = left.wam_match_status === "matched" ? 1 : 0;
      const rightMatched = right.wam_match_status === "matched" ? 1 : 0;
      if (leftMatched !== rightMatched) return rightMatched - leftMatched;

      return right.average_wage_yen - left.average_wage_yen;
    })
    .slice(0, 4);
}

function buildHosPeerReason(record) {
  const parts = [];
  const hosMean = meanFor(getHosRecords(state.records), "average_wage_yen");
  const wageGapToHos =
    isNumber(record.average_wage_yen) && isNumber(hosMean)
      ? record.average_wage_yen - hosMean
      : null;
  const wageGapToOverall =
    isNumber(record.average_wage_yen) && isNumber(state.dashboard?.analytics?.overall_wage_stats?.mean)
      ? record.average_wage_yen - state.dashboard.analytics.overall_wage_stats.mean
      : null;
  if (wageGapToHos != null) {
    parts.push(`HOS平均との差 ${formatSignedYen(wageGapToHos)}`);
  }
  if (wageGapToOverall != null) {
    parts.push(`全道平均との差 ${formatSignedYen(wageGapToOverall)}`);
  }
  if (hasWorkShortageRisk(record)) {
    parts.push("仕事不足候補");
  }
  if (isNumber(record.daily_user_capacity_ratio) && record.daily_user_capacity_ratio < 0.7) {
    parts.push("利用率70%未満");
  }
  if (record.wam_match_status !== "matched") {
    parts.push("人員詳細未取得");
  }
  const officialWorkplace = getHosOfficialWorkplace(record);
  if (officialWorkplace?.auditTone === "alert") {
    parts.push("公式表記差分あり");
  }
  parts.push(`利用率 ${formatPercent(record.daily_user_capacity_ratio)}`);
  return parts.join(" / ");
}

function buildHosBenchmarkReason(record) {
  const staffing = getStaffingComplianceLevel(record);
  const parts = [
    `平均工賃 ${formatMaybeYen(record.average_wage_yen)}`,
    `利用率 ${formatPercent(record.daily_user_capacity_ratio)}`,
  ];
  if (staffing) {
    parts.push(staffing.label);
  }
  return parts.join(" / ");
}

function renderHosWatchList(rootId, records, emptyLabel, buildReason) {
  const root = document.getElementById(rootId);
  if (!root) return;
  if (!records.length) {
    root.innerHTML = `<div class="empty-state"><h3>${escapeHtml(emptyLabel)}</h3></div>`;
    return;
  }

  root.innerHTML = records
    .map(
      (record) => `
        <article class="hos-watch-item">
          <div>
            <p class="strategy-kicker">${escapeHtml(record.municipality ?? "-")} / No.${escapeHtml(record.office_no ?? "-")}</p>
            <strong>${escapeHtml(record.office_name ?? "-")}</strong>
            <p class="detail-subtitle">${escapeHtml(record.corporation_name ?? "法人名未登録")}</p>
          </div>
          <div class="corporation-office-meta">
            <span class="metric-chip">${escapeHtml(`平均工賃 ${formatWageText(record.average_wage_yen)}`)}</span>
            <span class="metric-chip">${escapeHtml(`利用率 ${formatPercent(record.daily_user_capacity_ratio)}`)}</span>
            ${record.wam_match_status === "matched" ? `<span class="metric-chip">人員詳細あり</span>` : ""}
          </div>
          <p>${escapeHtml(buildReason(record))}</p>
          <div class="hos-watch-actions">
            <button class="table-link" data-open-office-detail="${escapeAttribute(record.office_no ?? "")}" type="button">詳細</button>
          </div>
        </article>
      `
    )
    .join("");
}

function renderHosAuditList(rootId, rows) {
  const root = document.getElementById(rootId);
  if (!root) return;
  if (!rows.length) {
    root.innerHTML = `<div class="empty-state"><h3>公式照合メモなし</h3></div>`;
    return;
  }

  root.innerHTML = rows
    .map((row) => {
      const sourceUrl = safeExternalUrl(row.sourceUrl);
      return `
        <article class="hos-audit-item hos-audit-${escapeAttribute(row.tone ?? "neutral")}">
          <div>
            <p class="strategy-kicker">${escapeHtml(row.tone === "alert" ? "要確認" : row.tone === "good" ? "確認済み" : "補足")}</p>
            <strong>${escapeHtml(row.title)}</strong>
          </div>
          <p>${escapeHtml(row.body)}</p>
          <div class="corporation-office-meta">
            ${(row.chips ?? []).map((chip) => `<span class="metric-chip">${escapeHtml(chip)}</span>`).join("")}
          </div>
          ${sourceUrl ? `<a class="table-link" href="${escapeAttribute(sourceUrl)}" target="_blank" rel="noreferrer">公式確認</a>` : ""}
        </article>
      `;
    })
    .join("");
}

function renderHosManagement() {
  const summary = document.getElementById("hosSummary");
  const statsRoot = document.getElementById("hosStatsGrid");
  const peerRoot = document.getElementById("hosPeerList");
  const benchmarkRoot = document.getElementById("hosBenchmarkList");
  const auditRoot = document.getElementById("hosAuditList");
  if (!summary || !statsRoot || !peerRoot || !benchmarkRoot || !auditRoot) return;

  const hosRecords = getHosRecords(state.records);
  const hosOffice = getHosPrimaryOffice(state.records);
  if (!hosOffice) {
    summary.textContent = "HOSの事業所データなし";
    statsRoot.innerHTML = document.getElementById("emptyStateTemplate").innerHTML;
    peerRoot.innerHTML = document.getElementById("emptyStateTemplate").innerHTML;
    benchmarkRoot.innerHTML = document.getElementById("emptyStateTemplate").innerHTML;
    auditRoot.innerHTML = document.getElementById("emptyStateTemplate").innerHTML;
    return;
  }

  const allRank = rankOfficeByWage(hosOffice.office_no, state.records);
  const overallMean = meanFor(state.records, "average_wage_yen");
  const hosMean = meanFor(hosRecords, "average_wage_yen");
  const hosMatchedCount = hosRecords.filter((record) => record.wam_match_status === "matched").length;
  const hosHighHighCount = hosRecords.filter((record) => record.market_position_quadrant === "高工賃 × 高稼働").length;
  const hosWorkShortageCount = hosRecords.filter((record) => hasWorkShortageRisk(record)).length;
  const hosLowUtilCount = hosRecords.filter(
    (record) => isNumber(record.daily_user_capacity_ratio) && record.daily_user_capacity_ratio < 0.7
  ).length;
  const officialTargets = getHosOfficialDashboardTargets();
  const officialNonTargets = getHosOfficialNonDashboardTargets();
  const recordedOfficeNames = new Set(hosRecords.map((record) => normalizeOfficeName(record.office_name)));
  const officialMatchedCount = officialTargets.filter((office) => recordedOfficeNames.has(normalizeOfficeName(office.name))).length;
  const auditRows = buildHosOfficialAuditRows(hosRecords);
  const auditIssueCount = auditRows.filter((row) => row.tone === "alert").length;
  const municipalitySummary = [...new Set(hosRecords.map((record) => record.municipality).filter(Boolean))]
    .sort((left, right) => (MUNICIPALITY_POPULATION[right] ?? -1) - (MUNICIPALITY_POPULATION[left] ?? -1))
    .map((municipality) => {
      const count = hosRecords.filter((record) => record.municipality === municipality).length;
      return `${municipality}${formatCount(count)}件`;
    })
    .join(" / ");
  const staffing = getStaffingComplianceLevel(hosOffice);
  const attentionCandidates = buildHosAttentionCandidates(hosRecords);
  const benchmarkCandidates = buildHosBenchmarkCandidates(state.records, hosOffice);

  summary.textContent = [
    `公式B型 ${formatCount(officialMatchedCount)} / ${formatCount(officialTargets.length)}収録`,
    municipalitySummary || null,
    `高工賃・高利用率 ${formatCount(hosHighHighCount)}件`,
    `仕事不足候補 ${formatCount(hosWorkShortageCount)}件`,
    `B型外 ${formatCount(officialNonTargets.length)}件は対象外`,
  ]
    .filter(Boolean)
    .join(" / ");

  const cards = [
    {
      label: "公式B型収録",
      value: `${formatCount(officialMatchedCount)} / ${formatCount(officialTargets.length)}`,
      hint: `公式掲載のうちB型外${formatCount(officialNonTargets.length)}件は別枠。HOS自社一覧はここから確認。`,
      action: "hos-corporation",
      tone: "history",
      cta: "一覧を見る",
    },
    {
      label: "HOS平均工賃",
      value: formatMaybeYen(hosMean),
      hint:
        overallMean != null && hosMean != null
          ? `全道平均との差 ${formatSignedYen(hosMean - overallMean)}`
          : "平均工賃を計算できない",
    },
    {
      label: "横展開したい主力",
      value: formatCount(hosHighHighCount),
      hint: `${formatPercent(hosRecords.length ? hosHighHighCount / hosRecords.length : null)} / 横展開したい主力拠点`,
      action: "benchmark",
      tone: "good",
      cta: "好事例一覧",
    },
    {
      label: "仕事が少ない候補",
      value: formatCount(hosWorkShortageCount),
      hint:
        hosWorkShortageCount > 0
          ? `HOS内で仕事量・受注先を先に見る候補。利用率70%未満は${formatCount(hosLowUtilCount)}件。`
          : `強い不足シグナルなし。利用率70%未満は${formatCount(hosLowUtilCount)}件。`,
      action: "work-shortage",
      tone: hosWorkShortageCount > 0 ? "alert" : "good",
      cta: "一覧を見る",
    },
    {
      label: "公式表記の要確認",
      value: formatCount(auditIssueCount),
      hint: auditIssueCount > 0 ? "住所表記など、役員会前に確認したい差分あり。" : "大きな表記差分なし",
      action: "hos-audit",
      tone: auditIssueCount > 0 ? "alert" : "good",
      cta: "照合を見る",
    },
    {
      label: "人員詳細一致",
      value: `${formatCount(hosMatchedCount)} / ${formatCount(hosRecords.length)}`,
      hint:
        hosMatchedCount < hosRecords.length
          ? "一部拠点でWAM人員詳細が未一致"
          : "HOS拠点はすべて人員詳細まで見える",
    },
    {
      label: "代表拠点の詳細",
      value: hosOffice.office_name ?? "-",
      hint: staffing
        ? `${hosOffice.municipality} / ${staffing.label} / ${allRank ? `北海道上位 ${formatPercent(allRank.rank / allRank.total)}` : "順位計算不可"}`
        : `${hosOffice.municipality} / まず詳細を確認したい代表拠点`,
      action: "hos-detail",
      tone: "history",
      cta: "詳細を見る",
    },
  ];

  statsRoot.innerHTML = cards
    .map(
      (card) => `
        <${card.action ? "button" : "article"} class="stat-card${card.action ? ` stat-card-action stat-card-${escapeAttribute(card.tone ?? "accent")}` : ""}"${card.action ? ` data-stat-action="${escapeAttribute(card.action)}" type="button"` : ""}>
          <p>${escapeHtml(card.label)}</p>
          <strong>${escapeHtml(card.value)}</strong>
          <em>${escapeHtml(card.hint)}</em>
          ${card.action ? `<span class="stat-card-cta">${escapeHtml(card.cta ?? "詳細を見る")}</span>` : ""}
        </${card.action ? "button" : "article"}>
      `
    )
    .join("");

  renderHosWatchList("hosPeerList", attentionCandidates, "先に確認したい HOS 拠点はまだ抽出できない", buildHosPeerReason);
  renderHosWatchList("hosBenchmarkList", benchmarkCandidates, "札幌・小樽の好事例はまだ抽出できない", buildHosBenchmarkReason);
  renderHosAuditList("hosAuditList", auditRows);
}

function renderHistoricalTrend(dashboard) {
  const root = document.getElementById("historyTrendGrid");
  const foot = document.getElementById("historyTrendFoot");
  const summary = document.getElementById("historyTrendSummary");
  if (!root || !foot || !summary) return;

  const history = dashboard.analytics?.historical_b_type_overview ?? HOKKAIDO_B_TYPE_HISTORY;
  const years = Array.isArray(history.years) ? history.years : [];
  if (!years.length) {
    root.innerHTML = document.getElementById("emptyStateTemplate").innerHTML;
    foot.innerHTML = "";
    summary.textContent = "過去3年の概要データなし";
    return;
  }

  const sortedYears = [...years].sort((left, right) =>
    String(left.fiscal_year_label ?? "").localeCompare(String(right.fiscal_year_label ?? ""), "ja")
  );
  const latest = sortedYears[sortedYears.length - 1];
  const baseline = sortedYears[0];
  const totalGrowthRate =
    isNumber(latest.average_wage_monthly_yen) &&
    isNumber(baseline.average_wage_monthly_yen) &&
    baseline.average_wage_monthly_yen > 0
      ? latest.average_wage_monthly_yen / baseline.average_wage_monthly_yen - 1
      : null;
  const facilityDelta =
    isNumber(latest.facility_count) && isNumber(baseline.facility_count)
      ? latest.facility_count - baseline.facility_count
      : null;

  summary.textContent = totalGrowthRate != null
    ? `${baseline.fiscal_year_label}比 ${formatSignedPercent(totalGrowthRate)} / 施設数 ${formatSignedCount(facilityDelta)}`
    : "直近3年の全道推移";

  root.innerHTML = sortedYears
    .map((year, index) => {
      const previous = sortedYears[index - 1];
      const diffYen =
        previous && isNumber(year.average_wage_monthly_yen) && isNumber(previous.average_wage_monthly_yen)
          ? year.average_wage_monthly_yen - previous.average_wage_monthly_yen
          : null;
      const diffRate =
        previous && isNumber(year.average_wage_monthly_yen) && isNumber(previous.average_wage_monthly_yen) && previous.average_wage_monthly_yen > 0
          ? year.average_wage_monthly_yen / previous.average_wage_monthly_yen - 1
          : null;
      const yearSourceUrl = safeExternalUrl(year.source_url);
      return `
        <article class="history-card${index === sortedYears.length - 1 ? " is-latest" : ""}${String(year.calculation_method_label ?? "").includes("変更") ? " is-caution" : ""}">
          <p class="section-kicker">${escapeHtml(year.fiscal_year_label ?? "-")}</p>
          <h3>平均工賃/月</h3>
          <strong>${escapeHtml(formatOfficialYen(year.average_wage_monthly_yen))}</strong>
          <p>北海道内 ${formatCount(year.facility_count ?? 0)} 事業所 / 平均工賃/時間 ${escapeHtml(formatOfficialYen(year.average_wage_hourly_yen))}</p>
          <div class="history-chip-row">
            <span class="history-chip${diffRate != null && diffRate > 0 ? " is-positive" : ""}">${escapeHtml(
              previous ? `前年度比 ${formatSignedPercent(diffRate)}` : "比較起点"
            )}</span>
            <span class="history-chip">${escapeHtml(previous && diffYen != null ? `前年差 ${formatSignedYen(diffYen)}` : String(year.calculation_method_label ?? "-"))}</span>
            ${
              previous && String(year.calculation_method_label ?? "").includes("変更")
                ? '<span class="history-chip is-caution">算定方法変更あり</span>'
                : `<span class="history-chip">${escapeHtml(String(year.calculation_method_label ?? "-"))}</span>`
            }
          </div>
          ${yearSourceUrl ? `<a class="history-link" href="${escapeAttribute(yearSourceUrl)}" target="_blank" rel="noreferrer">概要PDFを見る</a>` : ""}
        </article>
      `;
    })
    .join("");

  const pageUrl = safeExternalUrl(history.source_page_url);
  foot.innerHTML = `
    <article class="history-note">
      <strong>比較の読み方</strong>
      <p>${escapeHtml(history.caution ?? "年度比較の前提差があるため、増減は制度変更も含めて読む。")}</p>
      <div class="history-links">
        ${pageUrl ? `<a class="history-link" href="${escapeAttribute(pageUrl)}" target="_blank" rel="noreferrer">北海道公式ページ</a>` : ""}
        ${sortedYears
          .map((year) => {
            const yearSourceUrl = safeExternalUrl(year.source_url);
            return yearSourceUrl
              ? `<a class="history-link" href="${escapeAttribute(yearSourceUrl)}" target="_blank" rel="noreferrer">${escapeHtml(year.fiscal_year_label ?? "-")} 概要</a>`
              : "";
          })
          .join("")}
      </div>
    </article>
  `;
}

function renderQuality(dashboard, records = state.records) {
  const issuesRoot = document.getElementById("issuesList");
  const notesRoot = document.getElementById("notesList");
  const issues = dashboard.issues ?? [];
  const notes = dashboard.notes ?? [];
  const wageStats = computeLocalStats(numericValues(records, "average_wage_yen"));
  const wamMatchSummary = dashboard.analytics?.wam_match_summary ?? {};
  const focusLabel = wamMatchSummary.focus_label ?? "札幌市・小樽市";
  const focusRecordCount = wamMatchSummary.focus_record_count ?? 0;

  const summaryCards = [
    `<article class="note-card"><strong>表示対象</strong><p>工賃データは北海道全域 ${formatCount(
      records.length
    )} 件を表示している。札幌市・小樽市以外は参考比較で見られる。</p></article>`,
    `<article class="note-card"><strong>人員詳細あり とは</strong><p>福祉医療機構の公開情報と結び付いた事業所で、${escapeHtml(focusLabel)}レコード ${formatCount(
      wamMatchSummary.matched_record_count ?? 0
    )} / ${formatCount(focusRecordCount)} 件が対象である。</p></article>`,
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
  if (state.filters.noufukuOnly) filters.push("農福連携あり");
  if (state.filters.wamOnly) filters.push("人員詳細ありのみ");

  document.getElementById("activeFilterSummary").textContent = filters.length
    ? `${formatCount(records.length)} 件を表示中。`
    : `${formatCount(records.length)} 件を表示中。追加条件なし。`;
  const tagsRoot = document.getElementById("activeFilterTags");
  if (tagsRoot) {
    const tags = [
      ...(state.filters.municipality === "all" ? ["対象: 北海道全域"] : []),
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

  // 月額1.5万円ライン
  const targetRate = benchmarkWageRate(records);

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
      title: "🎯 月額1.5万円ライン",
      body: targetRate != null
        ? `月額 ${formatCount(REFERENCE_WAGE_LINE_YEN)} 円以上の事業所は ${formatPercent(targetRate)}。中央値は ${formatMaybeYen(wageStats.median)} で${wageStats.median && wageStats.median >= REFERENCE_WAGE_LINE_YEN ? "ラインを上回っている" : "まだラインに届いていない"}。`
        : "工賃データなし。",
    },
    {
      title: "💰 経営者向け: 利用率改善の収益効果",
      body: avgRevPerPoint != null
        ? `利用率70%未満の ${formatCount(lowUtilRecords.length)} 事業所の場合、利用率を1%改善すると平均で月 ${formatMaybeYen(avgRevPerPoint)} の増収。10%改善なら月 ${formatMaybeYen(avgRevPerPoint * 10)} の差になる。`
        : "利用率データが不足。",
    },
    {
      title: "👥 職員配置の立ち位置",
      body: matched.length
        ? `最も手厚い6:1水準が ${formatCount(tier6Count)} 件、標準より上の7.5:1水準が ${formatCount(tier75Count)} 件。${tier6Count === 0 ? "最上位の配置まで届いている事業所はまだない。人員増強で報酬単価アップを狙える余地がある。" : ""}`
        : "職員配置まで見える事業所がない。",
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
    ["growthList", "workShortageList", "fixList", "highHighList"].forEach((id) => {
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

  const workShortageCandidates = records
    .filter((record) => hasWorkShortageRisk(record))
    .sort((left, right) => workShortageScore(right) - workShortageScore(left))
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

  rootSummary.textContent = `高工賃・低利用率 ${formatCount(growthCandidates.length)} 件 / 仕事不足 ${formatCount(
    workShortageCandidates.length
  )} 件 / 低工賃・低利用率・人員過剰 ${formatCount(
    fixCandidates.length
  )} 件 / 高工賃・高利用率 ${formatCount(highHighCandidates.length)} 件`;

  renderStrategyList("growthList", growthCandidates, "高工賃・低利用率の事業所はまだない", buildGrowthReason);
  renderStrategyList("workShortageList", workShortageCandidates, "仕事が少なくて困っていそうな事業所はまだない", buildWorkShortageReason);
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
        <article class="strategy-item">
          <span class="strategy-kicker">${escapeHtml(record.municipality ?? "-")} / No.${escapeHtml(record.office_no ?? "-")}</span>
          <strong>${escapeHtml(record.office_name ?? "-")}</strong>
          <span>${renderCorporationTrigger(record.corporation_name)}</span>
          <p>${escapeHtml(buildReason(record))}</p>
          <div class="strategy-actions">
            <button class="table-link" data-select-office="${escapeAttribute(record.office_no ?? "")}" type="button">詳細</button>
          </div>
        </article>
      `
    )
    .join("");
}

function renderStats(records) {
  const matched = matchedRecords(records);
  const wages = numericValues(records, "average_wage_yen");
  const utilizationValues = numericValues(records, "daily_user_capacity_ratio");
  const wageStats = computeLocalStats(wages);
  const utilizationStats = computeLocalStats(utilizationValues);
  const nearRankUp = records.filter((r) => {
    const info = getWageTierUpInfo(r);
    return info && info.next && info.gapYen > 0 && info.gapYen <= 3000;
  }).length;
  const targetRate = benchmarkWageRate(records);
  const lowUtilCount = records.filter((record) => isNumber(record.daily_user_capacity_ratio) && record.daily_user_capacity_ratio < 0.7).length;
  const workShortageCount = records.filter((record) => hasWorkShortageRisk(record)).length;
  const growthOpportunityCount = records.filter((record) => record.market_position_quadrant === "高工賃 × 低稼働").length;
  const benchmarkOfficeCount = records.filter((record) => record.market_position_quadrant === "高工賃 × 高稼働").length;
  const fixPriorityCount = matched.filter(
    (record) =>
      record.market_position_quadrant === "低工賃 × 低稼働" &&
      record.wam_staffing_outlier_flag === "high"
  ).length;
  const wamMatchSummary = state.dashboard?.analytics?.wam_match_summary ?? {};
  const focusLabel = wamMatchSummary.focus_label ?? "札幌市・小樽市";
  const focusMunicipalities = new Set(wamMatchSummary.focus_municipalities ?? []);
  const focusRecords = focusMunicipalities.size
    ? records.filter((record) => focusMunicipalities.has(String(record.municipality ?? "")))
    : [];

  // 人員配置基準（R6年度 3段階）
  const staffingCritical = matched.filter((r) => {
    const c = getStaffingComplianceLevel(r);
    return c && (c.level === "critical" || c.level === "tier_10_1");
  }).length;
  const staffing6to1 = matched.filter((r) => {
    const c = getStaffingComplianceLevel(r);
    return c && c.level === "tier_6_1";
  }).length;
  const matchedCoverageHint = focusRecords.length
    ? `${focusLabel}の ${formatCount(focusRecords.length)} 件中 ${formatPercent(matched.length / focusRecords.length)}`
    : `${focusLabel}は現在の表示条件に含まれていない`;
  const staffing6Hint = matched.length
    ? `人員詳細ありの ${formatPercent(staffing6to1 / matched.length)} / 6:1の最上位水準`
    : "人員詳細ありの事業所で判定";
  const staffingCriticalHint = matched.length
    ? `人員詳細ありの ${formatPercent(staffingCritical / matched.length)} / 10:1に近く欠員に注意`
    : "人員詳細ありの事業所で判定";
  const historicalTrend = getHistoricalTrendSummary();

  const cards = [
    { label: "表示件数", value: formatCount(records.length), hint: `全 ${formatCount(state.records.length)} 件中` },
    { label: "中央の工賃", value: formatMaybeYen(wageStats.median), hint: "半分の事業所がこの金額以上 / 以下" },
    historicalTrend
      ? {
          label: "全道平均の3年推移",
          value: formatSignedPercent(historicalTrend.totalGrowthRate),
          hint: `${historicalTrend.baseline.fiscal_year_label}→${historicalTrend.latest.fiscal_year_label} / ${formatOfficialYen(historicalTrend.baseline.average_wage_monthly_yen)}→${formatOfficialYen(historicalTrend.latest.average_wage_monthly_yen)}。令和5から算定方法変更あり`,
          action: "history",
          tone: "history",
          cta: "3年推移を見る",
        }
      : null,
    { label: "中央の利用率", value: formatPercent(utilizationStats.median), hint: "定員に対する平均利用人数の真ん中" },
    { label: "1.5万円ライン超え", value: formatPercent(targetRate), hint: `月額 ${formatCount(REFERENCE_WAGE_LINE_YEN)} 円以上。報酬区分4以上の入口` },
    { label: "次の区分が近い", value: formatCount(nearRankUp), hint: "あと3,000円以内で報酬区分アップ" },
    { label: "利用率70%未満", value: formatCount(lowUtilCount), hint: "集客・定着の立て直し候補" },
    { label: "職員配置まで見える", value: formatCount(matched.length), hint: matchedCoverageHint },
    { label: "最上位の配置", value: formatCount(staffing6to1), hint: staffing6Hint },
    { label: "最低ライン運営", value: formatCount(staffingCritical), hint: staffingCriticalHint },
    { label: "稼働改善で伸びる", value: formatCount(growthOpportunityCount), hint: "高工賃だが利用率が低い" },
    { label: "立て直し優先", value: formatCount(fixPriorityCount), hint: "低工賃・低利用率・人員厚め" },
    {
      label: "仕事が少なくて困っていそう",
      value: formatCount(workShortageCount),
      hint: "作業量・受注先の確認を先にしたい事業所",
      action: "work-shortage",
      tone: "alert",
      cta: "一覧を見る",
    },
    {
      label: "好事例として見たい",
      value: formatCount(benchmarkOfficeCount),
      hint: "高工賃・高利用率の事業所",
      action: "benchmark",
      tone: "good",
      cta: "一覧を見る",
    },
  ].filter(Boolean);

  document.getElementById("statsGrid").innerHTML = cards
    .map(
      (card) => `
        <${card.action ? "button" : "article"} class="stat-card${card.action ? ` stat-card-action stat-card-${escapeAttribute(card.tone ?? "accent")}` : ""}"${card.action ? ` data-stat-action="${escapeAttribute(card.action)}" type="button"` : ""}>
          <p>${escapeHtml(card.label)}</p>
          <strong>${escapeHtml(card.value)}</strong>
          <em>${escapeHtml(card.hint)}</em>
          ${card.action ? `<span class="stat-card-cta">${escapeHtml(card.cta ?? "一覧を見る")}</span>` : ""}
        </${card.action ? "button" : "article"}>
      `
    )
    .join("");
}

function getHistoricalTrendSummary(dashboard = state.dashboard) {
  const history = dashboard?.analytics?.historical_b_type_overview ?? HOKKAIDO_B_TYPE_HISTORY;
  const years = Array.isArray(history?.years) ? history.years : [];
  if (!years.length) return null;

  const sortedYears = [...years].sort((left, right) =>
    String(left.fiscal_year_label ?? "").localeCompare(String(right.fiscal_year_label ?? ""), "ja")
  );
  const latest = sortedYears[sortedYears.length - 1];
  const baseline = sortedYears[0];
  const totalGrowthRate =
    isNumber(latest?.average_wage_monthly_yen) &&
    isNumber(baseline?.average_wage_monthly_yen) &&
    baseline.average_wage_monthly_yen > 0
      ? latest.average_wage_monthly_yen / baseline.average_wage_monthly_yen - 1
      : null;
  if (totalGrowthRate == null) return null;

  return { history, sortedYears, baseline, latest, totalGrowthRate };
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
    xLabel: "定員に対する支援職員",
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
        <article
          class="anomaly-card"
          data-flag="${escapeHtml(record.wage_outlier_flag ?? record.wam_staffing_outlier_flag ?? "none")}"
        >
          <div>
            <p class="section-kicker">${escapeHtml(record.municipality ?? "-")} / No.${escapeHtml(record.office_no ?? "-")}</p>
            <h3>${escapeHtml(record.office_name ?? "-")}</h3>
            ${renderCorporationSubtitle(record, false)}
          </div>
          <div class="anomaly-meta">
            ${record.wage_outlier_flag ? `<span class="metric-chip">${escapeHtml(`工賃 ${labelForSelect(record.wage_outlier_flag)}`)}</span>` : ""}
            ${record.wam_staffing_outlier_flag ? `<span class="metric-chip">${escapeHtml(`人員配置 ${staffingLevelLabel(record.wam_staffing_outlier_flag)}`)}</span>` : ""}
            ${(record.daily_user_capacity_ratio ?? 0) > 1.5 ? `<span class="metric-chip" style="background:rgba(220,38,38,0.1);color:#dc2626">利用率過大 ${formatPercent(record.daily_user_capacity_ratio)}</span>` : ""}
            <span class="metric-chip">${escapeHtml(`工賃 ${formatWageText(record.average_wage_yen)}`)}</span>
            <span class="metric-chip">${escapeHtml(`定員に対する支援職員 ${formatPercent(record.wam_key_staff_fte_per_capacity)}`)}</span>
          </div>
          <p>平均との差 ${escapeHtml(formatRatio(record.wage_ratio_to_overall_mean))} / 定員に対する支援職員 ${escapeHtml(formatPercent(record.wam_key_staff_fte_per_capacity))}</p>
          <p>${escapeHtml(labelForSelect(record.wam_staffing_efficiency_quadrant ?? "人員配置の分類なし"))} / ${escapeHtml(record.wam_primary_activity_type ?? "活動不明")}</p>
          <div class="anomaly-actions">
            <button class="table-link" data-select-office="${escapeAttribute(record.office_no ?? "")}" type="button">詳細</button>
          </div>
        </article>
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

  if (openButton) openButton.disabled = false;
  if (focusButton) focusButton.disabled = false;

  if (summaryRoot) {
    summaryRoot.innerHTML = `
      <article class="selected-office-card">
        <div class="selected-office-head">
          <div>
            <p class="section-kicker">${escapeHtml(record.municipality ?? "-")} / No.${escapeHtml(record.office_no ?? "-")}</p>
            <h3>${escapeHtml(record.office_name ?? "-")}</h3>
            ${renderCorporationSubtitle(record)}
          </div>
          <div class="selected-office-status">
            ${matchBadge(record.wam_match_status, record.wam_match_confidence)}
            ${record.wage_outlier_flag ? outlierBadge(record.wage_outlier_flag, record.wage_outlier_severity) : ""}
          </div>
        </div>
        <div class="selected-office-metrics">
          <span class="metric-chip">${escapeHtml(`平均工賃 ${formatWageText(record.average_wage_yen)}`)}</span>
          <span class="metric-chip">${escapeHtml(`利用率 ${formatPercent(record.daily_user_capacity_ratio)}`)}</span>
          <span class="metric-chip">${escapeHtml(`定員に対する支援職員 ${formatPercent(record.wam_key_staff_fte_per_capacity)}`)}</span>
        </div>
        <p class="selected-office-note">${escapeHtml(actionNotes[0] ?? "詳しい内訳は詳細を開いて確認できる。")}</p>
        <div class="selected-office-actions">
          ${officeUrl ? `<a class="ghost-button link-button" href="${escapeAttribute(officeUrl)}" target="_blank" rel="noreferrer">事業所ページ</a>` : ""}
        </div>
      </article>
    `;
  }

  if (!dialogRoot) return;

  const tierInfo = getWageTierUpInfo(record);
  const staffComp = getStaffingComplianceLevel(record);
  const comparisonContext = getDetailComparisonContext(record);
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
        ${renderCorporationSubtitle(record)}
      </div>
      <div class="detail-cta">
        ${officeUrl ? `<a class="solid-button link-button" href="${escapeAttribute(officeUrl)}" target="_blank" rel="noreferrer">事業所ページ</a>` : ""}
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
        "職員配置の立ち位置",
        staffComp ? staffComp.label : record.wam_match_status === "matched" ? "-" : "詳細なし",
        staffComp ? `${staffComp.qualifiedTier} / 職員数（常勤換算） ${formatFte(staffComp.staffFte)}` : ""
      )}
      ${detailKpi(
        "利用者1人あたり職員",
        isNumber(perUser) ? `${ratioFormatter.format(perUser)} 人` : "-",
        record.wam_match_status === "matched" ? `定員に対する支援職員 ${formatPercent(record.wam_key_staff_fte_per_capacity)}` : "人員詳細なし"
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
        <div class="detail-major-grid">
          ${renderDetailMajorStatItems(record, comparisonContext)}
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
        <h3>📊 ピアベンチマーク</h3>
        <ul class="detail-list">
          ${peer.activity ? `<li>【${escapeHtml(peer.activity.type)}】${formatCount(peer.activity.count)}事業所中 ${formatCount(peer.activity.rank)}位（中央値 ${formatMaybeYen(peer.activity.median)}）</li>` : "<li>主活動のピアデータなし</li>"}
          ${peer.capacity ? `<li>【${escapeHtml(peer.capacity.band)}】${formatCount(peer.capacity.count)}事業所中 ${formatCount(peer.capacity.rank)}位（中央値 ${formatMaybeYen(peer.capacity.median)}）</li>` : "<li>定員帯のピアデータなし</li>"}
          <li>市町村平均との差: ${escapeHtml(formatRatio(record.wage_ratio_to_municipality_mean))}</li>
          <li>法人種別平均との差: ${escapeHtml(formatRatio(record.wage_ratio_to_corporation_type_mean))}</li>
          <li>工賃と利用の位置: ${escapeHtml(labelForSelect(record.market_position_quadrant ?? "-"))}</li>
        </ul>
      </article>
      <article class="detail-card">
        <h3>💰 経営者向けシミュレーション</h3>
        <ul class="detail-list">
          ${tierInfo?.next ? `<li>ランクアップに必要な工賃改善: あと ${formatCount(tierInfo.gapYen)} 円</li>` : "<li>最上位の報酬算定区分</li>"}
          ${tierInfo?.revenueImpact ? `<li>ランクアップ時の増収: 月 ${formatMaybeYen(tierInfo.revenueImpact)}</li>` : ""}
          ${revPerPoint ? `<li>利用率1%改善: 月 ${formatMaybeYen(revPerPoint)} の増収</li>` : ""}
          ${missedAddons.length ? missedAddons.map((m) => `<li>💡 ${escapeHtml(m.name)}: ${escapeHtml(m.hint)}</li>`).join("") : "<li>主要加算の取り漏れは確認されない</li>"}
        </ul>
      </article>
      <article class="detail-card">
        <h3>👥 運営体制</h3>
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
        <h3>⚠️ 専門家の注視ポイント</h3>
        <ul class="detail-list">
          ${actionNotes.map((note) => `<li>${escapeHtml(note)}</li>`).join("")}
        </ul>
      </article>
      <article class="detail-card">
        <h3>📋 基本情報</h3>
        <ul class="detail-list">
          <li>住所: ${escapeHtml(composeAddress(record))}</li>
          <li>電話: ${escapeHtml(record.wam_office_phone ?? "-")}</li>
          <li>事業所番号: ${escapeHtml(record.wam_office_number ?? "-")}</li>
          <li>公開情報の定員: ${escapeHtml(formatNullable(record.wam_office_capacity))}</li>
          <li>新設: ${escapeHtml(record.is_new_office ? "あり" : "なし")}</li>
          <li>備考: ${escapeHtml(record.remarks ?? "-")}</li>
        </ul>
      </article>
    </div>
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
    tableBody.innerHTML = `<tr><td colspan="12">${document.getElementById("emptyStateTemplate").innerHTML}</td></tr>`;
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
          <td>${renderCorporationTrigger(record.corporation_name)}</td>
          <td class="numeric">${formatWage(record.average_wage_yen, record.average_wage_error)}</td>
          <td class="numeric">${formatRatio(record.wage_ratio_to_overall_mean)}</td>
          <td class="numeric">${formatPercent(record.daily_user_capacity_ratio)}</td>
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
                ${renderCorporationSubtitle(record, false)}
              </div>
              <button class="table-link" data-select-office="${escapeHtml(record.office_no)}" type="button">詳細</button>
            </div>
            <div class="record-card-metrics">
              <span class="metric-chip">${escapeHtml(`工賃 ${formatWageText(record.average_wage_yen)}`)}</span>
              <span class="metric-chip">${escapeHtml(`利用率 ${formatPercent(record.daily_user_capacity_ratio)}`)}</span>
              <span class="metric-chip">${escapeHtml(`定員に対する支援職員 ${formatPercent(record.wam_key_staff_fte_per_capacity)}`)}</span>
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

function workShortageScore(record) {
  const wageRatio = record.wage_ratio_to_overall_mean ?? 1;
  const municipalityRatio = record.wage_ratio_to_municipality_mean ?? wageRatio;
  const utilization = record.daily_user_capacity_ratio ?? 0.7;
  const staffingHeavy =
    record.wam_staffing_efficiency_quadrant === "低工賃 × 厚い人員" ||
    record.wam_staffing_outlier_flag === "high";

  return (
    Math.max(0.95 - wageRatio, 0) * 80 +
    Math.max(0.95 - municipalityRatio, 0) * 60 +
    Math.max(0.7 - utilization, 0) * 70 +
    (staffingHeavy ? 18 : 0)
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
  )} / 定員に対する支援職員 ${formatPercent(record.wam_key_staff_fte_per_capacity)}`;
}

function buildFixReason(record) {
  const workShortageNote = hasWorkShortageRisk(record) ? " / 仕事不足の可能性" : "";
  return `市町村平均との差 ${formatRatio(record.wage_ratio_to_municipality_mean)} / 利用率 ${formatPercent(
    record.daily_user_capacity_ratio
  )} / 人員配置 ${labelForSelect(record.wam_staffing_efficiency_quadrant ?? "-")}${workShortageNote}`;
}

function buildWorkShortageReason(record) {
  const parts = [
    `全道平均との差 ${formatRatio(record.wage_ratio_to_overall_mean)}`,
    `市町村平均との差 ${formatRatio(record.wage_ratio_to_municipality_mean)}`,
    `利用率 ${formatPercent(record.daily_user_capacity_ratio)}`,
  ];
  if (record.wam_staffing_efficiency_quadrant) {
    parts.push(`人員配置 ${labelForSelect(record.wam_staffing_efficiency_quadrant)}`);
  }
  return `${parts.join(" / ")}。作業量・受注先・施設外就労の確認候補。`;
}

function buildHighHighReason(record) {
  return `市町村平均との差 ${formatRatio(record.wage_ratio_to_municipality_mean)} / 利用率 ${formatPercent(
    record.daily_user_capacity_ratio
  )}`;
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
    notes.push(`【職員配置】${staffing.label}。${staffing.qualifiedTier}。${staffing.advice}`);
  } else if (record.wam_match_status !== "matched") {
    notes.push("【職員配置】公開情報と紐づいていないため、配置の立ち位置はまだ評価できない。");
  }

  // 3. 月額1.5万円ライン
  if (isNumber(record.average_wage_yen)) {
    if (record.average_wage_yen < REFERENCE_WAGE_LINE_YEN) {
      const gap = REFERENCE_WAGE_LINE_YEN - record.average_wage_yen;
      notes.push(`【工賃水準】月額 ${formatCount(REFERENCE_WAGE_LINE_YEN)} 円ライン未達。あと ${formatCount(Math.round(gap))} 円の向上余地がある。工賃向上計画の見直しを推奨。`);
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
  const usesAllRecords = filteredCount === state.records.length;

  if (filteredCount >= 5 && recordInFiltered && !usesAllRecords) {
    return {
      records: state.filteredRecords,
      label: `現在の絞り込み結果 ${formatCount(filteredCount)}件`,
      note: "平均との差・中央値差は現在の絞り込み結果を母集団にしている。絞り込みを変えると比較値も変わる。",
    };
  }

  if (filteredCount > 0 && filteredCount < 5 && recordInFiltered && !usesAllRecords) {
    return {
      records: state.records,
      label: `全道 ${formatCount(state.records.length)}件`,
      note: `現在の絞り込みが ${formatCount(filteredCount)} 件と少ないため、平均との差・中央値差は全道比較に切り替えている。`,
    };
  }

  if (!recordInFiltered) {
    return {
      records: state.records,
      label: `全道 ${formatCount(state.records.length)}件`,
      note: "この事業所は現在の一覧外のため、全道を母集団にして比較している。",
    };
  }

  return {
    records: state.records,
    label: `全道 ${formatCount(state.records.length)}件`,
    note: "平均との差・中央値差は全道比較で表示している。",
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
      key: "wam_key_staff_fte_per_capacity",
      label: "定員に対する支援職員",
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

function detailMajorStatCard(label, value, note, tone = "neutral") {
  return `
    <article class="detail-major-item detail-major-${escapeAttribute(tone)}">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(value)}</strong>
      <em>${escapeHtml(note)}</em>
    </article>
  `;
}

function renderDetailMajorStatItems(record, comparisonContext) {
  const wageStats = computeLocalStats(numericValues(comparisonContext.records, "average_wage_yen"));
  const utilizationStats = computeLocalStats(numericValues(comparisonContext.records, "daily_user_capacity_ratio"));
  const tierInfo = getWageTierUpInfo(record);
  const staffing = getStaffingComplianceLevel(record);
  const historicalTrend = getHistoricalTrendSummary();
  const hitsReferenceWage = isNumber(record.average_wage_yen) && record.average_wage_yen >= REFERENCE_WAGE_LINE_YEN;
  const lowUtilization = isNumber(record.daily_user_capacity_ratio) && record.daily_user_capacity_ratio < 0.7;
  const growthOpportunity = record.market_position_quadrant === "高工賃 × 低稼働";
  const fixPriority =
    record.market_position_quadrant === "低工賃 × 低稼働" &&
    record.wam_match_status === "matched" &&
    record.wam_staffing_outlier_flag === "high";
  const benchmarkCandidate = record.market_position_quadrant === "高工賃 × 高稼働";
  const workShortage = hasWorkShortageRisk(record);

  const cards = [
    detailMajorStatCard(
      "中央の工賃",
      isNumber(record.average_wage_yen) && isNumber(wageStats.median) ? formatSignedYen(record.average_wage_yen - wageStats.median) : "-",
      isNumber(wageStats.median) ? `中央値 ${formatMaybeYen(wageStats.median)}` : "比較できる工賃データなし",
      isNumber(record.average_wage_yen) && isNumber(wageStats.median) && record.average_wage_yen >= wageStats.median ? "good" : "alert"
    ),
    detailMajorStatCard(
      "全道平均の3年推移",
      historicalTrend ? formatSignedPercent(historicalTrend.totalGrowthRate) : "-",
      historicalTrend
        ? `${historicalTrend.baseline.fiscal_year_label}→${historicalTrend.latest.fiscal_year_label}`
        : "全道推移データなし",
      "info"
    ),
    detailMajorStatCard(
      "中央の利用率",
      isNumber(record.daily_user_capacity_ratio) && isNumber(utilizationStats.median)
        ? formatSignedPercentPoint(record.daily_user_capacity_ratio - utilizationStats.median)
        : "-",
      isNumber(utilizationStats.median) ? `中央値 ${formatPercent(utilizationStats.median)}` : "比較できる利用率データなし",
      isNumber(record.daily_user_capacity_ratio) && isNumber(utilizationStats.median) && record.daily_user_capacity_ratio >= utilizationStats.median
        ? "good"
        : "alert"
    ),
    detailMajorStatCard(
      "1.5万円ライン",
      hitsReferenceWage ? "達成" : "未達",
      hitsReferenceWage
        ? `${formatCount(REFERENCE_WAGE_LINE_YEN)}円を上回る`
        : isNumber(record.average_wage_yen)
          ? `あと ${formatCount(Math.round(REFERENCE_WAGE_LINE_YEN - record.average_wage_yen))}円`
          : "工賃データなし",
      hitsReferenceWage ? "good" : "alert"
    ),
    detailMajorStatCard(
      "次の区分",
      tierInfo?.next ? `あと ${formatCount(tierInfo.gapYen)}円` : tierInfo?.current ? "最上位" : "-",
      tierInfo?.next
        ? `${tierInfo.current.label}→${tierInfo.next.label}`
        : tierInfo?.current
          ? `${tierInfo.current.label}を維持`
          : "報酬区分を計算できない",
      tierInfo?.next && tierInfo.gapYen <= 3000 ? "good" : "neutral"
    ),
    detailMajorStatCard(
      "利用率70%ライン",
      lowUtilization ? "未達" : "クリア",
      isNumber(record.daily_user_capacity_ratio) ? `現在 ${formatPercent(record.daily_user_capacity_ratio)}` : "利用率データなし",
      lowUtilization ? "alert" : "good"
    ),
    detailMajorStatCard(
      "職員配置詳細",
      record.wam_match_status === "matched" ? "あり" : "なし",
      record.wam_match_status === "matched"
        ? `${matchConfidenceLabel(record.wam_match_confidence)}で一致`
        : "WAM詳細未一致",
      record.wam_match_status === "matched" ? "good" : "neutral"
    ),
    detailMajorStatCard(
      "最上位の配置",
      staffing?.level === "tier_6_1" ? "該当" : staffing ? "非該当" : "未確認",
      staffing?.level === "tier_6_1" ? staffing.qualifiedTier : staffing ? staffing.label : "人員詳細なし",
      staffing?.level === "tier_6_1" ? "good" : "neutral"
    ),
    detailMajorStatCard(
      "最低ライン運営",
      staffing?.level === "tier_10_1" ? "該当" : staffing?.level === "critical" ? "要確認" : staffing ? "非該当" : "未確認",
      staffing?.level === "tier_10_1"
        ? staffing.qualifiedTier
        : staffing?.level === "critical"
          ? staffing.advice
          : staffing
            ? "最低ラインより余裕あり"
            : "人員詳細なし",
      staffing?.level === "tier_10_1" || staffing?.level === "critical" ? "alert" : "neutral"
    ),
    detailMajorStatCard(
      "稼働改善で伸びる",
      growthOpportunity ? "該当" : "非該当",
      growthOpportunity ? "高工賃だが利用率改善の余地が大きい" : "典型パターンではない",
      growthOpportunity ? "good" : "neutral"
    ),
    detailMajorStatCard(
      "立て直し優先",
      fixPriority ? "該当" : "非該当",
      fixPriority ? "低工賃・低利用率・人員厚め" : "優先立て直しの型ではない",
      fixPriority ? "alert" : "neutral"
    ),
    detailMajorStatCard(
      "仕事不足の可能性",
      workShortage ? "あり" : "なし",
      workShortage ? "工賃と利用率の両面が弱い" : "強い不足シグナルはない",
      workShortage ? "alert" : "neutral"
    ),
    detailMajorStatCard(
      "好事例として見たい",
      benchmarkCandidate ? "該当" : "非該当",
      benchmarkCandidate ? "高工賃かつ高利用率" : "代表的な好事例型ではない",
      benchmarkCandidate ? "good" : "neutral"
    ),
  ];

  return cards.join("");
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
  const sapporoWardMatch = addressText.match(/札幌市([^\s0-9]+?区)/);
  if (sapporoWardMatch) {
    return sapporoWardMatch[1];
  }
  if (city) {
    const normalizedCity = city.replace(/^北海道/, "").trim();
    return normalizedCity === "札幌市" ? "札幌市（区未取得）" : normalizedCity;
  }
  return record.municipality === "札幌市" ? "札幌市（区未取得）" : record.municipality ?? null;
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
    ["wage_outlier_flag", "工賃水準"],
    ["wam_match_status", "人員詳細"],
    ["wam_welfare_staff_fte_total", "福祉職員（常勤換算）"],
    ["wam_key_staff_fte_per_capacity", "定員に対する支援職員"],
    ["wam_transport_available", "送迎"],
    ["wam_meal_support_addon", "食事加算"],
    ["wam_manager_multi_post", "管理者兼務"],
    ["wam_staffing_efficiency_quadrant", "工賃と人員配置の分類"],
    ["wam_staffing_outlier_flag", "人員配置"],
    ["derived_work_shortage_risk", "仕事不足の可能性"],
    ["wam_primary_activity_type", "主活動種別"],
    ["wam_office_number", "公開情報の事業所番号"],
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
      if (key === "derived_work_shortage_risk") {
        return escapeCsvCell(hasWorkShortageRisk(record) ? "あり" : "なし");
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
  link.download = "hokkaido-shuro-b-dashboard-filtered.csv";
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

function formatSignedPercent(value) {
  if (!isNumber(value)) return "-";
  return `${value >= 0 ? "+" : ""}${percentFormatter.format(value * 100)}%`;
}

function formatSignedPercentPoint(value) {
  if (!isNumber(value)) return "-";
  return `${value >= 0 ? "+" : ""}${percentFormatter.format(value * 100)}pt`;
}

function formatMaybeYen(value) {
  return isNumber(value) ? `${formatCount(Math.round(value))}円` : "-";
}

function formatOfficialYen(value) {
  if (!isNumber(value)) return "-";
  if (Number.isInteger(value)) {
    return `${formatCount(value)}円`;
  }
  return `${decimalFormatter.format(value)}円`;
}

function formatSignedYen(value) {
  if (!isNumber(value)) return "-";
  return `${value >= 0 ? "+" : ""}${formatCount(Math.round(value))}円`;
}

function formatSignedCount(value) {
  if (!isNumber(value)) return "-";
  return `${value >= 0 ? "+" : ""}${formatCount(value)}件`;
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
