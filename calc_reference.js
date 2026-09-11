// 独立実装のリファレンス計算(Excel数式の検算用) — Excel生成スクリプトとは別ロジックで再実装する
function lookupBracket(value, table) {
  // table: sorted ascending by lower bound, [lower, ...rest]
  let result = table[0];
  for (const row of table) {
    if (value >= row[0]) result = row; else break;
  }
  return result;
}

const kenpoLower = [0,63000,73000,83000,93000,101000,107000,114000,122000,130000,138000,146000,155000,165000,175000,185000,195000,210000,230000,250000,270000,290000,310000,330000,350000,370000,395000,425000,455000,485000,515000,545000,575000,605000,635000,665000,695000,730000,770000,810000,855000,905000,955000,1005000,1055000,1115000,1175000,1235000,1295000,1355000];
const kenpoStd   = [58000,68000,78000,88000,98000,104000,110000,118000,126000,134000,142000,150000,160000,170000,180000,190000,200000,220000,240000,260000,280000,300000,320000,340000,360000,380000,410000,440000,470000,500000,530000,560000,590000,620000,650000,680000,710000,750000,790000,830000,880000,930000,980000,1030000,1090000,1150000,1210000,1270000,1330000,1390000];
const kouseiLower = [0,93000,101000,107000,114000,122000,130000,138000,146000,155000,165000,175000,185000,195000,210000,230000,250000,270000,290000,310000,330000,350000,370000,395000,425000,455000,485000,515000,545000,575000,605000,635000];
const kouseiStd   = [88000,98000,104000,110000,118000,126000,134000,142000,150000,160000,170000,180000,190000,200000,220000,240000,260000,280000,300000,320000,340000,360000,380000,410000,440000,470000,500000,530000,560000,590000,620000,650000];

function vlookupStd(monthly, lowerArr, stdArr) {
  let idx = 0;
  for (let i = 0; i < lowerArr.length; i++) {
    if (monthly >= lowerArr[i]) idx = i; else break;
  }
  return stdArr[idx];
}

const taxBrackets = [
  [0, 0.05, 0],
  [1950000, 0.10, 97500],
  [3300000, 0.20, 427500],
  [6950000, 0.23, 636000],
  [9000000, 0.33, 1536000],
  [18000000, 0.40, 2796000],
  [40000000, 0.45, 4796000],
];

// 令和8年度税制改正(基礎控除の引上げ・給与所得控除の最低保障額引上げ)反映後の水準
// 出典: 国税庁「源泉所得税の改正のあらまし 令和8年4月」
const kisoIncomeTax = [[0,1040000],[4890000,670000],[6550000,620000],[23500000,480000],[24000000,320000],[24500000,160000],[25000000,0]];
const kisoJuminZei = 430000; // 個人住民税の基礎控除は令和8年度改正でも据置き(43万円)

// 配偶者控除・配偶者特別控除額(令和8・9年分)。行=配偶者の合計所得金額(下限)、列=[所得税(本人900万円以下/900万円超950万円以下/950万円超1000万円以下), 住民税(同左)]
// 出典: 国税庁パンフレット(配偶者特別控除額の表)+ 住民税は同表に対し住民税側の上限(33万/22万/11万)でキャップした額(市区町村公表の早見表と同一パターン)
const haiguTable = [
  // [配偶者合計所得 下限, 所得税tier0, 所得税tier1, 所得税tier2, 住民税tier0, 住民税tier1, 住民税tier2]
  [0,       380000, 260000, 130000, 330000, 220000, 110000],
  [950001,  360000, 240000, 120000, 330000, 220000, 110000],
  [1000001, 310000, 210000, 110000, 310000, 210000, 110000],
  [1050001, 260000, 180000, 90000,  260000, 180000, 90000],
  [1100001, 210000, 140000, 70000,  210000, 140000, 70000],
  [1150001, 160000, 110000, 60000,  160000, 110000, 60000],
  [1200001, 110000, 80000,  40000,  110000, 80000,  40000],
  [1250001, 60000,  40000,  20000,  60000,  40000,  20000],
  [1300001, 30000,  20000,  10000,  30000,  20000,  10000],
  [1330001, 0,      0,      0,      0,      0,      0],
];

// 特定親族特別控除(令和8年度税制改正で新設)。特定親族(19〜22歳、合計所得62万円超123万円以下)
// 本人の収入がある場合に適用。行=特定親族の合計所得金額(下限)、列=[所得税, 住民税]
// 出典: 国税庁パンフレット(特定親族特別控除額の表)。住民税は所得税額を45万円を上限にキャップした額
// (酒々井町公表の早見表と同一パターンで整合確認済み)。所得者本人の所得による段階はない(一律)。
const tokuteiShinzokuTable = [
  [620001,  630000, 450000],
  [850001,  610000, 450000],
  [900001,  510000, 450000],
  [950001,  410000, 410000],
  [1000001, 310000, 310000],
  [1050001, 210000, 210000],
  [1100001, 110000, 110000],
  [1150001, 60000,  60000],
  [1200001, 30000,  30000],
  [1230001, 0,      0],
];

function roundDown(v, digits) {
  const f = Math.pow(10, digits);
  return Math.floor(v * f + 1e-6) / f;
}

// 給与所得控除(令和8・9年分。最低保障額が65万円→74万円に引上げ)
// 国税庁パンフレット(1)ハの特例により、給与収入74万1,000円〜219万1,000円は「収入金額-74万円」
// (=控除額74万円固定)が(2)ロの通常式(190万円超220万円以下:収入×30%+8万円)に優先して適用される。
// これにより控除額は69.1万円〜220万円の間、実質的に74万円で連続する(2,200,000円ちょうどで
// 通常式に滑らかに接続: 2,200,000×30%+8万円=740,000)。219万1,000円〜220万円のごく狭い範囲の
// 段階的な端数調整(国税庁公表の給与所得の金額を1,000円単位で丸めた特例)は影響が僅少(数千円)
// のため簡略化し、220万円まで一律74万円として扱っています。
function kyuyoKoujoFunc(X) {
  if (X <= 2200000) return 740000;
  if (X <= 3600000) return X * 0.3 + 80000;
  if (X <= 6600000) return X * 0.2 + 440000;
  if (X <= 8500000) return X * 0.1 + 1100000;
  return 1950000;
}

// 配偶者の合計所得金額(給与収入から給与所得控除を差し引いた額)
function spouseTotalIncome(spouseIncome) {
  if (spouseIncome <= 0) return 0;
  return Math.max(spouseIncome - kyuyoKoujoFunc(spouseIncome), 0);
}

function filerTierIndex(kyuyoShotoku) {
  if (kyuyoShotoku <= 9000000) return 0;
  if (kyuyoShotoku <= 9500000) return 1;
  if (kyuyoShotoku <= 10000000) return 2;
  return -1;
}

function simulate(X, params) {
  const {
    profitBefore, ageOver40 = false, hasSpouse = true, spouseIncome = 0,
    generalDependents = 0, specificDependents = 0, elderlyDependents = 0,
    specificDependentsWithIncome = 0, specificDependentIncome = 0,
    kenpoRate = 0.1012, kaigoRate = 0.0162, kouseiRate = 0.183, kodomoRate = 0.0036, shienkinRate = 0.0023,
    corpRateLow = 0.2137, corpRateHigh = 0.3358,
    // 法人住民税の均等割。資本金等1,000万円以下・従業者50人以下の標準額(道府県2万円+市町村5万円)。
    // 資本金等1,000万円超1億円以下なら18万円など、規模に応じて変わるため上書きできるようにしている。
    houjinKintouwari = 70000,
    otherOfficersCost = 0, otherOfficersTakeHome = 0,
  } = params;

  const monthly = X / 12;
  const stdKenpo = X === 0 ? 0 : vlookupStd(monthly, kenpoLower, kenpoStd);
  const stdKousei = X === 0 ? 0 : vlookupStd(monthly, kouseiLower, kouseiStd);

  const kenpoRateTotal = kenpoRate + (ageOver40 ? kaigoRate : 0);
  const shienkin = Math.round(stdKenpo * shienkinRate / 2 * 12);
  const shakaihoKojin = Math.round(stdKenpo * kenpoRateTotal / 2 * 12) + Math.round(stdKousei * kouseiRate / 2 * 12) + shienkin;
  const shakaihoKaisha = Math.round(stdKenpo * kenpoRateTotal / 2 * 12) + Math.round(stdKousei * kouseiRate / 2 * 12) + Math.round(stdKousei * kodomoRate * 12) + shienkin;

  const kyuyoKoujo = kyuyoKoujoFunc(X);
  const kyuyoShotoku = Math.max(X - kyuyoKoujo, 0);

  const tier = filerTierIndex(kyuyoShotoku);
  const spouseIncomeTotal = spouseTotalIncome(spouseIncome);
  const haiguRow = (hasSpouse && tier >= 0) ? lookupBracket(spouseIncomeTotal, haiguTable) : null;
  const kiso = lookupBracket(kyuyoShotoku, kisoIncomeTax)[1];
  const haigu = haiguRow ? haiguRow[1 + tier] : 0;
  const haiguJumin_ = haiguRow ? haiguRow[4 + tier] : 0;

  const regularSpecificCount = Math.max(specificDependents - specificDependentsWithIncome, 0);
  const shinzokuTotalIncome = specificDependentsWithIncome > 0 ? spouseTotalIncome(specificDependentIncome) : 0;
  const shinzokuRow = specificDependentsWithIncome > 0 ? lookupBracket(shinzokuTotalIncome, tokuteiShinzokuTable) : null;
  const shinzokuPerPersonTax = (shinzokuRow && shinzokuTotalIncome >= 620001) ? shinzokuRow[1] : (specificDependentsWithIncome > 0 ? 630000 : 0);
  const shinzokuPerPersonJumin = (shinzokuRow && shinzokuTotalIncome >= 620001) ? shinzokuRow[2] : (specificDependentsWithIncome > 0 ? 450000 : 0);
  const tokuteiShinzokuTax = shinzokuPerPersonTax * specificDependentsWithIncome;
  const tokuteiShinzokuJumin = shinzokuPerPersonJumin * specificDependentsWithIncome;

  const fuyouTax = generalDependents * 380000 + regularSpecificCount * 630000 + tokuteiShinzokuTax + elderlyDependents * 580000;

  const shotokuKoujoGoukei = shakaihoKojin + kiso + haigu + fuyouTax;
  const kazeiShotoku = roundDown(Math.max(kyuyoShotoku - shotokuKoujoGoukei, 0), -3);

  const bracket = lookupBracket(kazeiShotoku, taxBrackets);
  const shotokuzei = Math.max(roundDown(kazeiShotoku * bracket[1] - bracket[2], 0), 0);
  const fukkou = roundDown(shotokuzei * 0.021, 0);
  const shotokuzeiTotal = shotokuzei + fukkou;

  const kisoJumin = kisoJuminZei;
  const haiguJumin = haiguJumin_;
  const fuyouJumin = generalDependents * 330000 + regularSpecificCount * 450000 + tokuteiShinzokuJumin + elderlyDependents * 450000;
  const juminKazeiShotoku = roundDown(Math.max(kyuyoShotoku - (shakaihoKojin + kisoJumin + haiguJumin + fuyouJumin), 0), -3);
  const juminShotokuwari = juminKazeiShotoku * 0.10;
  const juminKintouwari = 4000; // 個人住民税の均等割（標準税率）。令和6年度から道府県民税1,000円＋市町村民税3,000円＝4,000円（森林環境税1,000円は別建て）
  // 森林環境税(国税)。令和6年度から個人住民税の均等割と併せて年1,000円が賦課徴収される
  const shinrinKankyoZei = 1000;
  const juminTotal = juminShotokuwari + juminKintouwari + shinrinKankyoZei;

  const kojinFutanGoukei = shakaihoKojin + shotokuzeiTotal + juminTotal;
  const kojinTedori = X - kojinFutanGoukei;

  const kaishaCost = X + shakaihoKaisha;
  const rieki = profitBefore - kaishaCost - otherOfficersCost;
  // 所得に応じた部分(実効税率ベース)。赤字なら0
  let houjinShotokuBun;
  if (rieki <= 0) houjinShotokuBun = 0;
  else if (rieki <= 8000000) houjinShotokuBun = rieki * corpRateLow;
  else houjinShotokuBun = 8000000 * corpRateLow + (rieki - 8000000) * corpRateHigh;
  // 法人住民税の均等割は赤字でも課されるため、利益の有無にかかわらず加算する
  const houjinzei = houjinShotokuBun + houjinKintouwari;
  const naibuRyuho = rieki - houjinzei;
  const total = naibuRyuho + kojinTedori + otherOfficersTakeHome;
  const levyTotal = profitBefore - total;

  return { X, stdKenpo, stdKousei, shakaihoKojin, shakaihoKaisha, kyuyoKoujo, kyuyoShotoku,
    shotokuKoujoGoukei, kazeiShotoku, shotokuzei, fukkou, shotokuzeiTotal,
    juminShotokuwari, juminKintouwari, shinrinKankyoZei, juminTotal, kojinFutanGoukei, kojinTedori,
    kaishaCost, rieki, houjinShotokuBun, houjinKintouwari, houjinzei, naibuRyuho, total, levyTotal,
    otherOfficersCost, otherOfficersTakeHome };
}

// Node（検算スクリプト）とブラウザ（管理者用_計算過程.html）の両方から使えるようにする
if (typeof module === 'object' && module.exports) module.exports = { simulate };

if (typeof require !== 'undefined' && typeof module === 'object' && require.main === module) {
  const params = { profitBefore: 20000000, ageOver40: false, hasSpouse: true, spouseIncome: 0,
    generalDependents: 0, specificDependents: 0, elderlyDependents: 0 };
  for (const X of [3000000, 6000000, 10000000, 15000000, 24000000]) {
    console.log(X, JSON.stringify(simulate(X, params), null, 0));
  }
}
