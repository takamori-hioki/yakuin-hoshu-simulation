const ExcelJS = require('exceljs');
const path = require('path');

const OUT_PATH = process.argv[2] || path.join(__dirname, '役員報酬シミュレーション.xlsx');

const wb = new ExcelJS.Workbook();
wb.creator = 'AI所長';
wb.created = new Date();

const YEN = '#,##0"円"';
const YEN0 = '#,##0';
const PCT = '0.00%';
const N_OFFICERS = 5;

// ============================================================
// 入力シートの行レイアウト定数
// ============================================================
const COMPANY_PROFIT_ROW = 5;
const CAPITAL_ROW = 6;
const OFFICER_COUNT_ROW = 7;

function officerRows(k) {
  const header = 9 + k * 10;
  return {
    header,
    age: header + 1,
    spouse: header + 2,
    spouseIncome: header + 3,
    genDep: header + 4,
    specDep: header + 5,
    tokuteiCount: header + 6,
    tokuteiIncome: header + 7,
    elderDep: header + 8,
  };
}

const RATES_HEADER_ROW = 59;
const RATE_ROWS = { kenpo: 60, kaigo: 61, kousei: 62, kodomo: 63, shienkin: 64, corplow: 65, corphigh: 66, corpkintou: 67 };
const SIMRANGE_HEADER_ROW = 69;
const SIMRANGE_ROWS = { start: 70, end: 71, step: 72 };
const CAVEAT_ROW = 74;

// ============================================================
// 結果サマリーシートの行レイアウト定数
// ============================================================
const SUM_OFFICER_HEADER_ROW = 3;
function summaryOfficerRows(k) {
  const base = 5 + k * 3;
  return { input: base, tedori: base + 1, cost: base + 2 };
}
const SUM_AGG_HEADER_ROW = 21;
const SUM_AGG_ROWS = { totalX: 22, totalCost: 23, rieki: 24, houjinzei: 25, naibu: 26, totalTedori: 27, grandTotal: 28, rate: 29 };

const yellow = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };

function header(ws, row, text) {
  const cell = ws.getCell(`B${row}`);
  cell.value = text;
  cell.font = { bold: true, size: 12, color: { argb: 'FFFFFFFF' } };
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
  ws.mergeCells(`B${row}:D${row}`);
}

function label(ws, row, text, note) {
  ws.getCell(`B${row}`).value = text;
  ws.getCell(`B${row}`).font = { bold: false };
  ws.getCell(`B${row}`).alignment = { wrapText: true, vertical: 'top' };
  if (note) {
    ws.getCell(`D${row}`).value = note;
    ws.getCell(`D${row}`).font = { italic: true, size: 9, color: { argb: 'FF808080' } };
    ws.getCell(`D${row}`).alignment = { wrapText: true, vertical: 'top' };
  }
}

// ============================================================
// シート1: 入力
// ============================================================
const wsIn = wb.addWorksheet('入力', { properties: { tabColor: { argb: 'FF1F4E78' } } });
wsIn.columns = [
  { key: 'a', width: 4 },
  { key: 'b', width: 34 },
  { key: 'c', width: 18 },
  { key: 'd', width: 60 },
  { key: 'e', width: 18 },
];

wsIn.getCell('B1').value = '法人 役員報酬シミュレーション【入力シート】';
wsIn.getCell('B1').font = { bold: true, size: 16 };
wsIn.mergeCells('B1:D1');

wsIn.getCell('B2').value = 'このシートの黄色いセルに数値を入力すると、「シミュレーション1」〜「シミュレーション5」シートが自動で再計算されます。役員が複数いる場合は「役員の人数」を変更し、役員2以降の前提条件も入力してください。E列の「配偶者の給与収入」行に表示される金額は配偶者の合計所得金額の自動計算結果(入力不要)です。';
wsIn.getCell('B2').font = { italic: true, size: 10 };
wsIn.getCell('B2').alignment = { wrapText: true };
wsIn.mergeCells('B2:D3');

header(wsIn, 4, '【会社の前提条件】');
label(wsIn, COMPANY_PROFIT_ROW, '税引前利益（役員報酬控除前・年額）', '役員報酬を1円も払わなかった場合に法人に残る年間利益（減価償却等調整後）');
wsIn.getCell(`C${COMPANY_PROFIT_ROW}`).value = 20000000; wsIn.getCell(`C${COMPANY_PROFIT_ROW}`).numFmt = YEN; wsIn.getCell(`C${COMPANY_PROFIT_ROW}`).fill = yellow;
label(wsIn, CAPITAL_ROW, '資本金', '1億円以下の中小法人を想定した実効税率をデフォルト設定しています');
wsIn.getCell(`C${CAPITAL_ROW}`).value = 10000000; wsIn.getCell(`C${CAPITAL_ROW}`).numFmt = YEN; wsIn.getCell(`C${CAPITAL_ROW}`).fill = yellow;
label(wsIn, OFFICER_COUNT_ROW, '役員の人数', '2人以上の場合、下記の役員2〜5の前提条件欄も入力してください。人数を超える役員欄は自動的に集計から除外されます');
wsIn.getCell(`C${OFFICER_COUNT_ROW}`).value = 1; wsIn.getCell(`C${OFFICER_COUNT_ROW}`).fill = yellow;
wsIn.getCell(`C${OFFICER_COUNT_ROW}`).dataValidation = { type: 'list', allowBlank: false, formulae: ['"1,2,3,4,5"'] };

for (let k = 0; k < N_OFFICERS; k++) {
  const r = officerRows(k);
  header(wsIn, r.header, `【役員${k + 1}の前提条件】`);
  label(wsIn, r.age, '年齢区分', '40歳以上64歳以下は介護保険料（第2号被保険者）が加算されます');
  wsIn.getCell(`C${r.age}`).value = '40歳未満'; wsIn.getCell(`C${r.age}`).fill = yellow;
  wsIn.getCell(`C${r.age}`).dataValidation = { type: 'list', allowBlank: false, formulae: ['"40歳未満,40歳以上64歳以下"'] };
  label(wsIn, r.spouse, '配偶者の有無', '配偶者控除の判定に使用');
  wsIn.getCell(`C${r.spouse}`).value = '有'; wsIn.getCell(`C${r.spouse}`).fill = yellow;
  wsIn.getCell(`C${r.spouse}`).dataValidation = { type: 'list', allowBlank: false, formulae: ['"有,無"'] };
  label(wsIn, r.spouseIncome, '配偶者の給与収入（年間）', '136万円以下は配偶者控除、136万円超207万円以下は配偶者特別控除（段階的に減額）、207万円超は対象外（令和8・9年分）');
  wsIn.getCell(`C${r.spouseIncome}`).value = 0; wsIn.getCell(`C${r.spouseIncome}`).numFmt = YEN; wsIn.getCell(`C${r.spouseIncome}`).fill = yellow;
  wsIn.getCell(`E${r.spouseIncome}`).value = { formula: `IF(C${r.spouseIncome}<=0,0,MAX(C${r.spouseIncome}-IF(C${r.spouseIncome}<=2200000,740000,IF(C${r.spouseIncome}<=3600000,C${r.spouseIncome}*0.3+80000,IF(C${r.spouseIncome}<=6600000,C${r.spouseIncome}*0.2+440000,IF(C${r.spouseIncome}<=8500000,C${r.spouseIncome}*0.1+1100000,1950000)))),0))` };
  wsIn.getCell(`E${r.spouseIncome}`).numFmt = YEN;
  wsIn.getCell(`E${r.spouseIncome}`).font = { italic: true, size: 9, color: { argb: 'FF808080' } };
  label(wsIn, r.genDep, '一般の扶養親族の数', '16～18歳・23～69歳（同居老親等を除く）');
  wsIn.getCell(`C${r.genDep}`).value = 0; wsIn.getCell(`C${r.genDep}`).numFmt = YEN0; wsIn.getCell(`C${r.genDep}`).fill = yellow;
  label(wsIn, r.specDep, '特定扶養親族の数', '19～22歳（特定親族特別控除の対象は下記で別途入力）');
  wsIn.getCell(`C${r.specDep}`).value = 0; wsIn.getCell(`C${r.specDep}`).numFmt = YEN0; wsIn.getCell(`C${r.specDep}`).fill = yellow;
  label(wsIn, r.tokuteiCount, 'うち収入がある特定親族の人数', '合計所得62万円超123万円以下（給与収入136万円超197万円以下）で特定親族特別控除が段階的に適用されます（令和8年度改正で新設）。特定扶養親族の数を超えないようにしてください');
  wsIn.getCell(`C${r.tokuteiCount}`).value = 0; wsIn.getCell(`C${r.tokuteiCount}`).numFmt = YEN0; wsIn.getCell(`C${r.tokuteiCount}`).fill = yellow;
  label(wsIn, r.tokuteiIncome, 'その特定親族の給与収入（年間・全員同額と仮定）', '');
  wsIn.getCell(`C${r.tokuteiIncome}`).value = 0; wsIn.getCell(`C${r.tokuteiIncome}`).numFmt = YEN; wsIn.getCell(`C${r.tokuteiIncome}`).fill = yellow;
  wsIn.getCell(`E${r.tokuteiIncome}`).value = { formula: `IF(C${r.tokuteiIncome}<=0,0,MAX(C${r.tokuteiIncome}-IF(C${r.tokuteiIncome}<=2200000,740000,IF(C${r.tokuteiIncome}<=3600000,C${r.tokuteiIncome}*0.3+80000,IF(C${r.tokuteiIncome}<=6600000,C${r.tokuteiIncome}*0.2+440000,IF(C${r.tokuteiIncome}<=8500000,C${r.tokuteiIncome}*0.1+1100000,1950000)))),0))` };
  wsIn.getCell(`E${r.tokuteiIncome}`).numFmt = YEN;
  wsIn.getCell(`E${r.tokuteiIncome}`).font = { italic: true, size: 9, color: { argb: 'FF808080' } };
  label(wsIn, r.elderDep, '老人扶養親族の数（同居）', '70歳以上・同居老親等');
  wsIn.getCell(`C${r.elderDep}`).value = 0; wsIn.getCell(`C${r.elderDep}`).numFmt = YEN0; wsIn.getCell(`C${r.elderDep}`).fill = yellow;
}

header(wsIn, RATES_HEADER_ROW, '【料率等（編集可能・要検証・全役員共通）】');
label(wsIn, RATE_ROWS.kenpo, '健康保険料率（協会けんぽ・兵庫県）', '都道府県・年度により変動。デフォルトは兵庫県・令和8年3月分以降の目安値。最新の協会けんぽ公表料率で必ず検証してください');
wsIn.getCell(`C${RATE_ROWS.kenpo}`).value = 0.1012; wsIn.getCell(`C${RATE_ROWS.kenpo}`).numFmt = PCT; wsIn.getCell(`C${RATE_ROWS.kenpo}`).fill = yellow;
label(wsIn, RATE_ROWS.kaigo, '介護保険料率（40～64歳）', '全国一律。令和8年3月分以降の目安値。要検証');
wsIn.getCell(`C${RATE_ROWS.kaigo}`).value = 0.0162; wsIn.getCell(`C${RATE_ROWS.kaigo}`).numFmt = PCT; wsIn.getCell(`C${RATE_ROWS.kaigo}`).fill = yellow;
label(wsIn, RATE_ROWS.kousei, '厚生年金保険料率', '平成29年9月以降固定（労使合計）');
wsIn.getCell(`C${RATE_ROWS.kousei}`).value = 0.183; wsIn.getCell(`C${RATE_ROWS.kousei}`).numFmt = PCT; wsIn.getCell(`C${RATE_ROWS.kousei}`).fill = yellow;
label(wsIn, RATE_ROWS.kodomo, '子ども・子育て拠出金率（会社全額負担）', '厚生年金の標準報酬月額に乗じる。会社コスト計算にのみ影響し、個人の手取りには影響しません。下記の「子ども・子育て支援金率」とは別の制度です');
wsIn.getCell(`C${RATE_ROWS.kodomo}`).value = 0.0036; wsIn.getCell(`C${RATE_ROWS.kodomo}`).numFmt = PCT; wsIn.getCell(`C${RATE_ROWS.kodomo}`).fill = yellow;
label(wsIn, RATE_ROWS.shienkin, '子ども・子育て支援金率（労使折半）', '全国一律。令和8年4月分から新設。健康保険の標準報酬月額をベースに、労使で折半して負担します（個人負担・会社負担の両方に影響）');
wsIn.getCell(`C${RATE_ROWS.shienkin}`).value = 0.0023; wsIn.getCell(`C${RATE_ROWS.shienkin}`).numFmt = PCT; wsIn.getCell(`C${RATE_ROWS.shienkin}`).fill = yellow;
label(wsIn, RATE_ROWS.corplow, '法人実効税率（年800万円以下の部分）', '資本金1億円以下・標準税率ベースの概算実効税率。超過税率を課す自治体や事業年度により異なるため必ず検証してください');
wsIn.getCell(`C${RATE_ROWS.corplow}`).value = 0.2137; wsIn.getCell(`C${RATE_ROWS.corplow}`).numFmt = PCT; wsIn.getCell(`C${RATE_ROWS.corplow}`).fill = yellow;
label(wsIn, RATE_ROWS.corphigh, '法人実効税率（年800万円超の部分）', '同上（軽減税率の適用外部分）');
wsIn.getCell(`C${RATE_ROWS.corphigh}`).value = 0.3358; wsIn.getCell(`C${RATE_ROWS.corphigh}`).numFmt = PCT; wsIn.getCell(`C${RATE_ROWS.corphigh}`).fill = yellow;
label(wsIn, RATE_ROWS.corpkintou, '法人住民税 均等割（年額）', '資本金等1,000万円以下・従業者50人以下の標準額（道府県2万円＋市町村5万円）。赤字でも課されます。資本金等1,000万円超1億円以下なら18万円など、規模と所在自治体により異なります');
wsIn.getCell(`C${RATE_ROWS.corpkintou}`).value = 70000; wsIn.getCell(`C${RATE_ROWS.corpkintou}`).numFmt = YEN; wsIn.getCell(`C${RATE_ROWS.corpkintou}`).fill = yellow;

header(wsIn, SIMRANGE_HEADER_ROW, '【シミュレーション範囲（各役員のシミュレーションシートで共通）】');
label(wsIn, SIMRANGE_ROWS.start, '役員報酬（年額）開始', '');
wsIn.getCell(`C${SIMRANGE_ROWS.start}`).value = 0; wsIn.getCell(`C${SIMRANGE_ROWS.start}`).numFmt = YEN; wsIn.getCell(`C${SIMRANGE_ROWS.start}`).fill = yellow;
label(wsIn, SIMRANGE_ROWS.end, '役員報酬（年額）終了', '');
wsIn.getCell(`C${SIMRANGE_ROWS.end}`).value = 36000000; wsIn.getCell(`C${SIMRANGE_ROWS.end}`).numFmt = YEN; wsIn.getCell(`C${SIMRANGE_ROWS.end}`).fill = yellow;
label(wsIn, SIMRANGE_ROWS.step, '刻み幅', '');
wsIn.getCell(`C${SIMRANGE_ROWS.step}`).value = 500000; wsIn.getCell(`C${SIMRANGE_ROWS.step}`).numFmt = YEN; wsIn.getCell(`C${SIMRANGE_ROWS.step}`).fill = yellow;

wsIn.getCell(`B${CAVEAT_ROW}`).value = '※本シートは一般的なケースを想定した概算シミュレーションです。実際の申告・届出にあたっては必ず個別事情を確認し、税額・社会保険料は正式な計算方法で検算してください。住民税の調整控除、配偶者特別控除、扶養控除の細分（同居老親等以外の老人扶養等）、社会保険の定時決定・随時改定のタイミングのズレ、住民税の非課税限度額判定（所得が極めて低い場合の均等割非課税等）は簡略化のため考慮していません。個人住民税の均等割に併せて課される森林環境税（年1,000円）と、法人住民税の均等割（赤字でも発生。標準額7万円。入力シートで変更可）は織り込んでいます。防衛特別法人税（令和8年4月1日以後開始事業年度、基準法人税額の4％・年500万円控除）は反映していません。役員報酬0円付近の結果は参考値としての精度が低い点にご留意ください。役員が複数いる場合、各「シミュレーション」シートは「他の役員の報酬は結果サマリーシートで採用した金額のまま固定」した前提で、その役員自身の報酬額だけを動かした試算です。';
wsIn.getCell(`B${CAVEAT_ROW}`).font = { size: 9, color: { argb: 'FFC00000' } };
wsIn.getCell(`B${CAVEAT_ROW}`).alignment = { wrapText: true };
wsIn.mergeCells(`B${CAVEAT_ROW}:D${CAVEAT_ROW + 4}`);

for (let r = 4; r <= SIMRANGE_ROWS.step; r++) { wsIn.getRow(r).height = 18; }

// ============================================================
// シート2: 早見表（料率・控除額テーブル）
// ============================================================
const wsRef = wb.addWorksheet('早見表', { properties: { tabColor: { argb: 'FF808080' } } });

wsRef.getCell('A1').value = '健康保険 標準報酬月額等級表（協会けんぽ・50等級）';
wsRef.getCell('A1').font = { bold: true };
wsRef.getCell('A2').value = '等級'; wsRef.getCell('B2').value = '報酬月額(以上)'; wsRef.getCell('C2').value = '報酬月額(未満)'; wsRef.getCell('D2').value = '標準報酬月額';
['A2','B2','C2','D2'].forEach(c => wsRef.getCell(c).font = { bold: true });

const kenpoLower = [0,63000,73000,83000,93000,101000,107000,114000,122000,130000,138000,146000,155000,165000,175000,185000,195000,210000,230000,250000,270000,290000,310000,330000,350000,370000,395000,425000,455000,485000,515000,545000,575000,605000,635000,665000,695000,730000,770000,810000,855000,905000,955000,1005000,1055000,1115000,1175000,1235000,1295000,1355000];
const kenpoUpper = kenpoLower.slice(1).concat([9999999999]);
const kenpoStd   = [58000,68000,78000,88000,98000,104000,110000,118000,126000,134000,142000,150000,160000,170000,180000,190000,200000,220000,240000,260000,280000,300000,320000,340000,360000,380000,410000,440000,470000,500000,530000,560000,590000,620000,650000,680000,710000,750000,790000,830000,880000,930000,980000,1030000,1090000,1150000,1210000,1270000,1330000,1390000];

for (let i = 0; i < 50; i++) {
  const r = 3 + i;
  wsRef.getCell(`A${r}`).value = i + 1;
  wsRef.getCell(`B${r}`).value = kenpoLower[i]; wsRef.getCell(`B${r}`).numFmt = YEN0;
  wsRef.getCell(`C${r}`).value = kenpoUpper[i] > 100000000 ? '－' : kenpoUpper[i]; wsRef.getCell(`C${r}`).numFmt = YEN0;
  wsRef.getCell(`D${r}`).value = kenpoStd[i]; wsRef.getCell(`D${r}`).numFmt = YEN0;
}

wsRef.getCell('F1').value = '厚生年金保険 標準報酬月額等級表（32等級）';
wsRef.getCell('F1').font = { bold: true };
wsRef.getCell('F2').value = '等級'; wsRef.getCell('G2').value = '報酬月額(以上)'; wsRef.getCell('H2').value = '報酬月額(未満)'; wsRef.getCell('I2').value = '標準報酬月額';
['F2','G2','H2','I2'].forEach(c => wsRef.getCell(c).font = { bold: true });

const kouseiLower = [0,93000,101000,107000,114000,122000,130000,138000,146000,155000,165000,175000,185000,195000,210000,230000,250000,270000,290000,310000,330000,350000,370000,395000,425000,455000,485000,515000,545000,575000,605000,635000];
const kouseiUpper = kouseiLower.slice(1).concat([9999999999]);
const kouseiStd   = [88000,98000,104000,110000,118000,126000,134000,142000,150000,160000,170000,180000,190000,200000,220000,240000,260000,280000,300000,320000,340000,360000,380000,410000,440000,470000,500000,530000,560000,590000,620000,650000];

for (let i = 0; i < 32; i++) {
  const r = 3 + i;
  wsRef.getCell(`F${r}`).value = i + 1;
  wsRef.getCell(`G${r}`).value = kouseiLower[i]; wsRef.getCell(`G${r}`).numFmt = YEN0;
  wsRef.getCell(`H${r}`).value = kouseiUpper[i] > 100000000 ? '－' : kouseiUpper[i]; wsRef.getCell(`H${r}`).numFmt = YEN0;
  wsRef.getCell(`I${r}`).value = kouseiStd[i]; wsRef.getCell(`I${r}`).numFmt = YEN0;
}

// 所得税速算表
wsRef.getCell('K1').value = '所得税 速算表（課税所得ベース）';
wsRef.getCell('K1').font = { bold: true };
wsRef.getCell('K2').value = '課税所得(以上)'; wsRef.getCell('L2').value = '税率'; wsRef.getCell('M2').value = '控除額';
['K2','L2','M2'].forEach(c => wsRef.getCell(c).font = { bold: true });
const taxBrackets = [
  [0, 0.05, 0],
  [1950000, 0.10, 97500],
  [3300000, 0.20, 427500],
  [6950000, 0.23, 636000],
  [9000000, 0.33, 1536000],
  [18000000, 0.40, 2796000],
  [40000000, 0.45, 4796000],
];
taxBrackets.forEach((b, i) => {
  const r = 3 + i;
  wsRef.getCell(`K${r}`).value = b[0]; wsRef.getCell(`K${r}`).numFmt = YEN0;
  wsRef.getCell(`L${r}`).value = b[1]; wsRef.getCell(`L${r}`).numFmt = PCT;
  wsRef.getCell(`M${r}`).value = b[2]; wsRef.getCell(`M${r}`).numFmt = YEN0;
});

// 基礎控除（所得税・令和8年度税制改正後、令和8・9年分）
// 出典: 国税庁「源泉所得税の改正のあらまし 令和8年4月」
wsRef.getCell('O1').value = '基礎控除（所得税・令和8・9年分）';
wsRef.getCell('O1').font = { bold: true };
wsRef.getCell('O2').value = '合計所得(以上)'; wsRef.getCell('P2').value = '控除額';
const kisoIncomeTax = [[0,1040000],[4890000,670000],[6550000,620000],[23500000,480000],[24000000,320000],[24500000,160000],[25000000,0]];
kisoIncomeTax.forEach((b,i)=>{ const r=3+i; wsRef.getCell(`O${r}`).value=b[0]; wsRef.getCell(`O${r}`).numFmt=YEN0; wsRef.getCell(`P${r}`).value=b[1]; wsRef.getCell(`P${r}`).numFmt=YEN0; });

// 固定控除額（住民税等）と扶養控除
wsRef.getCell('R1').value = '固定控除額等';
wsRef.getCell('R1').font = { bold: true };
const fixedVals = [
  ['基礎控除（住民税・据置き。合計所得2,400万円以下）', 430000],
  ['住民税 均等割＋森林環境税（標準額。均等割4,000円＋森林環境税1,000円）', 5000],
  ['住民税 所得割率', 0.10],
  ['扶養控除（所得税）一般', 380000],
  ['扶養控除（所得税）特定', 630000],
  ['扶養控除（所得税）老人同居', 580000],
  ['扶養控除（住民税）一般', 330000],
  ['扶養控除（住民税）特定', 450000],
  ['扶養控除（住民税）老人同居', 450000],
];
fixedVals.forEach((v,i)=>{
  const r = 2+i;
  wsRef.getCell(`R${r}`).value = v[0];
  wsRef.getCell(`S${r}`).value = v[1];
  wsRef.getCell(`S${r}`).numFmt = (v[1] < 1) ? PCT : YEN0;
});

// 配偶者控除・配偶者特別控除額（令和8・9年分、本人の合計所得3区分×所得税/住民税）
// 出典: 国税庁パンフレット「配偶者特別控除額の表」。住民税は所得税額を住民税側の上限(33万/22万/11万)でキャップした額
wsRef.getCell('U1').value = '配偶者控除・配偶者特別控除額（令和8・9年分）';
wsRef.getCell('U1').font = { bold: true };
wsRef.getCell('U2').value = '配偶者合計所得(以上)';
wsRef.getCell('V2').value = '所得税(900万円以下)'; wsRef.getCell('W2').value = '所得税(900-950万円)'; wsRef.getCell('X2').value = '所得税(950-1000万円)';
wsRef.getCell('Y2').value = '住民税(900万円以下)'; wsRef.getCell('Z2').value = '住民税(900-950万円)'; wsRef.getCell('AA2').value = '住民税(950-1000万円)';
['U2','V2','W2','X2','Y2','Z2','AA2'].forEach(c => wsRef.getCell(c).font = { bold: true });
const haiguTable = [
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
haiguTable.forEach((row, i) => {
  const r = 3 + i;
  const cols = ['U','V','W','X','Y','Z','AA'];
  row.forEach((v, ci) => { wsRef.getCell(`${cols[ci]}${r}`).value = v; wsRef.getCell(`${cols[ci]}${r}`).numFmt = YEN0; });
});

// 特定親族特別控除額（令和8年度税制改正で新設。所得者本人の所得による段階はなく一律）
// 出典: 国税庁パンフレット「特定親族特別控除額の表」。住民税は所得税額を45万円上限でキャップした額
wsRef.getCell('AD1').value = '特定親族特別控除額（令和8・9年分）';
wsRef.getCell('AD1').font = { bold: true };
wsRef.getCell('AD2').value = '特定親族合計所得(以上)';
wsRef.getCell('AE2').value = '所得税'; wsRef.getCell('AF2').value = '住民税';
['AD2','AE2','AF2'].forEach(c => wsRef.getCell(c).font = { bold: true });
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
tokuteiShinzokuTable.forEach((row, i) => {
  const r = 3 + i;
  const cols = ['AD','AE','AF'];
  row.forEach((v, ci) => { wsRef.getCell(`${cols[ci]}${r}`).value = v; wsRef.getCell(`${cols[ci]}${r}`).numFmt = YEN0; });
});

wsRef.columns.forEach(c => c.width = 15);
wsRef.getColumn('A').width = 6; wsRef.getColumn('F').width = 6;
wsRef.getColumn('R').width = 42;
wsRef.getColumn('U').width = 18;
wsRef.getColumn('AD').width = 20;

// named cells for fixed values (row refs): S2 基礎控除住民税, S3 均等割, S4 所得割率
// S5-S7 扶養控除所得税(一般/特定/老人), S8-S10 扶養控除住民税(一般/特定/老人)

// ============================================================
// シート3〜7: シミュレーション1〜5（役員ごと）
// ============================================================
const headers = [
  '役員報酬(年額)', '役員報酬(月額)', '標準報酬月額(健保)', '標準報酬月額(厚年)',
  '社会保険料(個人負担)', '社会保険料(会社負担)', '給与所得控除', '給与所得',
  '所得控除合計(所得税)', '課税所得(所得税)', '所得税', '復興特別所得税', '所得税等合計',
  '住民税(所得割)', '住民税(均等割等)', '住民税合計',
  '個人負担合計', '個人手取り額', '個人手取り率',
  '会社コスト(この役員の報酬+法定福利費)', '役員報酬控除後利益(全役員ベース)', '法人税等', '内部留保(税引後)',
  'トータル手取り(内部留保+個人手取り合計)', '実効負担率(対税引前利益)'
];

const FIRST_ROW = 2;
const LAST_ROW = 301; // 300 data rows

function buildSimSheet(k) {
  const r0 = officerRows(k);
  const sumR = summaryOfficerRows(k);
  const ws = wb.addWorksheet(`シミュレーション${k + 1}`, { properties: { tabColor: { argb: 'FF548235' } } });

  ws.getRow(1).values = ['', ...headers];
  ws.getRow(1).font = { bold: true };
  ws.getRow(1).alignment = { wrapText: true, vertical: 'bottom', horizontal: 'center' };
  ws.getRow(1).height = 42;
  ws.views = [{ state: 'frozen', ySplit: 1, xSplit: 1 }];

  ws.getColumn(1).width = 5;
  for (let i = 0; i < headers.length; i++) {
    ws.getColumn(i + 2).width = 15;
  }

  const otherCostRef = `('結果サマリー'!$C$${SUM_AGG_ROWS.totalCost}-'結果サマリー'!$C$${sumR.cost})`;
  const otherTedoriRef = `('結果サマリー'!$C$${SUM_AGG_ROWS.totalTedori}-'結果サマリー'!$C$${sumR.tedori})`;

  for (let i = 0; i < (LAST_ROW - FIRST_ROW + 1); i++) {
    const r = FIRST_ROW + i;
    ws.getCell(`A${r}`).value = i + 1;

    if (i === 0) {
      ws.getCell(`B${r}`).value = { formula: `'入力'!$C$${SIMRANGE_ROWS.start}` };
    } else {
      ws.getCell(`B${r}`).value = { formula: `IF(B${r-1}="","",IF(B${r-1}+'入力'!$C$${SIMRANGE_ROWS.step}>'入力'!$C$${SIMRANGE_ROWS.end},"",B${r-1}+'入力'!$C$${SIMRANGE_ROWS.step}))` };
    }

    const IFB = (f) => `IF($B${r}="","",${f})`;

    // C: 月額報酬
    ws.getCell(`C${r}`).value = { formula: IFB(`$B${r}/12`) };
    // D: 標準報酬月額(健保) ※報酬0円の役員は社会保険の被保険者資格を取得しないため0とする
    ws.getCell(`D${r}`).value = { formula: IFB(`IF($B${r}=0,0,VLOOKUP($C${r},'早見表'!$B$3:$D$52,3,TRUE()))`) };
    // E: 標準報酬月額(厚年) ※同上
    ws.getCell(`E${r}`).value = { formula: IFB(`IF($B${r}=0,0,VLOOKUP($C${r},'早見表'!$G$3:$I$34,3,TRUE()))`) };
    // F: 社会保険料(個人負担・年額)
    ws.getCell(`F${r}`).value = { formula: IFB(`ROUND($D${r}*('入力'!$C$${RATE_ROWS.kenpo}+IF('入力'!$C$${r0.age}="40歳以上64歳以下",'入力'!$C$${RATE_ROWS.kaigo},0))/2*12,0)+ROUND($E${r}*'入力'!$C$${RATE_ROWS.kousei}/2*12,0)+ROUND($D${r}*'入力'!$C$${RATE_ROWS.shienkin}/2*12,0)`) };
    // G: 社会保険料(会社負担・年額)
    ws.getCell(`G${r}`).value = { formula: IFB(`ROUND($D${r}*('入力'!$C$${RATE_ROWS.kenpo}+IF('入力'!$C$${r0.age}="40歳以上64歳以下",'入力'!$C$${RATE_ROWS.kaigo},0))/2*12,0)+ROUND($E${r}*'入力'!$C$${RATE_ROWS.kousei}/2*12,0)+ROUND($E${r}*'入力'!$C$${RATE_ROWS.kodomo}*12,0)+ROUND($D${r}*'入力'!$C$${RATE_ROWS.shienkin}/2*12,0)`) };
    // H: 給与所得控除(令和8・9年分。最低保障額74万円。69.1万円〜220万円の端数調整特例は簡略化し220万円まで一律74万円としています)
    ws.getCell(`H${r}`).value = { formula: IFB(`IF($B${r}<=2200000,740000,IF($B${r}<=3600000,$B${r}*0.3+80000,IF($B${r}<=6600000,$B${r}*0.2+440000,IF($B${r}<=8500000,$B${r}*0.1+1100000,1950000))))`) };
    // I: 給与所得
    ws.getCell(`I${r}`).value = { formula: IFB(`MAX($B${r}-$H${r},0)`) };
    // J: 所得控除合計(所得税)
    const kisoTax = `VLOOKUP($I${r},'早見表'!$O$3:$P$9,2,TRUE())`;
    const filerTierCol = `IF($I${r}<=9000000,2,IF($I${r}<=9500000,3,4))`;
    const haiguTax = `IF(AND('入力'!$C$${r0.spouse}="有",$I${r}<=10000000),VLOOKUP('入力'!$E$${r0.spouseIncome},'早見表'!$U$3:$AA$12,${filerTierCol},TRUE()),0)`;
    const tokuteiTax = `IF('入力'!$C$${r0.tokuteiCount}<=0,0,IF('入力'!$E$${r0.tokuteiIncome}<620001,630000,VLOOKUP('入力'!$E$${r0.tokuteiIncome},'早見表'!$AD$3:$AF$12,2,TRUE()))*'入力'!$C$${r0.tokuteiCount})`;
    const regularSpecCount = `MAX('入力'!$C$${r0.specDep}-'入力'!$C$${r0.tokuteiCount},0)`;
    const fuyouTax = `'入力'!$C$${r0.genDep}*'早見表'!$S$5+${regularSpecCount}*'早見表'!$S$6+${tokuteiTax}+'入力'!$C$${r0.elderDep}*'早見表'!$S$7`;
    ws.getCell(`J${r}`).value = { formula: IFB(`$F${r}+${kisoTax}+${haiguTax}+${fuyouTax}`) };
    // K: 課税所得(所得税・1000円未満切捨)
    ws.getCell(`K${r}`).value = { formula: IFB(`ROUNDDOWN(MAX($I${r}-$J${r},0),-3)`) };
    // L: 所得税
    ws.getCell(`L${r}`).value = { formula: IFB(`ROUNDDOWN(MAX($K${r}*VLOOKUP($K${r},'早見表'!$K$3:$M$9,2,TRUE())-VLOOKUP($K${r},'早見表'!$K$3:$M$9,3,TRUE()),0),0)`) };
    // M: 復興特別所得税
    ws.getCell(`M${r}`).value = { formula: IFB(`ROUNDDOWN($L${r}*0.021,0)`) };
    // N: 所得税等合計
    ws.getCell(`N${r}`).value = { formula: IFB(`$L${r}+$M${r}`) };
    // O: 住民税(所得割)
    const juminTierCol = `IF($I${r}<=9000000,5,IF($I${r}<=9500000,6,7))`;
    const haiguJumin = `IF(AND('入力'!$C$${r0.spouse}="有",$I${r}<=10000000),VLOOKUP('入力'!$E$${r0.spouseIncome},'早見表'!$U$3:$AA$12,${juminTierCol},TRUE()),0)`;
    const tokuteiJumin = `IF('入力'!$C$${r0.tokuteiCount}<=0,0,IF('入力'!$E$${r0.tokuteiIncome}<620001,450000,VLOOKUP('入力'!$E$${r0.tokuteiIncome},'早見表'!$AD$3:$AF$12,3,TRUE()))*'入力'!$C$${r0.tokuteiCount})`;
    const fuyouJumin = `'入力'!$C$${r0.genDep}*'早見表'!$S$8+${regularSpecCount}*'早見表'!$S$9+${tokuteiJumin}+'入力'!$C$${r0.elderDep}*'早見表'!$S$10`;
    ws.getCell(`O${r}`).value = { formula: IFB(`ROUNDDOWN(MAX($I${r}-($F${r}+'早見表'!$S$2+${haiguJumin}+${fuyouJumin}),0),-3)*'早見表'!$S$4`) };
    // P: 住民税(均等割等)
    ws.getCell(`P${r}`).value = { formula: IFB(`'早見表'!$S$3`) };
    // Q: 住民税合計
    ws.getCell(`Q${r}`).value = { formula: IFB(`$O${r}+$P${r}`) };
    // R: 個人負担合計
    ws.getCell(`R${r}`).value = { formula: IFB(`$F${r}+$N${r}+$Q${r}`) };
    // S: 個人手取り額
    ws.getCell(`S${r}`).value = { formula: IFB(`$B${r}-$R${r}`) };
    // T: 個人手取り率
    ws.getCell(`T${r}`).value = { formula: IFB(`IF($B${r}=0,"",$S${r}/$B${r})`) };
    // U: 会社コスト(この役員のみ)
    ws.getCell(`U${r}`).value = { formula: IFB(`$B${r}+$G${r}`) };
    // V: 役員報酬控除後利益(全役員ベース。他の役員のコストは結果サマリーで採用中の金額を使用)
    ws.getCell(`V${r}`).value = { formula: IFB(`'入力'!$C$${COMPANY_PROFIT_ROW}-$U${r}-${otherCostRef}`) };
    // W: 法人税等
    ws.getCell(`W${r}`).value = { formula: IFB(`IF($V${r}<=0,0,IF($V${r}<=8000000,$V${r}*'入力'!$C$${RATE_ROWS.corplow},8000000*'入力'!$C$${RATE_ROWS.corplow}+($V${r}-8000000)*'入力'!$C$${RATE_ROWS.corphigh}))+'入力'!$C$${RATE_ROWS.corpkintou}`) };
    // X: 内部留保
    ws.getCell(`X${r}`).value = { formula: IFB(`$V${r}-$W${r}`) };
    // Y: トータル手取り(全役員ベース。他の役員の手取りは結果サマリーで採用中の金額を使用)
    ws.getCell(`Y${r}`).value = { formula: IFB(`$X${r}+$S${r}+${otherTedoriRef}`) };
    // Z: 実効負担率
    ws.getCell(`Z${r}`).value = { formula: IFB(`IF('入力'!$C$${COMPANY_PROFIT_ROW}=0,"",1-$Y${r}/'入力'!$C$${COMPANY_PROFIT_ROW})`) };

    // formats
    ws.getCell(`B${r}`).numFmt = YEN0;
    ['C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','U','V','W','X','Y'].forEach(c => ws.getCell(`${c}${r}`).numFmt = YEN0);
    ws.getCell(`T${r}`).numFmt = PCT;
    ws.getCell(`Z${r}`).numFmt = PCT;
  }

  for (let i = 0; i < headers.length; i++) {
    const col = String.fromCharCode('B'.charCodeAt(0) + i);
    ws.getCell(`${col}1`).value = headers[i];
  }
  ws.getCell('A1').value = '#';

  ws.addConditionalFormatting({
    ref: `Y${FIRST_ROW}:Y${LAST_ROW}`,
    rules: [{
      type: 'expression',
      formulae: [`Y${FIRST_ROW}=MAX($Y$${FIRST_ROW}:$Y$${LAST_ROW})`],
      style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFC6E0B4' } }, font: { bold: true } },
    }],
  });
  ws.addConditionalFormatting({
    ref: `Z${FIRST_ROW}:Z${LAST_ROW}`,
    rules: [{
      type: 'expression',
      formulae: [`AND(Z${FIRST_ROW}<>"",Z${FIRST_ROW}=MIN(IF($Z$${FIRST_ROW}:$Z$${LAST_ROW}<>"",$Z$${FIRST_ROW}:$Z$${LAST_ROW})))`],
      style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFC6E0B4' } }, font: { bold: true } },
    }],
  });

  return ws;
}

for (let k = 0; k < N_OFFICERS; k++) buildSimSheet(k);

// ============================================================
// シート8: 結果サマリー(会社全体の集計)
// ============================================================
const wsSum = wb.addWorksheet('結果サマリー', { properties: { tabColor: { argb: 'FFC00000' } } });
wsSum.columns = [{ width: 4 }, { width: 34 }, { width: 18 }, { width: 60 }];
wsSum.getCell('B1').value = '結果サマリー';
wsSum.getCell('B1').font = { bold: true, size: 16 };
wsSum.getCell('B2').value = '各役員の「採用報酬額」に、比較したい金額を入力してください（入力シートの刻み幅の倍数に丸めて各シミュレーションシートを検索します）。「入力」シートの役員の人数を超える役員欄は自動的に0円扱いとなり集計から除外されます。';
wsSum.getCell('B2').font = { italic: true, size: 9 };
wsSum.getCell('B2').alignment = { wrapText: true };
wsSum.mergeCells('B2:D2');

wsSum.getCell(`B${SUM_OFFICER_HEADER_ROW}`).value = '① 役員別 現在の設定';
wsSum.getCell(`B${SUM_OFFICER_HEADER_ROW}`).font = { bold: true };

for (let k = 0; k < N_OFFICERS; k++) {
  const s = summaryOfficerRows(k);
  wsSum.getCell(`B${s.input}`).value = `役員${k + 1} 採用報酬額(年額)`;
  wsSum.getCell(`C${s.input}`).value = k === 0 ? 5000000 : 0;
  wsSum.getCell(`C${s.input}`).numFmt = YEN; wsSum.getCell(`C${s.input}`).fill = yellow;
  wsSum.getCell(`D${s.input}`).value = { formula: `IF(${k + 1}<='入力'!$C$${OFFICER_COUNT_ROW},MROUND(C${s.input},'入力'!$C$${SIMRANGE_ROWS.step}),0)` };
  wsSum.getCell(`D${s.input}`).numFmt = YEN;
  wsSum.getCell(`E${s.input}`).value = '(集計に使う丸め後・人数ゲート後の金額)';
  wsSum.getCell(`E${s.input}`).font = { italic: true, size: 9, color: { argb: 'FF808080' } };

  wsSum.getCell(`B${s.tedori}`).value = '個人手取り';
  wsSum.getCell(`C${s.tedori}`).value = { formula: `IF(${k + 1}>'入力'!$C$${OFFICER_COUNT_ROW},0,IFERROR(INDEX('シミュレーション${k + 1}'!$S$2:$S$301,MATCH(D${s.input},'シミュレーション${k + 1}'!$B$2:$B$301,0)),0))` };
  wsSum.getCell(`C${s.tedori}`).numFmt = YEN;

  wsSum.getCell(`B${s.cost}`).value = '会社コスト(この役員分)';
  wsSum.getCell(`C${s.cost}`).value = { formula: `IF(${k + 1}>'入力'!$C$${OFFICER_COUNT_ROW},0,IFERROR(INDEX('シミュレーション${k + 1}'!$U$2:$U$301,MATCH(D${s.input},'シミュレーション${k + 1}'!$B$2:$B$301,0)),0))` };
  wsSum.getCell(`C${s.cost}`).numFmt = YEN;
}

wsSum.getCell(`B${SUM_AGG_HEADER_ROW}`).value = '② 会社全体(現在の設定に基づく集計)';
wsSum.getCell(`B${SUM_AGG_HEADER_ROW}`).font = { bold: true };

const tedoriCells = Array.from({ length: N_OFFICERS }, (_, k) => `C${summaryOfficerRows(k).tedori}`).join(',');
const costCells = Array.from({ length: N_OFFICERS }, (_, k) => `C${summaryOfficerRows(k).cost}`).join(',');
const xCells = Array.from({ length: N_OFFICERS }, (_, k) => `D${summaryOfficerRows(k).input}`).join(',');

wsSum.getCell(`B${SUM_AGG_ROWS.totalX}`).value = '役員報酬 合計(年額)';
wsSum.getCell(`C${SUM_AGG_ROWS.totalX}`).value = { formula: `SUM(${xCells})` };
wsSum.getCell(`C${SUM_AGG_ROWS.totalX}`).numFmt = YEN;

wsSum.getCell(`B${SUM_AGG_ROWS.totalCost}`).value = '会社コスト 合計(全役員)';
wsSum.getCell(`C${SUM_AGG_ROWS.totalCost}`).value = { formula: `SUM(${costCells})` };
wsSum.getCell(`C${SUM_AGG_ROWS.totalCost}`).numFmt = YEN;

wsSum.getCell(`B${SUM_AGG_ROWS.rieki}`).value = '役員報酬控除後利益';
wsSum.getCell(`C${SUM_AGG_ROWS.rieki}`).value = { formula: `'入力'!$C$${COMPANY_PROFIT_ROW}-C${SUM_AGG_ROWS.totalCost}` };
wsSum.getCell(`C${SUM_AGG_ROWS.rieki}`).numFmt = YEN;

wsSum.getCell(`B${SUM_AGG_ROWS.houjinzei}`).value = '法人税等';
wsSum.getCell(`C${SUM_AGG_ROWS.houjinzei}`).value = { formula: `IF(C${SUM_AGG_ROWS.rieki}<=0,0,IF(C${SUM_AGG_ROWS.rieki}<=8000000,C${SUM_AGG_ROWS.rieki}*'入力'!$C$${RATE_ROWS.corplow},8000000*'入力'!$C$${RATE_ROWS.corplow}+(C${SUM_AGG_ROWS.rieki}-8000000)*'入力'!$C$${RATE_ROWS.corphigh}))+'入力'!$C$${RATE_ROWS.corpkintou}` };
wsSum.getCell(`C${SUM_AGG_ROWS.houjinzei}`).numFmt = YEN;

wsSum.getCell(`B${SUM_AGG_ROWS.naibu}`).value = '内部留保(税引後)';
wsSum.getCell(`C${SUM_AGG_ROWS.naibu}`).value = { formula: `C${SUM_AGG_ROWS.rieki}-C${SUM_AGG_ROWS.houjinzei}` };
wsSum.getCell(`C${SUM_AGG_ROWS.naibu}`).numFmt = YEN;

wsSum.getCell(`B${SUM_AGG_ROWS.totalTedori}`).value = '個人手取り 合計(全役員)';
wsSum.getCell(`C${SUM_AGG_ROWS.totalTedori}`).value = { formula: `SUM(${tedoriCells})` };
wsSum.getCell(`C${SUM_AGG_ROWS.totalTedori}`).numFmt = YEN;

wsSum.getCell(`B${SUM_AGG_ROWS.grandTotal}`).value = 'グループ全体のトータル手取り(内部留保+個人手取り合計)';
wsSum.getCell(`C${SUM_AGG_ROWS.grandTotal}`).value = { formula: `C${SUM_AGG_ROWS.naibu}+C${SUM_AGG_ROWS.totalTedori}` };
wsSum.getCell(`C${SUM_AGG_ROWS.grandTotal}`).numFmt = YEN;

wsSum.getCell(`B${SUM_AGG_ROWS.rate}`).value = '実効負担率(税金・社会保険料 ÷ 税引前利益)';
wsSum.getCell(`C${SUM_AGG_ROWS.rate}`).value = { formula: `IF('入力'!$C$${COMPANY_PROFIT_ROW}=0,"",1-C${SUM_AGG_ROWS.grandTotal}/'入力'!$C$${COMPANY_PROFIT_ROW})` };
wsSum.getCell(`C${SUM_AGG_ROWS.rate}`).numFmt = PCT;

const noteRow = SUM_AGG_ROWS.rate + 2;
wsSum.getCell(`B${noteRow}`).value = '※「グループ全体のトータル手取り」は内部留保に対する将来の配当課税・退職金課税等を考慮しない単純合算指標です。あくまで比較の目安としてご利用ください。各役員の「シミュレーション」シートに表示される最適額は、他の役員の採用報酬額をここでの設定のまま固定した場合の、その役員1人分の最適化結果です。全役員を同時に最適化する場合は、役員ごとに①の採用報酬額を更新しながら繰り返し調整してください。';
wsSum.getCell(`B${noteRow}`).font = { size: 9, italic: true, color: { argb: 'FFC00000' } };
wsSum.getCell(`B${noteRow}`).alignment = { wrapText: true };
wsSum.mergeCells(`B${noteRow}:D${noteRow + 3}`);

for (let r = 3; r <= noteRow + 3; r++) wsSum.getRow(r).height = 16;

// order sheets
wb.views = [{ firstSheet: 0, activeTab: 0 }];

wb.xlsx.writeFile(OUT_PATH).then(() => {
  console.log('written to', OUT_PATH);
}).catch(err => {
  console.error(err);
  process.exit(1);
});
