// 特定親族特別控除(令和8年度税制改正で新設)がExcel上で正しく機能しているかの専用検証
const ExcelJS = require('exceljs');
const { HyperFormula } = require('hyperformula');
const path = require('path');

const XLSX_PATH = path.join(__dirname, '役員報酬シミュレーション.xlsx');
const OFFICER1_SPEC_DEP_ROW = 14;   // officerRows(0).specDep = 9+5
const OFFICER1_TOKUTEI_COUNT_ROW = 15; // officerRows(0).tokuteiCount = 9+6
const OFFICER1_TOKUTEI_INCOME_ROW = 16; // officerRows(0).tokuteiIncome = 9+7
const OFFICER1_SPOUSE_ROW = 11; // officerRows(0).spouse = 9+2
const OFFICER1_X_ROW = 13; // シミュレーション1 row matching X=6,000,000

async function main() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(XLSX_PATH);
  const sheetNames = ['入力', '早見表', 'シミュレーション1', 'シミュレーション2', 'シミュレーション3', 'シミュレーション4', 'シミュレーション5', '結果サマリー'];
  const sheetsData = {};
  for (const name of sheetNames) {
    const ws = wb.getWorksheet(name);
    const data = [];
    for (let r = 1; r <= ws.rowCount; r++) {
      const row = [];
      const wsRow = ws.getRow(r);
      for (let c = 1; c <= ws.columnCount; c++) {
        const cell = wsRow.getCell(c);
        if (cell.formula) row.push('=' + cell.formula);
        else if (cell.value && typeof cell.value === 'object' && cell.value.formula) row.push('=' + cell.value.formula);
        else if (cell.value !== null && cell.value !== undefined) row.push(cell.value);
        else row.push(null);
      }
      data.push(row);
    }
    sheetsData[name] = data;
  }
  const hf = HyperFormula.buildFromSheets(sheetsData, { licenseKey: 'gpl-v3' });
  const sid = {};
  for (const name of sheetNames) sid[name] = hf.getSheetId(name);
  const colIdx = (letter) => letter.split('').reduce((a, c) => a * 26 + (c.charCodeAt(0) - 64), 0) - 1;
  const setVal = (sheetName, addr, value) => {
    const m = addr.match(/^([A-Z]+)(\d+)$/);
    hf.setCellContents({ sheet: sid[sheetName], row: parseInt(m[2], 10) - 1, col: colIdx(m[1]) }, [[value]]);
  };
  const getVal = (sheetName, addr) => {
    const m = addr.match(/^([A-Z]+)(\d+)$/);
    return hf.getCellValue({ sheet: sid[sheetName], row: parseInt(m[2], 10) - 1, col: colIdx(m[1]) });
  };

  let allPass = true;
  function check(label, actual, expected, tol) {
    tol = tol || 1;
    const ok = Math.abs(actual - expected) <= tol;
    if (!ok) allPass = false;
    console.log(`${ok ? 'OK ' : 'NG!'} ${label}: excel=${actual} expected=${expected}`);
  }

  setVal('入力', `C${OFFICER1_SPOUSE_ROW}`, '無'); // 配偶者控除を外して特定親族控除だけを見る
  setVal('入力', `C${OFFICER1_SPEC_DEP_ROW}`, 1);

  // J = F(社会保険料個人負担) + kiso(1,040,000) + haigu(0) + fuyouTax
  const cases = [
    [0, 630000],       // 収入0 -> 通常の特定扶養控除(満額)
    [1400000, 630000], // 合計所得66万 -> 62-85万バンドでも実は63万(=満額と同額)
    [1600000, 610000], // 合計所得86万 -> 85-90万 -> 61万
    [1650000, 510000], // 合計所得91万 -> 90-95万 -> 51万
    [1700000, 410000], // 合計所得96万 -> 95-100万 -> 41万
    [1900000, 60000],  // 合計所得116万 -> 115-120万 -> 6万
    [2000000, 0],      // 合計所得126万 -> 123万超 -> 対象外
  ];
  for (const [income, expected] of cases) {
    setVal('入力', `C${OFFICER1_TOKUTEI_COUNT_ROW}`, 1);
    setVal('入力', `C${OFFICER1_TOKUTEI_INCOME_ROW}`, income);
    const j = getVal('シミュレーション1', `J${OFFICER1_X_ROW}`);
    const f = getVal('シミュレーション1', `F${OFFICER1_X_ROW}`);
    const fuyouTax = j - f - 1040000;
    check(`特定親族収入=${income}`, fuyouTax, expected);
  }

  // 2人のうち1人だけ収入がある場合: 残り1人は満額(63万)+収入がある1人(85-90万->61万) = 124万
  setVal('入力', `C${OFFICER1_SPEC_DEP_ROW}`, 2);
  setVal('入力', `C${OFFICER1_TOKUTEI_COUNT_ROW}`, 1);
  setVal('入力', `C${OFFICER1_TOKUTEI_INCOME_ROW}`, 1600000);
  {
    const j = getVal('シミュレーション1', `J${OFFICER1_X_ROW}`);
    const f = getVal('シミュレーション1', `F${OFFICER1_X_ROW}`);
    const fuyouTax = j - f - 1040000;
    check('2人中1人だけ収入あり', fuyouTax, 630000 + 610000);
  }

  console.log('\n' + (allPass ? '★★★ 特定親族特別控除(Excel) — ALL PASS ★★★' : '!!! MISMATCH FOUND !!!'));
}

main().catch(e => { console.error(e); process.exit(1); });
