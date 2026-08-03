// 配偶者控除・配偶者特別控除(令和8・9年分)の段階的な減額がExcel上で正しく機能しているかの専用検証
const ExcelJS = require('exceljs');
const { HyperFormula } = require('hyperformula');
const path = require('path');

const XLSX_PATH = path.join(__dirname, '役員報酬シミュレーション.xlsx');
const OFFICER1_SPOUSE_INCOME_ROW = 12; // officerRows(0).spouseIncome
const OFFICER1_X_ROW = 13; // シミュレーション1 row matching X=6,000,000 (from earlier test: row=13)

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

  // シミュレーション1のX=6,000,000行(row13)を使い、officer1の配偶者収入を段階的に変えてJ列(所得控除合計)経由でhaigu額を逆算する。
  // kiso(1,040,000, kyuyoShotoku=4,360,000<=489万) + haigu + fuyouTax(0) が J に入っているので、haigu = J - 1,040,000
  const cases = [
    [0, 380000],
    [1360000, 380000],
    [1690000, 380000],
    [1752000, 310000],
    [1840000, 260000],
    [1900000, 160000],
    [2070000, 30000],
    [2200000, 0],
  ];
  for (const [spouseIncome, expectedHaigu] of cases) {
    setVal('入力', `C${OFFICER1_SPOUSE_INCOME_ROW}`, spouseIncome);
    const j = getVal('シミュレーション1', `J${OFFICER1_X_ROW}`);
    const f = getVal('シミュレーション1', `F${OFFICER1_X_ROW}`);
    const haigu = j - f - 1040000;
    check(`配偶者収入=${spouseIncome}`, haigu, expectedHaigu);
  }

  console.log('\n' + (allPass ? '★★★ 配偶者控除/配偶者特別控除(Excel) — ALL PASS ★★★' : '!!! MISMATCH FOUND !!!'));
}

main().catch(e => { console.error(e); process.exit(1); });
