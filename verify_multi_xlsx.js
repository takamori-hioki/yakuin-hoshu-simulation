const ExcelJS = require('exceljs');
const { HyperFormula } = require('hyperformula');
const { simulate } = require('./calc_reference');
const path = require('path');

const XLSX_PATH = path.join(__dirname, '役員報酬シミュレーション.xlsx');

// mirror gen_xlsx.js row layout constants
const COMPANY_PROFIT_ROW = 5;
const OFFICER_COUNT_ROW = 7;
function officerRows(k) {
  const header = 9 + k * 10;
  return { header, age: header + 1, spouse: header + 2, spouseIncome: header + 3, genDep: header + 4, specDep: header + 5,
    tokuteiCount: header + 6, tokuteiIncome: header + 7, elderDep: header + 8 };
}
function summaryOfficerRows(k) {
  const base = 5 + k * 3;
  return { input: base, tedori: base + 1, cost: base + 2 };
}
const SUM_AGG_ROWS = { totalX: 22, totalCost: 23, rieki: 24, houjinzei: 25, naibu: 26, totalTedori: 27, grandTotal: 28, rate: 29 };

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
      const maxCol = ws.columnCount;
      for (let c = 1; c <= maxCol; c++) {
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
  const cellVal = (sheetName, addr) => {
    const m = addr.match(/^([A-Z]+)(\d+)$/);
    return hf.getCellValue({ sheet: sid[sheetName], row: parseInt(m[2], 10) - 1, col: colIdx(m[1]) });
  };
  const setVal = (sheetName, addr, value) => {
    const m = addr.match(/^([A-Z]+)(\d+)$/);
    hf.setCellContents({ sheet: sid[sheetName], row: parseInt(m[2], 10) - 1, col: colIdx(m[1]) }, [[value]]);
  };

  function findRowForX(sheetName, target) {
    const sheetId = sid[sheetName];
    const dims = hf.getSheetDimensions(sheetId);
    for (let r = 1; r < dims.height; r++) {
      const v = hf.getCellValue({ sheet: sheetId, row: r, col: 1 });
      if (v === target) return r;
    }
    return -1;
  }

  let allPass = true;
  function check(label, cond, detail) {
    if (!cond) { allPass = false; console.log(`NG [${label}] ${detail || ''}`); }
    else console.log(`OK  [${label}] ${detail || ''}`);
  }

  // ================================================================
  // Test A: baseline single-officer (default state: 人数=1, officer1 defaults)
  // ================================================================
  console.log('--- Test A: single officer baseline (シミュレーション1) ---');
  const paramsBase = { profitBefore: 20000000, ageOver40: false, hasSpouse: true, spouseIncome: 0, generalDependents: 0, specificDependents: 0, elderlyDependents: 0 };
  for (const X of [3000000, 6000000, 10000000]) {
    const ref = simulate(X, paramsBase);
    const r = findRowForX('シミュレーション1', X);
    const excelTotal = hf.getCellValue({ sheet: sid['シミュレーション1'], row: r, col: colIdx('Y') });
    const excelTedori = hf.getCellValue({ sheet: sid['シミュレーション1'], row: r, col: colIdx('S') });
    check(`X=${X} total`, Math.abs(excelTotal - ref.total) < 1.5, `excel=${excelTotal} ref=${ref.total}`);
    check(`X=${X} kojinTedori`, Math.abs(excelTedori - ref.kojinTedori) < 1.5, `excel=${excelTedori} ref=${ref.kojinTedori}`);
  }

  // ================================================================
  // Test B: 3 officers with distinct params
  // ================================================================
  console.log('\n--- Test B: 3 officers ---');
  setVal('入力', `C${OFFICER_COUNT_ROW}`, 3);

  const officerDefs = [
    { ageOver40: false, hasSpouse: true, spouseIncome: 0, generalDependents: 1, specificDependents: 0, elderlyDependents: 0, X: 6000000 },
    { ageOver40: true, hasSpouse: false, spouseIncome: 0, generalDependents: 0, specificDependents: 0, elderlyDependents: 0, X: 4000000 },
    { ageOver40: true, hasSpouse: true, spouseIncome: 500000, generalDependents: 0, specificDependents: 1, elderlyDependents: 1, X: 3000000 },
  ];

  officerDefs.forEach((def, k) => {
    const r0 = officerRows(k);
    setVal('入力', `C${r0.age}`, def.ageOver40 ? '40歳以上64歳以下' : '40歳未満');
    setVal('入力', `C${r0.spouse}`, def.hasSpouse ? '有' : '無');
    setVal('入力', `C${r0.spouseIncome}`, def.spouseIncome);
    setVal('入力', `C${r0.genDep}`, def.generalDependents);
    setVal('入力', `C${r0.specDep}`, def.specificDependents);
    setVal('入力', `C${r0.elderDep}`, def.elderlyDependents);

    const s = summaryOfficerRows(k);
    setVal('結果サマリー', `C${s.input}`, def.X);
  });
  // officers 4,5 stay at defaults (X=0) but are gated off anyway since 人数=3

  const profitBefore = cellVal('入力', `C${COMPANY_PROFIT_ROW}`);
  check('profitBefore unchanged', profitBefore === 20000000, `=${profitBefore}`);

  // independent reference aggregate computed purely from calc_reference.js
  let sumX = 0, sumKaishaCost = 0, sumKojinTedori = 0;
  const refPerOfficer = officerDefs.map(def => simulate(def.X, {
    profitBefore, ageOver40: def.ageOver40, hasSpouse: def.hasSpouse, spouseIncome: def.spouseIncome,
    generalDependents: def.generalDependents, specificDependents: def.specificDependents, elderlyDependents: def.elderlyDependents,
  }));
  refPerOfficer.forEach(r => { sumX += r.X; sumKaishaCost += r.kaishaCost; sumKojinTedori += r.kojinTedori; });
  const rieki = profitBefore - sumKaishaCost;
  let houjinzei;
  if (rieki <= 0) houjinzei = 0;
  else if (rieki <= 8000000) houjinzei = rieki * 0.2137;
  else houjinzei = 8000000 * 0.2137 + (rieki - 8000000) * 0.3358;
  const naibuRyuho = rieki - houjinzei;
  const grandTotal = naibuRyuho + sumKojinTedori;

  const excelTotalX = cellVal('結果サマリー', `C${SUM_AGG_ROWS.totalX}`);
  const excelTotalCost = cellVal('結果サマリー', `C${SUM_AGG_ROWS.totalCost}`);
  const excelRieki = cellVal('結果サマリー', `C${SUM_AGG_ROWS.rieki}`);
  const excelHoujinzei = cellVal('結果サマリー', `C${SUM_AGG_ROWS.houjinzei}`);
  const excelNaibu = cellVal('結果サマリー', `C${SUM_AGG_ROWS.naibu}`);
  const excelTotalTedori = cellVal('結果サマリー', `C${SUM_AGG_ROWS.totalTedori}`);
  const excelGrandTotal = cellVal('結果サマリー', `C${SUM_AGG_ROWS.grandTotal}`);

  check('agg sumX', excelTotalX === sumX, `excel=${excelTotalX} ref=${sumX}`);
  check('agg sumKaishaCost', Math.abs(excelTotalCost - sumKaishaCost) < 1.5, `excel=${excelTotalCost} ref=${sumKaishaCost}`);
  check('agg rieki', Math.abs(excelRieki - rieki) < 1.5, `excel=${excelRieki} ref=${rieki}`);
  check('agg houjinzei', Math.abs(excelHoujinzei - houjinzei) < 1.5, `excel=${excelHoujinzei} ref=${houjinzei}`);
  check('agg naibuRyuho', Math.abs(excelNaibu - naibuRyuho) < 1.5, `excel=${excelNaibu} ref=${naibuRyuho}`);
  check('agg sumKojinTedori', Math.abs(excelTotalTedori - sumKojinTedori) < 1.5, `excel=${excelTotalTedori} ref=${sumKojinTedori}`);
  check('agg grandTotal', Math.abs(excelGrandTotal - grandTotal) < 1.5, `excel=${excelGrandTotal} ref=${grandTotal}`);

  // marginal invariant: each officer's own シミュレーションk sheet, at that officer's own X row,
  // must show Y == grandTotal (since others are held at their 結果サマリー selection)
  officerDefs.forEach((def, k) => {
    const sheetName = `シミュレーション${k + 1}`;
    const r = findRowForX(sheetName, def.X);
    check(`row found ${sheetName} X=${def.X}`, r >= 0, `row=${r}`);
    if (r < 0) return;
    const y = hf.getCellValue({ sheet: sid[sheetName], row: r, col: colIdx('Y') });
    const v = hf.getCellValue({ sheet: sid[sheetName], row: r, col: colIdx('V') });
    check(`marginal Y ${sheetName}`, Math.abs(y - grandTotal) < 1.5, `excel=${y} ref=${grandTotal}`);
    check(`marginal V(rieki) ${sheetName}`, Math.abs(v - rieki) < 1.5, `excel=${v} ref=${rieki}`);
  });

  // officers 4 & 5 must be fully gated to zero (人数=3)
  const s4 = summaryOfficerRows(3), s5 = summaryOfficerRows(4);
  check('officer4 tedori gated to 0', cellVal('結果サマリー', `C${s4.tedori}`) === 0);
  check('officer4 cost gated to 0', cellVal('結果サマリー', `C${s4.cost}`) === 0);
  check('officer5 tedori gated to 0', cellVal('結果サマリー', `C${s5.tedori}`) === 0);
  check('officer5 cost gated to 0', cellVal('結果サマリー', `C${s5.cost}`) === 0);

  console.log('\n\n' + (allPass ? '★★★ MULTI-OFFICER EXCEL VERIFICATION — ALL PASS ★★★' : '!!! SOME CHECKS FAILED !!!'));
}

main().catch(e => { console.error(e); process.exit(1); });
