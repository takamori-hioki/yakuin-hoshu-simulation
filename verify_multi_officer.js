const fs = require('fs');
const path = require('path');

const html = fs.readFileSync(path.join(__dirname, 'yakuin_hoshu_sim.html'), 'utf8');

const startMarker = '// 計算エンジン';
const endMarker = '// ============================================================\n  // officer card DOM';
const startIdx = html.indexOf(startMarker);
const endIdx = html.indexOf(endMarker, startIdx);
if (startIdx < 0 || endIdx < 0) throw new Error('markers not found');
let block = html.slice(startIdx, endIdx);

const mod = { exports: {} };
const wrapped = `
  ${block}
  module.exports = { simulate, state, makeOfficer, officerParams, othersAggregate, computeAggregate, buildSeriesForOfficer, findOptimal };
`;
const fn = new Function('module', wrapped);
fn(mod);
const { state, makeOfficer, officerParams, othersAggregate, computeAggregate, buildSeriesForOfficer, findOptimal } = mod.exports;

let allPass = true;
function check(label, cond, detail) {
  if (!cond) { allPass = false; console.log(`NG [${label}] ${detail || ''}`); }
}

// ---- scenario: 3 officers, distinct params ----
state.profitBefore = 30000000;
state.officers = [
  Object.assign(makeOfficer(), { ageOver40: false, hasSpouse: true, spouseIncome: 0, generalDependents: 1, selectedX: 6000000 }),
  Object.assign(makeOfficer(), { ageOver40: true, hasSpouse: false, selectedX: 4000000 }),
  Object.assign(makeOfficer(), { ageOver40: true, hasSpouse: true, spouseIncome: 500000, specificDependents: 1, elderlyDependents: 1, selectedX: 3000000 }),
];

const agg = computeAggregate();

// 1) sums are simple additions of independent per-officer simulate() calls
let sumXCheck = 0, sumKaishaCheck = 0, sumTedoriCheck = 0, sumShakaihoKaishaCheck = 0;
state.officers.forEach((o, i) => {
  const r = mod.exports.simulate(o.selectedX, officerParams(i, 0, 0));
  sumXCheck += r.X; sumKaishaCheck += r.kaishaCost; sumTedoriCheck += r.kojinTedori; sumShakaihoKaishaCheck += r.shakaihoKaisha;
});
check('sumX', agg.sumX === sumXCheck, `${agg.sumX} vs ${sumXCheck}`);
check('sumKaishaCost', agg.sumKaishaCost === sumKaishaCheck, `${agg.sumKaishaCost} vs ${sumKaishaCheck}`);
check('sumKojinTedori', agg.sumKojinTedori === sumTedoriCheck, `${agg.sumKojinTedori} vs ${sumTedoriCheck}`);
check('sumShakaihoKaisha', agg.sumShakaihoKaisha === sumShakaihoKaishaCheck, `${agg.sumShakaihoKaisha} vs ${sumShakaihoKaishaCheck}`);

// 2) corp tax recomputed independently from rieki
const rieki = state.profitBefore - agg.sumKaishaCost;
let houjinShotokuBun;
if (rieki <= 0) houjinShotokuBun = 0;
else if (rieki <= 8000000) houjinShotokuBun = rieki * state.corpRateLow;
else houjinShotokuBun = 8000000 * state.corpRateLow + (rieki - 8000000) * state.corpRateHigh;
// 法人住民税の均等割は赤字でも課され、役員が何人いても法人単位で年1回
const houjinzei = houjinShotokuBun + state.houjinKintouwari;
const naibuRyuho = rieki - houjinzei;
check('rieki', agg.rieki === rieki, `${agg.rieki} vs ${rieki}`);
check('houjinzei', Math.abs(agg.houjinzei - houjinzei) < 1, `${agg.houjinzei} vs ${houjinzei}`);
check('naibuRyuho', Math.abs(agg.naibuRyuho - naibuRyuho) < 1, `${agg.naibuRyuho} vs ${naibuRyuho}`);
check('grandTotal', Math.abs(agg.grandTotal - (naibuRyuho + agg.sumKojinTedori)) < 1);

// 3) invariant: each officer's own marginal series, evaluated AT its own currently-selected X,
//    must reproduce the exact same company-wide totals as computeAggregate() (since "others" are
//    held at their current selection too).
state.officers.forEach((o, i) => {
  const others = othersAggregate(i);
  const r = mod.exports.simulate(o.selectedX, officerParams(i, others.cost, others.takeHome));
  check(`marginal-total officer${i}`, Math.abs(r.total - agg.grandTotal) < 1, `${r.total} vs ${agg.grandTotal}`);
  check(`marginal-rieki officer${i}`, Math.abs(r.rieki - agg.rieki) < 1, `${r.rieki} vs ${agg.rieki}`);
  check(`marginal-houjinzei officer${i}`, Math.abs(r.houjinzei - agg.houjinzei) < 1, `${r.houjinzei} vs ${agg.houjinzei}`);
  check(`marginal-naibuRyuho officer${i}`, Math.abs(r.naibuRyuho - agg.naibuRyuho) < 1, `${r.naibuRyuho} vs ${agg.naibuRyuho}`);
});

// 4) buildSeriesForOfficer / findOptimal sanity: optimal point's total must be >= total at officer's
//    own current selection (since it searches over that officer's X holding others fixed)
state.officers.forEach((o, i) => {
  const { pts } = buildSeriesForOfficer(i);
  const opt = findOptimal(pts);
  const others = othersAggregate(i);
  const atCurrent = mod.exports.simulate(o.selectedX, officerParams(i, others.cost, others.takeHome));
  check(`optimal>=current officer${i}`, opt.total >= atCurrent.total - 1, `${opt.total} vs ${atCurrent.total}`);
});

// ---- single-officer regression: with 1 officer, aggregate must equal that officer's own simulate() with no others ----
state.profitBefore = 20000000;
state.officers = [Object.assign(makeOfficer(), { selectedX: 7000000 })];
const agg1 = computeAggregate();
const solo = mod.exports.simulate(7000000, officerParams(0, 0, 0));
check('solo-total', Math.abs(agg1.grandTotal - solo.total) < 1, `${agg1.grandTotal} vs ${solo.total}`);
check('solo-houjinzei', Math.abs(agg1.houjinzei - solo.houjinzei) < 1, `${agg1.houjinzei} vs ${solo.houjinzei}`);
check('solo-naibuRyuho', Math.abs(agg1.naibuRyuho - solo.naibuRyuho) < 1, `${agg1.naibuRyuho} vs ${solo.naibuRyuho}`);

console.log(allPass ? '★★★ MULTI-OFFICER AGGREGATION MATCHES INDEPENDENT CALC — ALL PASS ★★★' : '!!! MISMATCH FOUND !!!');
