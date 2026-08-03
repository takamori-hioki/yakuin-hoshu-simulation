const fs = require('fs');
const path = require('path');
const { simulate: refSimulate } = require('./calc_reference');

const html = fs.readFileSync(path.join(__dirname, 'yakuin_hoshu_sim.html'), 'utf8');

// Extract the calc engine block from the artifact's <script> (tables through simulate function)
const startMarker = '// 計算エンジン';
const endMarker = '// ============================================================\n  // state';
const startIdx = html.indexOf(startMarker);
const endIdx = html.indexOf(endMarker, startIdx);
if (startIdx < 0 || endIdx < 0) throw new Error('markers not found');
let block = html.slice(startIdx, endIdx);

// eval in isolated scope, exposing `simulate`
const mod = { exports: {} };
const wrapped = `
  ${block}
  module.exports = { simulate };
`;
const fn = new Function('module', wrapped);
fn(mod);
const artifactSimulate = mod.exports.simulate;

const paramsBase = { profitBefore: 20000000, ageOver40: false, hasSpouse: true, spouseIncome: 0,
  generalDependents: 0, specificDependents: 0, elderlyDependents: 0,
  kenpoRate: 0.1012, kaigoRate: 0.0162, kouseiRate: 0.183, kodomoRate: 0.0036, shienkinRate: 0.0023,
  corpRateLow: 0.2137, corpRateHigh: 0.3358 };

const refParamsBase = { profitBefore: 20000000, ageOver40: false, hasSpouse: true, spouseIncome: 0,
  generalDependents: 0, specificDependents: 0, elderlyDependents: 0 };

let allPass = true;
const testXs = [0, 1000000, 3000000, 5000000, 6000000, 8500000, 10000000, 15000000, 24000000, 36000000, 50000000];
const testCases = [
  { label: 'base', p: {}, refP: {} },
  { label: '40-64 age', p: { ageOver40: true }, refP: { ageOver40: true } },
  { label: 'spouse+deps', p: { generalDependents: 1, specificDependents: 1, elderlyDependents: 1 }, refP: { generalDependents: 1, specificDependents: 1, elderlyDependents: 1 } },
  { label: 'no spouse', p: { hasSpouse: false }, refP: { hasSpouse: false } },
  { label: 'spouse partial phase-out', p: { spouseIncome: 1800000 }, refP: { spouseIncome: 1800000 } },
  { label: 'tokutei shinzoku partial', p: { specificDependents: 2, specificDependentsWithIncome: 1, specificDependentIncome: 1650000 },
    refP: { specificDependents: 2, specificDependentsWithIncome: 1, specificDependentIncome: 1650000 } },
];

for (const tc of testCases) {
  const p = Object.assign({}, paramsBase, tc.p);
  const refP = Object.assign({}, refParamsBase, tc.refP);
  for (const X of testXs) {
    const a = artifactSimulate(X, p);
    const r = refSimulate(X, refP);
    const keys = ['stdKenpo','stdKousei','shakaihoKojin','shakaihoKaisha','kyuyoKoujo','kyuyoShotoku','kazeiShotoku','shotokuzei','fukkou','shotokuzeiTotal','juminShotokuwari','juminTotal','kojinFutanGoukei','kojinTedori','kaishaCost','rieki','houjinzei','naibuRyuho','total'];
    for (const k of keys) {
      const diff = Math.abs((a[k]||0) - (r[k]||0));
      if (diff > 1) {
        allPass = false;
        console.log(`NG [${tc.label}] X=${X} ${k}: artifact=${a[k]} ref=${r[k]} diff=${diff}`);
      }
    }
  }
}

console.log(allPass ? '★★★ ARTIFACT CALC MATCHES REFERENCE — ALL PASS ★★★' : '!!! MISMATCH FOUND !!!');
