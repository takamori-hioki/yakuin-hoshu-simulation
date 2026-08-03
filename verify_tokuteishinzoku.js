const { simulate } = require('./calc_reference');

let allPass = true;
function check(label, actual, expected, tol) {
  tol = tol || 1;
  const ok = Math.abs(actual - expected) <= tol;
  if (!ok) allPass = false;
  console.log(`${ok ? 'OK ' : 'NG!'} ${label}: actual=${actual} expected=${expected}`);
}

const base = { profitBefore: 20000000, ageOver40: false, hasSpouse: false,
  generalDependents: 0, specificDependents: 1, elderlyDependents: 0 };

// helper: extract implied (fuyouTax, fuyouJumin) contribution via shotokuKoujoGoukei / juminKazeiShotoku diff
function fuyouAmounts(p) {
  const r = simulate(6000000, p); // kyuyoShotoku=4,360,000<=489万 -> kiso=1,040,000, haigu=0(no spouse)
  const fuyouTax = r.shotokuKoujoGoukei - r.shakaihoKojin - 1040000;
  const juminKazeiShotoku = r.juminShotokuwari / 0.10;
  const impliedDeduction = r.kyuyoShotoku - juminKazeiShotoku; // shakaihoKojin + 430,000(kisoJumin) + fuyouJumin
  const fuyouJumin = impliedDeduction - r.shakaihoKojin - 430000;
  return { fuyouTax, fuyouJumin };
}

console.log('--- 特定親族特別控除なし(通常の特定扶養控除、1人) ---');
{
  const { fuyouTax, fuyouJumin } = fuyouAmounts(base);
  check('fuyouTax (regular tokutei, no income)', fuyouTax, 630000);
  check('fuyouJumin (regular tokutei, no income)', fuyouJumin, 450000, 1500);
}

console.log('\n--- 特定親族特別控除あり(1人、収入により段階的に減額) ---');
const cases = [
  // [収入, 期待所得税控除額, 期待住民税控除額]
  [1200000, 630000, 450000],   // 給与収入120万 -> 合計所得120万-74万=46万 -> 62万以下なので通常の特定扶養控除(満額)
  [1400000, 630000, 450000],   // 給与収入140万 -> 合計所得66万 -> 62万超85万以下 -> 63万/45万
  [1600000, 610000, 450000],   // 給与収入160万 -> 合計所得86万 -> 85万超90万以下 -> 61万/45万
  [1650000, 510000, 450000],   // 給与収入165万 -> 合計所得91万 -> 90万超95万以下 -> 51万/45万
  [1700000, 410000, 410000],   // 給与収入170万 -> 合計所得96万 -> 95万超100万以下 -> 41万/41万
  [1750000, 310000, 310000],   // 給与収入175万 -> 合計所得101万 -> 100万超105万以下 -> 31万/31万
  [1800000, 210000, 210000],   // 給与収入180万 -> 合計所得106万 -> 105万超110万以下 -> 21万/21万
  [1850000, 110000, 110000],   // 給与収入185万 -> 合計所得111万 -> 110万超115万以下 -> 11万/11万
  [1900000, 60000, 60000],     // 給与収入190万 -> 合計所得116万 -> 115万超120万以下 -> 6万/6万
  [1950000, 30000, 30000],     // 給与収入195万 -> 合計所得121万 -> 120万超123万以下 -> 3万/3万
  [2000000, 0, 0],             // 給与収入200万 -> 合計所得126万 -> 123万超 -> 対象外
];
for (const [income, expTax, expJumin] of cases) {
  const p = Object.assign({}, base, { specificDependentsWithIncome: 1, specificDependentIncome: income });
  const { fuyouTax, fuyouJumin } = fuyouAmounts(p);
  check(`収入=${income} 所得税控除`, fuyouTax, expTax);
  check(`収入=${income} 住民税控除`, fuyouJumin, expJumin, 1500);
}

console.log('\n--- 特定扶養親族2人中1人だけ収入がある場合(残り1人は満額) ---');
{
  const p = Object.assign({}, base, { specificDependents: 2, specificDependentsWithIncome: 1, specificDependentIncome: 1600000 });
  const { fuyouTax, fuyouJumin } = fuyouAmounts(p);
  // 残り1人(regularSpecificCount=1) 63万 + 収入ある1人 61万 = 124万
  check('2人中1人収入あり(所得税)', fuyouTax, 630000 + 610000);
  check('2人中1人収入あり(住民税)', fuyouJumin, 450000 + 450000, 1500);
}

console.log('\n' + (allPass ? '★★★ 特定親族特別控除 — ALL PASS ★★★' : '!!! MISMATCH FOUND !!!'));
