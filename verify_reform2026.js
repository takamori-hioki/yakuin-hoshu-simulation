// 令和8年度税制改正(基礎控除・給与所得控除・配偶者控除/配偶者特別控除)の反映を、
// 手計算で検証した期待値と突き合わせる独立検証スクリプト。
const { simulate } = require('./calc_reference');

let allPass = true;
function check(label, actual, expected, tol) {
  tol = tol || 1;
  const ok = Math.abs(actual - expected) <= tol;
  if (!ok) allPass = false;
  console.log(`${ok ? 'OK ' : 'NG!'} ${label}: actual=${actual} expected=${expected}`);
}

const base = { profitBefore: 20000000, ageOver40: false, hasSpouse: true, spouseIncome: 0,
  generalDependents: 0, specificDependents: 0, elderlyDependents: 0 };

console.log('--- 給与所得控除(令和8・9年分) ---');
// X<=1,900,000 -> 740,000固定
check('kyuyoKoujo X=1,000,000', simulate(1000000, base).kyuyoKoujo, 740000);
check('kyuyoKoujo X=1,900,000', simulate(1900000, base).kyuyoKoujo, 740000);
// X=2,000,000 -> ハの特例により220万円まで74万円で連続(680,000ではない)
check('kyuyoKoujo X=2,000,000', simulate(2000000, base).kyuyoKoujo, 740000);
// X=2,200,000 -> 境界。ここで通常式(660,000+80,000=740,000)に滑らかに接続
check('kyuyoKoujo X=2,200,000 (境界)', simulate(2200000, base).kyuyoKoujo, 740000);
// X=2,300,000 -> 通常式: 2,300,000*0.3+80,000=770,000 (220万円超で式適用)
check('kyuyoKoujo X=2,300,000', simulate(2300000, base).kyuyoKoujo, 770000);
// X=3,600,000 -> 3,600,000*0.3+80,000=1,160,000
check('kyuyoKoujo X=3,600,000', simulate(3600000, base).kyuyoKoujo, 1160000);
// 高額側は改正なし: X=10,000,000 -> 1,950,000固定
check('kyuyoKoujo X=10,000,000 (unchanged)', simulate(10000000, base).kyuyoKoujo, 1950000);

console.log('\n--- 基礎控除(所得税・令和8・9年分) ---');
// 合計所得(給与所得)489万円以下->104万円。X=6,000,000: kyuyoShotoku=6,000,000-1,640,000=4,360,000<=4,890,000
check('kiso via kyuyoShotoku=4,360,000 (X=6,000,000)', simulate(6000000, base).kyuyoKoujo && (simulate(6000000, base).kyuyoShotoku), 4360000);
{
  const r = simulate(6000000, base);
  // shotokuKoujoGoukei = shakaihoKojin + kiso(1,040,000) + haigu(380,000, spouseIncome=0)
  const kiso = r.shotokuKoujoGoukei - r.shakaihoKojin - 380000; // no dependents
  check('kiso amount (kyuyoShotoku=4,360,000<=489万)', kiso, 1040000);
}
{
  // X such that kyuyoShotoku lands in 489万超655万以下: try X=8,000,000 -> kyuyoKoujo=8,000,000*0.1+1,100,000=1,900,000, kyuyoShotoku=6,100,000 (489万<6,100,000<=655万)
  const r = simulate(8000000, base);
  check('kyuyoShotoku for X=8,000,000', r.kyuyoShotoku, 6100000);
  const kiso = r.shotokuKoujoGoukei - r.shakaihoKojin - 380000;
  check('kiso amount (655万超489万以下 bracket)', kiso, 670000);
}
{
  // kyuyoShotoku in 655万超2350万以下: X=20,000,000 -> kyuyoKoujo=1,950,000(capped), kyuyoShotoku=18,050,000
  // filerTier=-1 here (kyuyoShotoku>10,000,000) so haigu=0, not 380,000 - account for that.
  const r = simulate(20000000, base);
  check('kyuyoShotoku for X=20,000,000', r.kyuyoShotoku, 18050000);
  const kiso = r.shotokuKoujoGoukei - r.shakaihoKojin - 0;
  check('kiso amount (655万超2350万以下 bracket)', kiso, 620000);
}
{
  // kyuyoShotoku > 2350万: X=27,000,000 -> kyuyoKoujo=1,950,000, kyuyoShotoku=25,050,000 -> phase-out bottom (0)
  // filerTier=-1 here too, so haigu=0.
  const r = simulate(27000000, base);
  check('kyuyoShotoku for X=27,000,000', r.kyuyoShotoku, 25050000);
  const kiso = r.shotokuKoujoGoukei - r.shakaihoKojin - 0;
  check('kiso amount (>2500万 existing phase-out, should be 0)', kiso, 0);
}

console.log('\n--- 配偶者控除・配偶者特別控除(所得税, 本人900万円以下) ---');
// spouse収入0 -> 合計所得0 -> 38万円
check('spouseIncome=0', spouseHaigu(0), 380000);
// spouse収入136万(1,360,000) -> 合計所得=1,360,000-740,000=620,000 -> まだ38万円(62万円超95万円以下バンドも38万円)
check('spouseIncome=1,360,000 (境界)', spouseHaigu(1360000), 380000);
// spouse収入169万(1,690,000) -> 合計所得=1,690,000-740,000=950,000 -> ちょうど95万円->38万円のはず(下限含む)
check('spouseIncome=1,690,000 (合計所得95万ちょうど)', spouseHaigu(1690000), 380000);
// spouse収入175万2000(合計所得175.2万-74万=101.2万->100万円超105万円以下->31万円)
check('spouseIncome=1,752,000 (合計所得101.2万)', spouseHaigu(1752000), 310000);
// spouse収入184万(合計所得184-74=110万ちょうど->105万円超110万円以下->26万円)
check('spouseIncome=1,840,000 (合計所得110万)', spouseHaigu(1840000), 260000);
// spouse収入190万(合計所得190-74=116万->115万円超120万円以下->16万円)
check('spouseIncome=1,900,000 (合計所得116万)', spouseHaigu(1900000), 160000);
// spouse収入207万(合計所得: kyuyoKoujoFunc(2,070,000)=2,070,000*0.3+80,000=701,000; 合計所得=2,070,000-701,000=1,369,000->130万円超133万円以下->3万円
check('spouseIncome=2,070,000 (境界、合計所得約136.9万)', spouseHaigu(2070000), 30000);
// spouse収入220万(合計所得: kyuyoKoujoFunc(2,200,000)=2,200,000*0.3+80,000=740,000; 合計所得=1,460,000->133万円超->対象外0円
check('spouseIncome=2,200,000 (133万円超で対象外)', spouseHaigu(2200000), 0);

function spouseHaigu(spouseIncome) {
  const p = Object.assign({}, base, { spouseIncome });
  const r = simulate(6000000, p); // kyuyoShotoku=4,360,000 <=900万 tier0
  return r.shotokuKoujoGoukei - r.shakaihoKojin - 1040000; // 引いてkisoの1,040,000を除けばhaiguが残る(扶養0のため)
}

console.log('\n--- 配偶者控除・配偶者特別控除(住民税, 本人900万円以下) ---');
function spouseHaiguJumin(spouseIncome) {
  const p = Object.assign({}, base, { spouseIncome });
  const r = simulate(6000000, p);
  const juminKazeiShotoku_withHaigu = r.kyuyoShotoku - (r.shakaihoKojin + 430000 + 0 /*haiguJumin unknown*/);
  // 代わりにjuminTotalから逆算するのは複雑なので、haiguJuminを直接検証するために内部式を再現する
  return null;
}
// 住民税側は簡易検証: spouseIncome=0のとき住民税基礎控除43万+配偶者控除33万が使われているかを、
// juminKazeiShotokuの計算結果から確認する(kyuyoShotoku=4,360,000, shakaihoKojin=859,500)
{
  const r0 = simulate(6000000, base);
  // juminKazeiShotoku = roundDown(max(kyuyoShotoku-(shakaiho+kisoJumin+haiguJumin+fuyouJumin),0),-3)
  // juminShotokuwari = juminKazeiShotoku*0.10 なので、juminShotokuwari/0.10 = juminKazeiShotoku(1000円単位に丸め後)
  const juminKazeiShotoku = r0.juminShotokuwari / 0.10;
  const impliedDeduction = r0.kyuyoShotoku - juminKazeiShotoku; // shakaihoKojin+kisoJumin(43万)+haiguJumin(?)
  const impliedHaiguJumin = impliedDeduction - r0.shakaihoKojin - 430000;
  check('haiguJumin (spouseIncome=0, tier0)', impliedHaiguJumin, 330000, 1500); // roundDown to 1000 introduces up to ~1000 slack
}

console.log('\n' + (allPass ? '★★★ 令和8年度税制改正 反映値 — ALL PASS ★★★' : '!!! MISMATCH FOUND !!!'));
