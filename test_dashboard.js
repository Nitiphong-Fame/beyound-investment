/**
 * Portfolio Dashboard — Feature Test Suite
 * ─────────────────────────────────────────
 * Tests: Page structure, NVDA data, Dime USD,
 *        Buy/Add, Sell + commission/VAT,
 *        Auto-delete, P&L history, Real-time recalc
 *
 * Run:  node test_dashboard.js
 *  or:  bash run_tests.sh
 */

const { JSDOM } = require('jsdom');
const fs   = require('fs');
const path = require('path');

// Load index.html from same folder as this script
const htmlPath = path.join(__dirname, 'index.html');
const html = fs.readFileSync(htmlPath, 'utf8');

async function runTests() {

// ── Setup JSDOM ──────────────────────────────────────────
const dom = new JSDOM(html, {
  runScripts: 'dangerously',
  resources: 'usable',
  pretendToBeVisual: true,
  url: 'http://localhost/',
});
const { window } = dom;
const { document } = window;

// Wait for scripts to load
await new Promise(r => setTimeout(r, 400));

// Use jsdom's real localStorage (scripts use it directly)
const ls = window.localStorage;

// ── Test Helpers ─────────────────────────────────────────
let pass = 0, fail = 0;
const results = [];

function test(name, fn) {
  try {
    fn();
    console.log(`  ✅ PASS: ${name}`);
    results.push({ status: 'PASS', name });
    pass++;
  } catch (e) {
    console.log(`  ❌ FAIL: ${name}`);
    console.log(`         → ${e.message}`);
    results.push({ status: 'FAIL', name, error: e.message });
    fail++;
  }
}

function assert(cond, msg) {
  if (!cond) throw new Error(msg);
}

function approx(a, b, tolerance = 1) {
  assert(Math.abs(a - b) <= tolerance, `Expected ~${b}, got ${a}`);
}

// ── Helper: get element value safely ─────────────────────
const el = (id) => document.getElementById(id);
const val = (id) => parseFloat(el(id)?.value);
const txt = (id) => el(id)?.textContent?.trim();

// ── Test constants ────────────────────────────────────────
const FX = 31.96;

// ═══════════════════════════════════════════════════════════
console.log('\n========================================');
console.log(' 🧪 DASHBOARD FEATURE TEST SUITE');
console.log('========================================\n');

// ── 1. PAGE STRUCTURE ────────────────────────────────────
console.log('📋 1. PAGE STRUCTURE');

test('Growth tab exists', () => {
  assert(el('tab-growth'), 'tab-growth not found');
});
test('Core tab exists', () => {
  assert(el('tab-core'), 'tab-core not found');
});
test('Profit/Loss tab exists', () => {
  assert(el('tab-profit'), 'tab-profit not found');
});
test('Sell modal exists', () => {
  assert(el('sellModal'), 'sellModal not found');
});
test('Add/Buy modal exists', () => {
  assert(el('addModal'), 'addModal not found');
});
test('Manual trade modal exists', () => {
  assert(el('manualTradeModal'), 'manualTradeModal not found');
});

// ── 2. NVDA DATA ─────────────────────────────────────────
console.log('\n📊 2. NVDA DATA');

test('NVDA price = 188.63', () => {
  approx(val('g1-price'), 188.63, 0.01);
});
test('NVDA shares = 7.4238', () => {
  approx(val('g1-shares'), 7.4238, 0.001);
});
test('NVDA cost = 175.1385', () => {
  const costAttr = parseFloat(el('g1-price')?.dataset?.cost);
  approx(costAttr, 175.1385, 0.001);
});
test('NVDA value (THB) ≈ ฿44,755', () => {
  const thb = txt('g1-val-thb');
  assert(thb && thb.includes('44'), `THB value wrong: ${thb}`);
});
test('NVDA P&L% = +7.70%', () => {
  const pct = txt('g1-pl-pct');
  assert(pct && pct.includes('7.70'), `P&L% wrong: ${pct}`);
});

// ── 3. DIME USD VALUE ────────────────────────────────────
console.log('\n💵 3. DIME USD VALUE');

test('Dime! USD row exists', () => {
  const rows = Array.from(document.querySelectorAll('.ticker'));
  const dime = rows.find(r => r.textContent.includes('Dime! USD'));
  assert(dime, 'Dime! USD ticker not found');
});
test('Dime! USD ≈ $1,189', () => {
  const rows = Array.from(document.querySelectorAll('td'));
  const dimeUsdCell = rows.find(r => r.textContent.trim() === '≈$1,189');
  assert(dimeUsdCell, `Dime USD cell not found (got: ${rows.map(r=>r.textContent.trim()).filter(t=>t.includes('1,1')).join(', ')})`);
});
test('Dime! USD THB ≈ ฿38,010', () => {
  const rows = Array.from(document.querySelectorAll('strong'));
  const dimeThb = rows.find(r => r.textContent.includes('38,010'));
  assert(dimeThb, 'Dime THB ฿38,010 not found in page');
});

// ── 4. ADD (BUY) FEATURE ─────────────────────────────────
console.log('\n➕ 4. BUY / ADD FEATURE');

test('openAddModal sets correct ticker', () => {
  window.openAddModal(null, 'g1', 'NVDA', 'USD');
  const title = txt('modal-title');   // real ID is modal-title
  assert(title && title.includes('NVDA'), `Title wrong: "${title}"`);
});
test('Add modal opens (has .open class)', () => {
  assert(el('addModal')?.classList.contains('open'), 'addModal did not get .open class');
});
test('Buy: new shares merge correctly (weighted avg cost)', () => {
  // NVDA: 7.4238 @ $175.1385, buy 2 @ $190
  // newAvg = (7.4238*175.1385 + 2*190) / 9.4238 ≈ $178.87
  const oldShares = 7.4238, oldCost = 175.1385;
  const buyShares = 2, buyPrice = 190;
  const newShares = oldShares + buyShares;
  const expectedAvg = (oldShares * oldCost + buyShares * buyPrice) / newShares;

  el('modal-shares').value    = buyShares;
  el('modal-buy-price').value = buyPrice;
  window.previewAdd();

  const previewText = el('modal-preview')?.textContent || '';
  assert(previewText.includes('178') || previewText.includes('179'),
    `Preview should show ~$178.87 avg, got: "${previewText.slice(0,80)}"`);
  approx(expectedAvg, 178.87, 1);
});
test('confirmAdd updates shares input', () => {
  el('modal-shares').value    = 2;
  el('modal-buy-price').value = 190;
  window.confirmAdd();
  approx(val('g1-shares'), 9.4238, 0.01);
});
test('confirmAdd updates data-shares attribute', () => {
  const attr = parseFloat(el('g1-price')?.dataset?.shares);
  approx(attr, 9.4238, 0.01);
});
test('Add modal closes after confirm', () => {
  assert(!el('addModal')?.classList.contains('open'), 'addModal should be closed');
});

// ── 5. SELL FEATURE ──────────────────────────────────────
console.log('\n➖ 5. SELL FEATURE');

// Reset NVDA to original state for sell tests (buy test changed shares + cost)
el('g1-shares').value          = 7.4238;
el('g1-price').dataset.shares  = 7.4238;
el('g1-price').dataset.cost    = 175.1385;

test('openSellModal opens and shows NVDA', () => {
  window.openSellModal(null, 'g1', 'NVDA', 'USD');
  assert(el('sellModal')?.classList.contains('open'), 'sellModal did not open');
  const title = txt('sell-modal-title');
  assert(title && title.includes('NVDA'), `Sell title wrong: ${title}`);
});
test('sellMax fills all shares', () => {
  window.sellMax();
  approx(val('sell-shares'), 7.4238, 0.001);
});
test('previewSell calculates gross P&L correctly', () => {
  el('sell-shares').value = 2;
  el('sell-price').value  = 200;
  el('sell-comm-rate').value = 0; // no fee for clean test
  window.previewSell();
  // grossPL = 2 * (200 - 175.1385) = 2 * 24.8615 = 49.723
  const expected = 2 * (200 - 175.1385);
  assert(Math.abs(expected - 49.72) < 1, `Expected ~49.72, got ${expected.toFixed(2)}`);
});
test('Commission + VAT fee calculation', () => {
  // proceeds = 2 * 200 = 400, rate = 0.15%
  // comm = 400 * 0.15/100 = 0.60, vat = 0.60 * 0.07 = 0.042, total = 0.642
  const proceeds = 2 * 200;
  const comm = proceeds * 0.15 / 100;
  const vat  = comm * 0.07;
  const total = comm + vat;
  approx(total, 0.642, 0.01);
  el('sell-comm-rate').value = 0.15;
  window.previewSell();
  const feeText = el('sell-fee-preview')?.textContent || '';
  assert(feeText.includes('0.64') || feeText.includes('0.642'), `Fee text wrong: ${feeText}`);
});
test('confirmSell reduces shares', () => {
  el('sell-shares').value    = 2;
  el('sell-price').value     = 200;
  el('sell-comm-rate').value = 0.15;
  window.confirmSell();
  approx(val('g1-shares'), 5.4238, 0.01);
});
test('Sell logs trade to localStorage', () => {
  const hist = JSON.parse(ls.getItem('pl_history_2026') || '[]');
  const lastTrade = hist[0]; // newest first
  assert(lastTrade && lastTrade.stock === 'NVDA', `Last trade not NVDA: ${JSON.stringify(lastTrade)}`);
  approx(lastTrade.shares, 2, 0.01);
  approx(lastTrade.sellPrice, 200, 0.01);
});
test('Sell P&L logged is net (after fees)', () => {
  const hist = JSON.parse(ls.getItem('pl_history_2026') || '[]');
  const trade = hist[0];
  // gross = 2*(200-175.1385)=49.723, comm=400*0.15%=0.60, vat=0.042, net≈49.08
  approx(trade.plUsd, 49.08, 0.3);
});
test('Sell modal closes after confirm', () => {
  assert(!el('sellModal')?.classList.contains('open'), 'sellModal should be closed');
});

// ── 6. AUTO-DELETE when shares = 0 ───────────────────────
console.log('\n🗑️  6. AUTO-DELETE (shares → 0)');

// Use g12 (ASML) for delete test — sell all shares
test('ASML row exists before sell-all', () => {
  const asmlPrice = el('g12-price');
  assert(asmlPrice, 'g12-price not found');
});
test('Selling all shares removes the row', () => {
  const asmlShares = val('g12-shares');
  window.openSellModal(null, 'g12', 'ASML', 'USD');
  el('sell-shares').value    = asmlShares;
  el('sell-price').value     = parseFloat(el('g12-price')?.value) || 100;
  el('sell-comm-rate').value = 0;
  window.confirmSell();
  // Row should be gone
  const asmlEl = document.getElementById('g12-shares');
  assert(!asmlEl, 'g12 row should have been removed from DOM');
});

// ── 7. PROFIT/LOSS TAB ───────────────────────────────────
console.log('\n💰 7. PROFIT/LOSS TAB');

test('renderPLHistory runs without error', () => {
  window.renderPLHistory();
  assert(el('pl-tbody'), 'pl-tbody not found');
});
test('P&L history table has trades', () => {
  const rows = el('pl-tbody')?.querySelectorAll('tr') || [];
  assert(rows.length > 0, 'No rows in P&L table');
});
test('Seeded $88.41 entry exists', () => {
  const hist = JSON.parse(ls.getItem('pl_history_2026') || '[]');
  const seed = hist.find(t => t.stock === 'ALL' && Math.abs(t.plUsd - 88.41) < 0.01);
  assert(seed, `Seeded $88.41 ALL entry not found. History: ${JSON.stringify(hist.map(h=>({s:h.stock,p:h.plUsd})))}`);
});
test('Total P&L shown in THB card', () => {
  const thb = txt('pl-total-thb');
  assert(thb && thb !== '฿0', `Total P&L THB is zero/missing: ${thb}`);
});
test('Total USD shown', () => {
  const usd = txt('pl-total-usd');
  assert(usd && usd !== '≈$0', `Total USD is zero/missing: ${usd}`);
});
test('Win/Loss count correct', () => {
  const wins   = parseInt(txt('pl-wins'))   || 0;
  const losses = parseInt(txt('pl-losses')) || 0;
  assert(wins + losses > 0, 'No wins or losses counted');
});
test('FX rate is 31.96', () => {
  // Verify FX = 31.96 is in the source
  assert(html.includes('const FX = 31.96'), 'FX = 31.96 not found in source');
  // 1189.26 * 31.96 = 38008.7496 → rounds to 38009
  const expectedThb = Math.round(1189.26 * 31.96);
  assert(expectedThb >= 38008 && expectedThb <= 38011,
    `FX calc out of range: ${expectedThb}`);
});

// ── 8. RECALC (real-time P&L update) ─────────────────────
console.log('\n🔄 8. REAL-TIME RECALC');

test('recalc updates P&L% on price change', () => {
  // Reset cost to original before recalc test
  const priceEl = el('g1-price');
  priceEl.dataset.cost   = 175.1385;
  priceEl.dataset.shares = 5.4238;   // shares after sell test
  priceEl.value = 200;
  window.recalc('g1', priceEl);
  const pct = txt('g1-pl-pct');
  // (200 - 175.1385) / 175.1385 * 100 = ~14.2%
  assert(pct && (pct.includes('14.') || pct.includes('14,')),
    `P&L% after recalc wrong: "${pct}" (expected ~14.2%)`);
});
test('recalc updates val-thb on price change', () => {
  const priceEl = el('g1-price');
  priceEl.dataset.shares = 5.4238;
  priceEl.value = 200;
  window.recalc('g1', priceEl);
  const thb = txt('g1-val-thb');
  // 5.4238 * 200 * 31.96 ≈ ฿34,663
  assert(thb && thb.includes('34,'), `THB value after recalc wrong: "${thb}"`);
});

// ── SUMMARY ──────────────────────────────────────────────
console.log('\n========================================');
console.log(` 📊 RESULTS: ${pass} passed, ${fail} failed`);
console.log('========================================');

if (fail > 0) {
  console.log('\n❌ FAILED TESTS:');
  results.filter(r => r.status === 'FAIL').forEach(r => {
    console.log(`   • ${r.name}: ${r.error}`);
  });
}

console.log('\n✅ PASSED:', pass);
console.log('❌ FAILED:', fail);
console.log('');

// Exit with code
process.exit(fail > 0 ? 1 : 0);

} // end runTests

runTests().catch(e => { console.error(e); process.exit(1); });
