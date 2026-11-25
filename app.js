// 산업연관표 분석 웹앱 — 엑셀을 SheetJS로 읽고 브라우저에서 계산
const FILES = {
  large:  ['대분류(33)',  'io-2023-large.xlsx'],
  medium: ['중분류(83)',  'io-2023-medium.xlsx'],
  small:  ['소분류(165)', 'io-2023-small.xlsx'],
  basic:  ['기본부문(380)', 'io-2023-basic.xlsx'],
};
const WB = {}, CACHE = {}, MODEL = {};
const $ = id => document.getElementById(id);
const clean = s => String(s ?? '').replace(/\s+/g, ' ').trim();
const fmt = (x, d = 0) => (x == null || isNaN(x)) ? '' : Number(x).toLocaleString('ko-KR', { maximumFractionDigits: d });
const num = x => (typeof x === 'number' ? x : 0);

function unitOf(name) {
  if (/거래표/.test(name)) return '단위: 백만원';
  if (/유발액/.test(name)) return '단위: 백만원 (구성비 행은 %, 유발계수 행은 계수)';
  return '단위: 계수(단위 없음)';
}

// ---------- 불러오기 ----------
async function load(key) {
  if (WB[key]) return;
  $('status').textContent = FILES[key][0] + ' 불러오는 중…';
  const res = await fetch(FILES[key][1]);
  if (!res.ok) throw new Error(FILES[key][1] + ' 를 찾을 수 없음 (' + res.status + ')');
  const buf = await res.arrayBuffer();
  WB[key] = XLSX.read(buf, { type: 'array' });
  CACHE[key] = {};
  MODEL[key] = buildModel(key);
  $('status').innerHTML = '<span class="ok">' + FILES[key][0] + ' 로딩 완료 (산업 ' + MODEL[key].n + '개)</span>';
  refreshClsSelect();
}
function rowsOf(key, sheet) {
  CACHE[key][sheet] ??= XLSX.utils.sheet_to_json(WB[key].Sheets[sheet], { header: 1, raw: true, defval: null });
  return CACHE[key][sheet];
}

// ---------- 모델 구성 ----------
function indexSheet(R) {
  const col = {}, row = {}, inds = [];
  (R[4] || []).forEach((c, j) => { if (c != null) col[String(c).trim()] = j; });
  for (let i = 6; i < R.length; i++) {
    const c = R[i] && R[i][0]; if (c == null) continue;
    const k = String(c).trim(); row[k] = i;
    if (!k.startsWith('9')) inds.push({ code: k, name: clean(R[i][1]), i });
  }
  return { col, row, inds };
}
function buildModel(key) {
  const T = rowsOf(key, 'A표_총거래표(생산자)'), t = indexSheet(T);
  const D = rowsOf(key, 'A표_국산거래표(생산자)'), d = indexSheet(D);
  const Lr = rowsOf(key, '생산유발계수'), l = indexSheet(Lr);
  const n = t.inds.length;
  const g = (R, ix, rc, cc) => num(R[ix.row[rc]]?.[ix.col[cc]]);
  const inds = t.inds.map((o, a) => {
    const c = o.code, r = o.i;
    const X = num(T[r][t.col['9310']]);
    const VA = g(T, t, '9690', c), TI = g(T, t, '9790', c);
    return {
      code: c, name: o.name, X, VA, TI,
      W: g(T, t, '9610', c), IMID: g(T, t, '9590', c),
      EXP: num(T[r][t.col['9140']]), DEXP: num(D[d.inds[a].i][d.col['9140']]),
      IMP: num(T[r][t.col['9329']]), SUP: num(T[r][t.col['9390']]),
      v: TI ? VA / TI : 0,
      infl: g(Lr, l, '9801', c), sens: num(Lr[l.inds[a].i][l.col['9802']]),
    };
  });
  const L = l.inds.map(o => t.inds.map(p => num(Lr[o.i][l.col[p.code]])));
  return { n, inds, L };
}

// ---------- 표 그리기 ----------
function drawTable(el, heads, rows, opt = {}) {
  let data = rows.slice(), sortCol = -1, asc = true;
  const render = () => {
    el.innerHTML = '<thead><tr>' + heads.map((h, i) => `<th data-i="${i}" class="${i < (opt.left ?? 1) ? 'l' : ''}">${h}</th>`).join('') +
      '</tr></thead><tbody>' + data.map((r, ri) => '<tr data-r="' + ri + '">' + r.map((v, i) =>
        `<td class="${i < (opt.left ?? 1) ? 'l' : ''}">${typeof v === 'number' ? fmt(v, opt.dec?.[i] ?? 0) : (v ?? '')}</td>`).join('') + '</tr>').join('') + '</tbody>';
    el.querySelectorAll('th').forEach(th => th.onclick = () => {
      const i = +th.dataset.i; asc = sortCol === i ? !asc : true; sortCol = i;
      data.sort((a, b) => (a[i] > b[i] ? 1 : a[i] < b[i] ? -1 : 0) * (asc ? 1 : -1)); render();
    });
    if (opt.onRow) el.querySelectorAll('tbody tr').forEach(tr => tr.onclick = () => opt.onRow(data[+tr.dataset.r]));
  };
  render();
}

// ---------- 2. 원자료 탐색 ----------
const curKey = () => $('cls').value;
function refreshClsSelect() {
  const sel = $('cls'), old = sel.value;
  sel.innerHTML = Object.keys(WB).map(k => `<option value="${k}">${FILES[k][0]}</option>`).join('');
  if (old && WB[old]) sel.value = old;
  onClsChange();
}
function onClsChange() {
  const k = curKey(); if (!k) return;
  $('sheet').innerHTML = WB[k].SheetNames.map(s => `<option>${s}</option>`).join('');
  $('simInd').innerHTML = MODEL[k].inds.map((o, i) => `<option value="${i}">${o.code} ${o.name}</option>`).join('');
  showSheet(); showOverview();
}
function showSheet() {
  const k = curKey(); if (!k) return;
  const s = $('sheet').value, R = rowsOf(k, s), q = $('q').value.trim().toLowerCase();
  $('sheetUnit').textContent = unitOf(s);
  const rows = [];
  for (let i = 6; i < R.length; i++) {
    if (R[i][0] == null) continue;
    const code = String(R[i][0]).trim(), name = clean(R[i][1]);
    if (!q || code.toLowerCase().includes(q) || name.toLowerCase().includes(q)) rows.push([code, name, i]);
  }
  drawTable($('rowList'), ['코드', '산업(행)'], rows.map(r => [r[0], r[1]]), { left: 2, onRow: r => showRow(R, rows.find(x => x[0] === r[0] && x[1] === r[1])[2]) });
}
function showRow(R, i) {
  const cq = $('cq').value.trim().toLowerCase();
  $('rowTitle').innerHTML = '<b>' + String(R[i][0]).trim() + ' ' + clean(R[i][1]) + '</b> · <span class="unit">' + unitOf($('sheet').value) + '</span>';
  const out = [];
  for (let j = 2; j < R[4].length; j++) {
    if (R[4][j] == null) continue;
    const code = String(R[4][j]).trim(), name = clean(R[5][j]);
    if (!cq || code.toLowerCase().includes(cq) || name.toLowerCase().includes(cq)) out.push([code, name, R[i][j]]);
  }
  drawTable($('rowDetail'), ['열 코드', '열 이름', '값'], out, { left: 2, dec: [0, 0, 4] });
  window._lastRow = [R, i];
}

// ---------- 3. 산업구조 ----------
function showOverview() {
  const k = curKey(); if (!k) return;
  const m = MODEL[k];
  const rows = m.inds.map(o => [o.code, o.name, o.X, o.VA,
    o.TI ? 100 * o.VA / o.TI : 0, o.TI ? 100 * o.IMID / o.TI : 0, o.X ? 100 * o.EXP / o.X : 0, o.SUP ? 100 * o.IMP / o.SUP : 0, o.infl, o.sens]);
  drawTable($('overview'), ['코드', '산업', '총산출(백만원)', '부가가치(백만원)', '부가가치율(%)', '중간투입률(%)', '수출비중(%)', '수입의존도(%)', '영향력계수', '감응도계수'],
    rows, { left: 2, dec: [0, 0, 0, 0, 1, 1, 1, 1, 3, 3] });
  window._ov = rows;
}
$('csv').onclick = () => {
  const head = '코드,산업,총산출(백만원),부가가치(백만원),부가가치율(%),중간투입률(%),수출비중(%),수입의존도(%),영향력계수,감응도계수\n';
  const body = (window._ov || []).map(r => r.map(v => '"' + v + '"').join(',')).join('\n');
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob(['\ufeff' + head + body], { type: 'text/csv' }));
  a.download = 'industry-structure-' + curKey() + '.csv'; a.click();
};

// ---------- 4. 시뮬레이션 ----------
function simulate(k, dF, title) {
  const m = MODEL[k], n = m.n, dX = new Array(n).fill(0);
  for (let a = 0; a < n; a++) { let s = 0; for (let b = 0; b < n; b++) if (dF[b]) s += m.L[a][b] * dF[b]; dX[a] = s; }
  const sF = dF.reduce((x, y) => x + y, 0), sX = dX.reduce((x, y) => x + y, 0);
  const dV = dX.map((x, a) => x * m.inds[a].v), sV = dV.reduce((x, y) => x + y, 0);
  return { dX, dV, sF, sX, sV, title };
}
function showSim(r, k) {
  const m = MODEL[k];
  $('simSummary').innerHTML = `<p><b>${r.title}</b><br>최종수요 변화 ${fmt(r.sF)} 백만원 → 생산 변화 ${fmt(r.sX)} 백만원, 부가가치 변화 ${fmt(r.sV)} 백만원` +
    (r.sF ? ` (생산유발승수 ${(r.sX / r.sF).toFixed(3)}, 단위 없음)` : '') + `</p>`;
  const rows = m.inds.map((o, a) => [o.code, o.name, r.dX[a], r.dV[a]]).sort((x, y) => Math.abs(y[2]) - Math.abs(x[2]));
  drawTable($('simTable'), ['코드', '산업', '생산 변화(백만원)', '부가가치 변화(백만원)'], rows, { left: 2 });
}
$('runInd').onclick = () => {
  const k = curKey(); if (!k) return alert('먼저 자료를 불러오세요');
  const a = +$('simInd').value, amt = +$('simAmt').value, dF = new Array(MODEL[k].n).fill(0); dF[a] = amt;
  showSim(simulate(k, dF, MODEL[k].inds[a].name + ' 국산품 최종수요 ' + fmt(amt) + ' 백만원 변화'), k);
};
function exportShock(k, pct) { return simulate(k, MODEL[k].inds.map(o => o.DEXP * pct / 100), '전 산업 수출 ' + pct + '% 변화'); }
$('runExp').onclick = () => {
  const k = curKey(); if (!k) return alert('먼저 자료를 불러오세요');
  showSim(exportShock(k, +$('simRate').value), k);
};

// ---------- 5. 분류 비교 ----------
$('cmp').onclick = () => {
  const rows = [];
  for (const k of ['large', 'medium', 'small', 'basic']) {
    if (!MODEL[k]) continue;
    const m = MODEL[k], s = f => m.inds.reduce((x, o) => x + o[f], 0), r = exportShock(k, +$('simRate').value);
    rows.push([FILES[k][0], m.n, s('X') / 1e6, 100 * s('VA') / s('TI'), 100 * s('EXP') / s('X'), 100 * s('IMP') / s('SUP'),
      Math.max(...m.inds.map(o => o.infl)), r.sX / 1e6, r.sV / 1e6]);
  }
  if (!rows.length) return alert('먼저 자료를 불러오세요');
  drawTable($('cmpTable'), ['분류', '산업 수', '총산출(조원)', '부가가치율(%)', '수출비중(%)', '수입의존도(%)', '최대 영향력계수', '수출변화 시 생산 변화(조원)', '부가가치 변화(조원)'],
    rows, { dec: [0, 0, 0, 1, 1, 1, 3, 1, 1] });
};

// ---------- 이벤트 ----------
const run = async k => { try { await load(k); } catch (e) { $('status').innerHTML = '<span class="err">' + e.message + '</span>'; } };
document.querySelectorAll('[data-load]').forEach(b => b.onclick = () => run(b.dataset.load));
$('loadAll').onclick = async () => { for (const k of ['large', 'medium', 'small', 'basic']) await run(k); };
$('cls').onchange = onClsChange; $('sheet').onchange = showSheet; $('q').oninput = showSheet;
$('cq').oninput = () => { if (window._lastRow) showRow(...window._lastRow); };
