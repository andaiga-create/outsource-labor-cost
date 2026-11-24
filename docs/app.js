const $ = (s, r = document) => r.querySelector(s);
const css = n => getComputedStyle(document.documentElement).getPropertyValue(n).trim();
const KINDS = ['AS', '김해공장'];
const SEGS = ['상선', '특수선', '유상'];
const SHOT = location.search.includes('shot');
const BLANK = '(미분류)';
let ROWS = [], DATA, SET, charts = {}, latestY, latestM;
// 기본값: 엑셀 보고서와 같게 최신 연도 + 유상 제외(상선·특수선)
const DEFAULTS = () => ({ 연: new Set([latestY]), 월: new Set(), 구분: new Set(), 선종: new Set(['상선', '특수선']), 업체: new Set() });
let SEL = { 연: new Set(), 월: new Set(), 구분: new Set(), 선종: new Set(), 업체: new Set() };
const CHIP_GROUPS = [['연', '연도'], ['월', '월'], ['구분', '구분'], ['선종', '선종 (유상 = 상선·특수선 외)']];

const fmt = v => { const a = Math.abs(v); return a >= 1e8 ? (v / 1e8).toFixed(2) + '억' : a >= 1e4 ? Math.round(v / 1e4).toLocaleString() + '만' : Math.round(v).toLocaleString(); };
const won = v => Math.round(v).toLocaleString() + '원';
const pct = (v, d = 1) => (v * 100).toFixed(d) + '%';
const sum = (a, f) => a.reduce((s, r) => s + f(r), 0);
const kindColor = () => ({ AS: css('--c1'), 김해공장: css('--c2') });
const segColor = () => ({ 상선: css('--c1'), 특수선: css('--c3'), 유상: css('--seg3') });

function skeleton() {
  document.body.innerHTML = `
  <div class="app">
    <aside class="side">
      <div class="brand"><span class="logo"></span><div><b>외주인건비</b><small>AS · 김해공장 대시보드</small></div></div>
      <div id="slicers"></div>
      <div class="grp dd" id="vendDD"></div>
      <div class="foot" id="foot"></div>
    </aside>
    <main>
      <header><div><h1 id="ttl">외주인건비 현황</h1><p id="sub"></p></div><button class="reset" id="reset">필터 초기화</button></header>
      <section class="kpis" id="kpis"></section>
      <section class="row r3">
        <div class="card"><h2>연도별 외주인건비 <small>막대에 마우스를 올리면 매출 대비 비율</small></h2><div class="box"><canvas id="cYear"></canvas></div></div>
        <div class="card"><h2 id="tMonth">월별 외주인건비</h2><div class="box"><canvas id="cMonth"></canvas></div></div>
        <div class="card"><h2 id="tCum">누적 추이</h2><div class="box"><canvas id="cCum"></canvas></div></div>
      </section>
      <section class="row r4">
        <div class="card"><h2>장비별 TOP 15</h2><div class="box"><canvas id="cEquip"></canvas></div></div>
        <div class="card"><h2>호선별 TOP 10</h2><div class="box"><canvas id="cShip"></canvas></div></div>
        <div class="card"><h2>업체별</h2><div class="box"><canvas id="cVendor"></canvas></div></div>
        <div class="card"><h2>선종별 <small>선종 필터와 무관</small></h2><div class="box"><canvas id="cSeg"></canvas></div></div>
      </section>
    </main>
  </div>`;
}

async function init() {
  skeleton();
  const [d, s] = await Promise.all([
    fetch('data.json', { cache: 'no-cache' }).then(r => r.json()),
    fetch('settings.json', { cache: 'no-cache' }).then(r => r.json()).catch(() => ({}))]);
  DATA = d; SET = s;
  ROWS = d.rows.map(r => { const o = {}; d.dims.forEach((k, i) => o[k] = r[i]); o.v = r[d.dims.length]; o.n = r[d.dims.length + 1]; return o; });
  [latestY, latestM] = d.latest.split('-').map(Number);
  SEL = DEFAULTS();
  buildVendDD();
  $('#reset').onclick = () => { SEL = DEFAULTS(); render(); };
  $('#foot').innerHTML = `데이터 기준일<br><b>${d.latest_date}</b><br>갱신 ${d.generated}<br>원본 ${d.files}개 파일 · ${d.records.toLocaleString()}행`;
  Chart.defaults.font.family = 'Pretendard Variable, Pretendard, "Malgun Gothic", sans-serif';
  Chart.defaults.color = css('--mute');
  if (SHOT) Chart.defaults.animation = false;
  render();
}

function options(g) {
  if (g === '월') return [...Array(12)].map((_, i) => [i + 1, i + 1]);
  if (g === '구분') return KINDS.map(k => [k, k]);
  if (g === '선종') return SEGS.map(k => [k, k]);
  const tot = {};
  ROWS.forEach(r => { tot[r[g]] = (tot[r[g]] || 0) + r.v; });
  let keys = Object.keys(tot);
  if (g === '연') keys.sort((a, b) => b - a); else keys.sort((a, b) => tot[b] - tot[a]);
  return keys.map(k => [g === '연' ? +k : k, k]);
}
const selYears = () => SEL.연.size ? [...SEL.연] : [latestY];
function availMonths() {
  const ys = selYears(), m = new Set();
  ROWS.forEach(r => { if (ys.includes(r.연)) m.add(r.월); });
  return m;
}
function renderSlicers() {
  const avail = availMonths();
  $('#slicers').innerHTML = CHIP_GROUPS.map(([g, label]) => `
    <div class="grp g-${g}"><div class="gh"><span>${label}</span>
      <a data-clear="${g}" class="${SEL[g].size ? 'on' : ''}" title="선택 해제">✕</a></div>
      <div class="chips">${options(g).map(([v, t]) => `<button class="chip ${SEL[g].has(v) ? 'on' : ''}" data-g="${g}" data-v="${v}" ${g === '월' && !avail.has(v) ? 'disabled' : ''}>${t}${g === '연' ? '년' : g === '월' ? '월' : ''}</button>`).join('')}</div></div>`).join('');
  document.querySelectorAll('.chip').forEach(b => b.onclick = () => {
    const g = b.dataset.g, v = (g === '연' || g === '월') ? +b.dataset.v : b.dataset.v;
    SEL[g].has(v) ? SEL[g].delete(v) : SEL[g].add(v); render();
  });
  document.querySelectorAll('[data-clear]').forEach(a => a.onclick = () => { SEL[a.dataset.clear].clear(); render(); });
}
function buildVendDD() {
  const box = $('#vendDD'), opts = options('업체').map(o => o[0]);
  box.innerHTML = `<div class="gh"><span>업체</span><a id="vClear" title="선택 해제">✕</a></div>
    <button class="ddbtn" id="ddBtn"><span id="ddLabel">전체</span><i>▾</i></button>
    <div class="ddpanel" id="ddPanel" hidden>
      <input id="ddSearch" placeholder="업체 검색">
      <div class="ddlist" id="ddList">${opts.map(v => `<label data-v="${v}"><input type="checkbox" value="${v}"><span>${v}</span></label>`).join('')}</div>
      <div class="ddfoot"><button id="ddAll">전체 선택</button><button id="ddNone">선택 해제</button></div></div>`;
  $('#ddBtn').onclick = e => { e.stopPropagation(); const p = $('#ddPanel'); p.hidden = !p.hidden; if (!p.hidden) $('#ddSearch').focus(); };
  document.addEventListener('click', e => { if (!box.contains(e.target)) $('#ddPanel').hidden = true; });
  $('#ddSearch').oninput = e => { const q = e.target.value.trim().toLowerCase(); document.querySelectorAll('#ddList label').forEach(l => l.style.display = l.dataset.v.toLowerCase().includes(q) ? '' : 'none'); };
  $('#ddList').onchange = e => { const v = e.target.value; e.target.checked ? SEL.업체.add(v) : SEL.업체.delete(v); render(); };
  $('#ddAll').onclick = () => { opts.forEach(v => SEL.업체.add(v)); render(); };
  $('#ddNone').onclick = () => { SEL.업체.clear(); render(); };
  $('#vClear').onclick = () => { SEL.업체.clear(); render(); };
}
function syncVendDD() {
  const sel = [...SEL.업체];
  $('#ddLabel').textContent = !sel.length ? '전체' : sel.length <= 2 ? sel.join(', ') : `${sel[0]} 외 ${sel.length - 1}`;
  $('#ddBtn').classList.toggle('on', sel.length > 0);
  $('#vClear').className = sel.length ? 'on' : '';
  document.querySelectorAll('#ddList input').forEach(i => i.checked = SEL.업체.has(i.value));
}

const pass = (r, o = {}) => (o.year || !SEL.연.size || SEL.연.has(r.연)) && (o.month || !SEL.월.size || SEL.월.has(r.월)) &&
  (!SEL.구분.size || SEL.구분.has(r.구분)) && (o.seg || !SEL.선종.size || SEL.선종.has(r.선종)) && (!SEL.업체.size || SEL.업체.has(r.업체));
const group = (rows, kf) => { const m = new Map(); rows.forEach(r => { const k = kf(r); m.set(k, (m.get(k) || 0) + r.v); }); return m; };

function render() {
  renderSlicers(); syncVendDD();
  const rows = ROWS.filter(r => pass(r));
  const ys = selYears().sort((a, b) => a - b);
  const total = sum(rows, r => r.v), cnt = sum(rows, r => r.n);
  const ylabel = SEL.연.size ? ys.join('·') + '년' : '전체 기간';
  const mlabel = SEL.월.size ? ' ' + [...SEL.월].sort((a, b) => a - b).join('·') + '월' : '';
  $('#sub').textContent = `${ylabel}${mlabel} · ${SEL.구분.size ? [...SEL.구분].join('·') : 'AS + 김해공장'} · ${SEL.선종.size ? [...SEL.선종].join('·') : '전체 선종'}`;
  renderKpis(rows, total, cnt, ys);
  chartYear(); chartMonth(ys); chartCum(ys);
  hbar('cEquip', group(rows, r => r.장비 || BLANK), 15);
  hbar('cShip', group(rows, r => r.호선 || BLANK), 10);
  hbar('cVendor', group(rows, r => r.업체 || BLANK), 12);
  chartSeg();
}

function renderKpis(rows, total, cnt, ys) {
  const oneY = SEL.연.size === 1 ? ys[0] : null;
  const k = KINDS.map(x => sum(rows.filter(r => r.구분 === x), r => r.v));
  // 전년 동기: 선택 연도 각각 -1년, 같은 월 범위 (최신 연도는 최신 월까지)
  let prev = null;
  if (SEL.연.size) {
    const lim = y => (y === latestY && !SEL.월.size) ? latestM : 12;
    prev = sum(ROWS.filter(r => pass(r, { year: true }) && SEL.연.has(r.연 + 1) && r.월 <= lim(r.연 + 1)), r => r.v);
  }
  const tgt = oneY && SET.target ? SET.target[oneY] : null;
  const yTotal = oneY ? sum(ROWS.filter(r => r.연 === oneY && pass(r, { month: true })), r => r.v) : null;
  const rev = SEL.연.size ? ys.map(y => (SET.revenue || {})[y]).filter(x => x != null) : [];
  const ratioOk = SEL.연.size && rev.length === ys.length && !SEL.월.size;
  const ratio = ratioOk ? total / sum(rev, x => x) : null;
  const dlt = prev ? (total - prev) / prev : null;
  const kc = kindColor();
  $('#kpis').innerHTML = `
    <div class="card kpi hero"><span class="t">외주인건비 합계</span><span class="v">${fmt(total)}<em>원</em></span><span class="s">${won(total)} · ${cnt.toLocaleString()}건</span></div>
    <div class="card kpi"><span class="t">${oneY ? oneY + '년 목표 대비' : '연간 목표 대비'}</span>${tgt
      ? `<span class="v">${pct(yTotal / tgt, 0)}</span><span class="s">${fmt(yTotal)} / 목표 ${fmt(tgt)}</span><div class="bar"><i class="${yTotal > tgt ? 'over' : ''}" style="width:${Math.min(100, yTotal / tgt * 100)}%"></i></div>`
      : `<span class="v sm">—</span><span class="s">${oneY ? 'settings.json 에 목표 없음' : '연도를 하나만 선택하세요'}</span>`}</div>
    <div class="card kpi"><span class="t">전년 동기 대비</span>${dlt != null
      ? `<span class="v ${dlt > 0 ? 'up' : 'down'}">${dlt > 0 ? '▲' : '▼'} ${pct(Math.abs(dlt))}</span><span class="s">전년 동기 ${fmt(prev)}원</span>`
      : `<span class="v sm">—</span><span class="s">비교할 전년 자료 없음</span>`}</div>
    <div class="card kpi"><span class="t">구분별</span>
      <div class="split"><span><i style="background:${kc.AS}"></i>AS <b>${fmt(k[0])}</b></span></div>
      <div class="split"><span><i style="background:${kc.김해공장}"></i>김해공장 <b>${fmt(k[1])}</b></span></div>
      <span class="s" style="margin-top:4px">AS 비중 ${total ? pct(k[0] / total, 0) : '—'}</span></div>
    <div class="card kpi"><span class="t">매출 대비 비율</span>${ratio != null
      ? `<span class="v">${pct(ratio, 2)}</span><span class="s">매출 ${fmt(sum(rev, x => x))}원 기준</span>`
      : `<span class="v sm">—</span><span class="s">${SEL.월.size ? '월 선택 시 연 매출과 비교 불가' : '연도를 선택하세요 (매출은 settings.json)'}</span>`}</div>`;
}

const baseOpts = (extra = {}) => ({
  responsive: true, maintainAspectRatio: false,
  interaction: { mode: 'index', intersect: false },
  plugins: { legend: { display: false }, tooltip: { callbacks: { label: c => `${c.dataset.label ? c.dataset.label + ': ' : ''}${won(c.parsed.y ?? c.parsed)}` } } },
  scales: { x: { grid: { display: false }, border: { display: false } }, y: { grid: { color: css('--grid') }, border: { display: false }, ticks: { callback: v => fmt(v) } } },
  ...extra
});
function draw(id, cfg) { if (charts[id]) charts[id].destroy(); charts[id] = new Chart($('#' + id), cfg); }
const barDs = (label, data, color, extra = {}) => ({ label, data, backgroundColor: color, borderRadius: 4, borderSkipped: 'start', borderWidth: { top: 2 }, borderColor: 'transparent', maxBarThickness: 34, ...extra });

function chartYear() {
  const rows = ROWS.filter(r => pass(r, { year: true, month: true }));
  const years = [...new Set(ROWS.map(r => r.연))].sort();
  const kc = kindColor(), sel = selYears();
  const dim = (c, y) => SEL.연.size && !sel.includes(y) ? c + '55' : c;
  const by = KINDS.map(k => years.map(y => sum(rows.filter(r => r.연 === y && r.구분 === k), r => r.v)));
  draw('cYear', {
    type: 'bar',
    data: { labels: years.map(String), datasets: KINDS.map((k, i) => barDs(k, by[i], years.map(y => dim(kc[k], y)), { stack: 's' })) },
    options: baseOpts({
      plugins: {
        legend: { display: true, position: 'top', align: 'end', labels: { boxWidth: 10, boxHeight: 10, useBorderRadius: true, borderRadius: 3, generateLabels: () => KINDS.map((k, i) => ({ text: k, fillStyle: kc[k], strokeStyle: kc[k], fontColor: css('--mute'), datasetIndex: i })) } },
        tooltip: {
          callbacks: {
            label: c => `${c.dataset.label}: ${won(c.parsed.y)}`,
            footer: items => {
              const y = +items[0].label, t = by[0][items[0].dataIndex] + by[1][items[0].dataIndex];
              const rv = (SET.revenue || {})[y];
              return [`합계 ${won(t)}`, rv ? `매출 ${fmt(rv)}원 대비 ${pct(t / rv, 2)}` : '매출 미입력'];
            }
          }
        }
      },
      scales: { x: { stacked: true, grid: { display: false }, border: { display: false }, ticks: { autoSkip: false, maxRotation: 0, font: { size: 11 } } }, y: { stacked: true, grid: { color: css('--grid') }, border: { display: false }, ticks: { callback: v => fmt(v) } } },
      onClick: (e, els) => { if (els.length) { const y = years[els[0].index]; SEL.연 = new Set([y]); SEL.월.clear(); render(); } }
    })
  });
}

function chartMonth(ys) {
  const rows = ROWS.filter(r => pass(r, { month: true }) && ys.includes(r.연));
  const kc = kindColor();
  $('#tMonth').textContent = `월별 외주인건비 (${ys.join('·')}년)`;
  const months = [...Array(12)].map((_, i) => i + 1);
  draw('cMonth', {
    type: 'bar',
    data: { labels: months.map(m => m + ''), datasets: KINDS.map(k => barDs(k, months.map(m => sum(rows.filter(r => r.월 === m && r.구분 === k), r => r.v)), months.map(m => SEL.월.size && !SEL.월.has(m) ? kc[k] + '55' : kc[k]), { stack: 's' })) },
    options: baseOpts({
      plugins: { legend: { display: true, position: 'top', align: 'end', labels: { boxWidth: 10, boxHeight: 10, useBorderRadius: true, borderRadius: 3 } },
        tooltip: { callbacks: { title: it => it[0].label + '월', label: c => `${c.dataset.label}: ${won(c.parsed.y)}`, footer: it => `합계 ${won(sum(it, i => i.parsed.y))}` } } },
      scales: { x: { stacked: true, grid: { display: false }, border: { display: false }, ticks: { autoSkip: false, maxRotation: 0, font: { size: 11 } } }, y: { stacked: true, grid: { color: css('--grid') }, border: { display: false }, ticks: { callback: v => fmt(v) } } }
    })
  });
}

function chartCum(ys) {
  const y = Math.max(...ys), py = y - 1;
  const rows = ROWS.filter(r => pass(r, { year: true, month: true }));
  const cum = yr => { let s = 0; return [...Array(12)].map((_, i) => { if (yr === latestY && i + 1 > latestM) return null; s += sum(rows.filter(r => r.연 === yr && r.월 === i + 1), r => r.v); return s; }); };
  const tgt = SET.target ? SET.target[y] : null;
  $('#tCum').textContent = `누적 추이 (${y}년 vs ${py}년)`;
  const ds = [
    { label: `${y}년`, data: cum(y), borderColor: css('--c1'), backgroundColor: css('--c1'), borderWidth: 2, pointRadius: 3, pointHoverRadius: 5, tension: .25 },
    { label: `${py}년`, data: cum(py), borderColor: css('--prior'), backgroundColor: css('--prior'), borderWidth: 2, borderDash: [5, 4], pointRadius: 0, pointHoverRadius: 4, tension: .25 }];
  if (tgt) ds.push({ label: '연간 목표', data: Array(12).fill(tgt), borderColor: css('--bad'), borderWidth: 1.5, borderDash: [2, 3], pointRadius: 0, pointHoverRadius: 0 });
  draw('cCum', {
    type: 'line',
    data: { labels: [...Array(12)].map((_, i) => String(i + 1)), datasets: ds },
    options: baseOpts({ plugins: { legend: { display: true, position: 'top', align: 'end', labels: { boxWidth: 14, boxHeight: 2 } }, tooltip: { callbacks: { title: it => it[0].label + '월', label: c => `${c.dataset.label}: ${won(c.parsed.y)}` } } }, scales: { x: { grid: { display: false }, border: { display: false }, ticks: { maxRotation: 0 } }, y: { grid: { color: css('--grid') }, border: { display: false }, ticks: { callback: v => fmt(v) } } } })
  });
}

function hbar(id, map, n) {
  const arr = [...map.entries()].filter(([, v]) => v).sort((a, b) => b[1] - a[1]).slice(0, n);
  draw(id, {
    type: 'bar',
    data: { labels: arr.map(a => a[0]), datasets: [barDs('', arr.map(a => a[1]), css('--c1'), { maxBarThickness: 16 })] },
    options: baseOpts({
      indexAxis: 'y', interaction: { mode: 'nearest', axis: 'y', intersect: false },
      plugins: { legend: { display: false }, tooltip: { callbacks: { title: it => it[0].label, label: c => won(c.parsed.x) } } },
      scales: { x: { grid: { color: css('--grid') }, border: { display: false }, ticks: { callback: v => fmt(v), maxTicksLimit: 4 } }, y: { grid: { display: false }, border: { display: false }, ticks: { autoSkip: false, font: { size: 11.5 }, color: css('--ink2'), callback: function (v) { const t = String(this.getLabelForValue(v)).replace(/^\(주\)\s*|\s*주식회사$/g, ''); return t.length > 9 ? t.slice(0, 8) + '…' : t; } } } }
    })
  });
}

function chartSeg() {
  const rows = ROWS.filter(r => pass(r, { seg: true }));
  const sc = segColor(), vals = SEGS.map(s => sum(rows.filter(r => r.선종 === s), r => r.v)), tot = sum(vals, x => x);
  draw('cSeg', {
    type: 'doughnut',
    data: { labels: SEGS, datasets: [{ data: vals, backgroundColor: SEGS.map(s => sc[s]), borderColor: css('--card'), borderWidth: 2, hoverOffset: 4 }] },
    options: {
      responsive: true, maintainAspectRatio: false, cutout: '62%',
      plugins: {
        legend: { position: 'bottom', labels: { boxWidth: 10, boxHeight: 10, useBorderRadius: true, borderRadius: 3, generateLabels: ch => SEGS.map((s, i) => ({ text: `${s} ${fmt(vals[i])} (${tot ? pct(vals[i] / tot, 0) : '-'})`, fillStyle: sc[s], strokeStyle: sc[s], fontColor: css('--ink2'), index: i })) } },
        tooltip: { callbacks: { label: c => `${c.label}: ${won(c.parsed)} (${pct(c.parsed / tot)})` } }
      }
    }
  });
}

init();
