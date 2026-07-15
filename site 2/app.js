/* Pelada Musical — Analytics (Instagram + YouTube + Spotify)
   Lê dados.xlsx (3 abas) e monta o painel. Sem backend: roda em qualquer
   hospedagem estática (Netlify, GitHub Pages, etc). */

const PLATFORMS = {
  instagram: {
    key: 'instagram', sheet: 'Instagram', label: 'Instagram', color: '#e1306c',
    primaryKey: 'Visualizações', primaryLabel: 'Visualizações', unit: 'views',
    totalKey: 'Seguidores (total)', totalLabel: 'Seguidores',
    newKey: 'Novos seguidores informados', newLabel: 'Novos seguidores',
    secondaryKey: '% Não seguidores', secondaryLabel: 'Público novo (não seguidores)', secondaryIsPct: true,
    variationKey: 'Variação de views informada (%)',
    obsKey: 'Observação do Instagram'
  },
  youtube: {
    key: 'youtube', sheet: 'YouTube', label: 'YouTube', color: '#ff3b3b',
    primaryKey: 'Visualizações', primaryLabel: 'Visualizações', unit: 'views',
    totalKey: 'Inscritos (total)', totalLabel: 'Inscritos',
    newKey: 'Novos inscritos', newLabel: 'Novos inscritos',
    secondaryKey: 'Horas assistidas', secondaryLabel: 'Horas assistidas', secondaryIsPct: false,
    variationKey: null,
    obsKey: 'Observação'
  },
  spotify: {
    key: 'spotify', sheet: 'Spotify', label: 'Spotify', color: '#1db954',
    primaryKey: 'Streams', primaryLabel: 'Streams', unit: 'streams',
    totalKey: 'Seguidores (total)', totalLabel: 'Seguidores',
    newKey: 'Novos seguidores', newLabel: 'Novos seguidores',
    secondaryKey: 'Ouvintes mensais', secondaryLabel: 'Ouvintes mensais', secondaryIsPct: false,
    variationKey: null,
    obsKey: 'Observação'
  }
};

let DATA = { instagram: [], youtube: [], spotify: [] };
let CHARTS = {};
const nfInt = new Intl.NumberFormat('pt-BR');
const nfCompact = new Intl.NumberFormat('pt-BR', { notation: 'compact', maximumFractionDigits: 1 });
const nfPct = (v) => (v === null || v === undefined || isNaN(v)) ? '—' : `${v > 0 ? '+' : ''}${v.toFixed(1)}%`;

/* ---------------- carregar planilha ---------------- */

async function loadExcel(manual) {
  setStatus('loading', manual ? 'Atualizando...' : 'Carregando dados.xlsx...');
  try {
    const res = await fetch('dados.xlsx?t=' + Date.now());
    if (!res.ok) throw new Error('arquivo não encontrado');
    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array', cellDates: true });
    ingestWorkbook(wb);
  } catch (err) {
    setStatus('error', 'Não achei dados.xlsx aqui perto. Use "Abrir outra planilha".');
    console.error(err);
  }
}

function ingestWorkbook(wb) {
  Object.values(PLATFORMS).forEach(cfg => {
    DATA[cfg.key] = parseSheet(wb, cfg);
  });
  const anyData = Object.values(DATA).some(rows => rows.length);
  if (!anyData) {
    setStatus('error', 'Planilha carregada, mas nenhuma aba reconhecida (Instagram/YouTube/Spotify).');
    return;
  }
  setStatus('ok', `Atualizado em ${new Date().toLocaleString('pt-BR')}`);
  document.getElementById('updated').textContent =
    `Dados sincronizados • ${new Date().toLocaleDateString('pt-BR', { day: '2-digit', month: 'long', year: 'numeric' })}`;
  renderAll();
}

function parseSheet(wb, cfg) {
  const ws = wb.Sheets[cfg.sheet];
  if (!ws) return [];
  const json = XLSX.utils.sheet_to_json(ws, { defval: null });
  const rows = json
    .filter(r => r['Mês/Ano'] && r[cfg.primaryKey] !== null && r[cfg.primaryKey] !== undefined && r[cfg.primaryKey] !== 0 || (r['Mês/Ano'] && Number(r[cfg.primaryKey]) > 0))
    .map(r => ({
      label: String(r['Mês/Ano']).trim(),
      primary: Number(r[cfg.primaryKey]) || 0,
      total: Number(r[cfg.totalKey]) || 0,
      newF: r[cfg.newKey] !== null && r[cfg.newKey] !== undefined ? Number(r[cfg.newKey]) : null,
      secondary: r[cfg.secondaryKey] !== null && r[cfg.secondaryKey] !== undefined ? Number(r[cfg.secondaryKey]) : null,
      variationRaw: cfg.variationKey ? r[cfg.variationKey] : null,
      obs: r[cfg.obsKey] || ''
    }));
  // variação calculada quando não informada
  rows.forEach((row, i) => {
    if (row.variationRaw !== null && row.variationRaw !== undefined && row.variationRaw !== '') {
      row.variation = Number(row.variationRaw);
    } else if (i > 0 && rows[i - 1].primary) {
      row.variation = ((row.primary - rows[i - 1].primary) / rows[i - 1].primary) * 100;
    } else {
      row.variation = null;
    }
  });
  return rows;
}

function setStatus(state, text) {
  const dot = document.getElementById('dot');
  dot.className = state === 'ok' ? 'ok' : '';
  dot.style.background = state === 'error' ? '#ff6767' : (state === 'ok' ? '' : '#ffb020');
  document.getElementById('statusText').textContent = text;
}

/* ---------------- render geral ---------------- */

function renderAll() {
  renderOverview();
  Object.keys(PLATFORMS).forEach(renderPlatform);
  renderHistorico();
}

function bestMonth(rows) {
  if (!rows.length) return null;
  return rows.reduce((a, b) => (b.primary > a.primary ? b : a));
}

function renderOverview() {
  const ig = DATA.instagram, yt = DATA.youtube, sp = DATA.spotify;
  const sum = rows => rows.reduce((s, r) => s + r.primary, 0);
  const totalAll = sum(ig) + sum(yt) + sum(sp);

  document.getElementById('ov-total').textContent = totalAll ? nfCompact.format(totalAll) : '—';
  document.getElementById('ov-total-sub').textContent = totalAll
    ? `${nfInt.format(totalAll)} views/streams somando as 3 redes`
    : 'Ainda sem dados suficientes';

  document.getElementById('ov-ig').textContent = ig.length ? nfCompact.format(sum(ig)) : '—';
  document.getElementById('ov-yt').textContent = yt.length ? nfCompact.format(sum(yt)) : '—';
  document.getElementById('ov-sp').textContent = sp.length ? nfCompact.format(sum(sp)) : '—';
  document.getElementById('ov-ig-sub').textContent = ig.length ? `${ig.length} meses` : 'sem dados na aba Instagram';
  document.getElementById('ov-yt-sub').textContent = yt.length ? `${yt.length} meses` : 'sem dados na aba YouTube';
  document.getElementById('ov-sp-sub').textContent = sp.length ? `${sp.length} meses` : 'sem dados na aba Spotify';

  const followersSum = (ig.length ? ig[ig.length - 1].total : 0) + (yt.length ? yt[yt.length - 1].total : 0) + (sp.length ? sp[sp.length - 1].total : 0);
  document.getElementById('ov-followers').textContent = followersSum ? nfInt.format(followersSum) : '—';

  // melhor mês geral (maior primary entre todas as redes)
  const candidates = [
    ig.length && { ...bestMonth(ig), platform: 'Instagram' },
    yt.length && { ...bestMonth(yt), platform: 'YouTube' },
    sp.length && { ...bestMonth(sp), platform: 'Spotify' }
  ].filter(Boolean);
  if (candidates.length) {
    const best = candidates.reduce((a, b) => (b.primary > a.primary ? b : a));
    document.getElementById('ov-bestmonth').textContent = best.label;
    document.getElementById('ov-bestmonth-sub').textContent = `${best.platform} • ${nfInt.format(best.primary)}`;
  }

  const monthsCount = Math.max(ig.length, yt.length, sp.length);
  document.getElementById('ov-months').textContent = monthsCount || '—';

  // rede com maior crescimento (últimos 2 meses)
  const growth = Object.entries({ Instagram: ig, YouTube: yt, Spotify: sp })
    .map(([name, rows]) => rows.length >= 2 ? { name, v: rows[rows.length - 1].variation } : null)
    .filter(Boolean);
  if (growth.length) {
    const top = growth.reduce((a, b) => (b.v > a.v ? b : a));
    document.getElementById('ov-topgrowth').textContent = top.name;
  } else {
    document.getElementById('ov-topgrowth').textContent = '—';
  }

  drawOverviewChart(ig, yt, sp);
  drawShareChart(sum(ig), sum(yt), sum(sp));
  document.getElementById('ov-executive').innerHTML = buildOverviewExecutive(ig, yt, sp, candidates);
}

function drawOverviewChart(ig, yt, sp) {
  const labels = [ig, yt, sp].reduce((longest, rows) => rows.length > longest.length ? rows.map(r => r.label) : longest, []);
  const normalize = rows => {
    if (!rows.length) return labels.map(() => null);
    const max = Math.max(...rows.map(r => r.primary), 1);
    const byLabel = Object.fromEntries(rows.map(r => [r.label, (r.primary / max) * 100]));
    return labels.map(l => byLabel[l] ?? null);
  };
  drawChart('ovChart', {
    type: 'line',
    data: {
      labels,
      datasets: [
        { label: 'Instagram', data: normalize(ig), borderColor: PLATFORMS.instagram.color, backgroundColor: 'transparent', tension: .35, spanGaps: true },
        { label: 'YouTube', data: normalize(yt), borderColor: PLATFORMS.youtube.color, backgroundColor: 'transparent', tension: .35, spanGaps: true },
        { label: 'Spotify', data: normalize(sp), borderColor: PLATFORMS.spotify.color, backgroundColor: 'transparent', tension: .35, spanGaps: true }
      ]
    },
    options: baseChartOptions({ y: { display: false }, legend: true, tooltipSuffix: '% do pico da própria rede' })
  });
}

function drawShareChart(a, b, c) {
  const total = a + b + c || 1;
  drawChart('ovShareChart', {
    type: 'doughnut',
    data: {
      labels: ['Instagram', 'YouTube', 'Spotify'],
      datasets: [{ data: [a, b, c], backgroundColor: [PLATFORMS.instagram.color, PLATFORMS.youtube.color, PLATFORMS.spotify.color], borderWidth: 0 }]
    },
    options: {
      responsive: true, maintainAspectRatio: false, cutout: '68%',
      plugins: {
        legend: { position: 'bottom', labels: { color: '#8aa093', boxWidth: 10, font: { size: 11 } } },
        tooltip: { callbacks: { label: (ctx) => `${ctx.label}: ${((ctx.raw / total) * 100).toFixed(1)}%` } }
      }
    }
  });
}

function buildOverviewExecutive(ig, yt, sp, candidates) {
  if (!candidates.length) return '<p>Ainda não há dados suficientes. Preencha as abas Instagram, YouTube e Spotify em dados.xlsx.</p>';
  const parts = [];
  const best = candidates.reduce((a, b) => (b.primary > a.primary ? b : a));
  parts.push(`<p><b>${best.platform}</b> foi o destaque do período, com <b>${nfInt.format(best.primary)}</b> em ${best.label.toLowerCase()}.</p>`);
  [['Instagram', ig], ['YouTube', yt], ['Spotify', sp]].forEach(([name, rows]) => {
    if (!rows.length) { parts.push(`<p><b>${name}:</b> sem dados carregados ainda.</p>`); return; }
    const last = rows[rows.length - 1];
    const trend = last.variation === null ? 'estável' : (last.variation >= 0 ? `crescimento de ${nfPct(last.variation)}` : `queda de ${nfPct(last.variation)}`);
    parts.push(`<p><b>${name}:</b> ${last.label} fechou com ${nfInt.format(last.primary)} e ${trend} frente ao mês anterior. Total acumulado no período: ${nfInt.format(rows.reduce((s, r) => s + r.primary, 0))}.</p>`);
  });
  return parts.join('');
}

/* ---------------- render por plataforma ---------------- */

function renderPlatform(key) {
  const cfg = PLATFORMS[key];
  const rows = DATA[key];
  const view = document.getElementById('view-' + key);

  if (!rows.length) {
    view.innerHTML = `<div class="panel"><div class="empty">Nenhum dado encontrado na aba <b>${cfg.sheet}</b> de dados.xlsx ainda.<br>Adicione as colunas e linhas conforme a aba "Como Atualizar" da planilha.</div></div>`;
    return;
  }

  view.innerHTML = platformTemplate(cfg);

  const totalPrimary = rows.reduce((s, r) => s + r.primary, 0);
  const best = bestMonth(rows);
  const last = rows[rows.length - 1];
  const avgPrimary = totalPrimary / rows.length;
  const avgSecondary = rows.filter(r => r.secondary !== null).length
    ? rows.reduce((s, r) => s + (r.secondary || 0), 0) / rows.filter(r => r.secondary !== null).length : null;

  q(view, '[data-role="heroLabel"]').textContent = `${cfg.primaryLabel} acumuladas`;
  q(view, '[data-role="total"]').textContent = nfCompact.format(totalPrimary);
  q(view, '[data-role="totalSub"]').textContent = `${nfInt.format(totalPrimary)} ${cfg.unit} em ${rows.length} ${rows.length === 1 ? 'mês' : 'meses'}`;
  q(view, '[data-role="bestMonth"]').textContent = best.label;
  q(view, '[data-role="bestSub"]').textContent = `${nfInt.format(best.primary)} ${cfg.unit}`;

  q(view, '[data-role="kpiFollowers"]').textContent = nfInt.format(last.total);
  q(view, '[data-role="kpiFollowersLabel"]').textContent = cfg.totalLabel + ' atuais';
  const gain = rows.length > 1 ? last.total - rows[0].total : last.newF;
  q(view, '[data-role="kpiFollowersSub"]').textContent = gain !== null ? `${gain >= 0 ? '+' : ''}${nfInt.format(gain)} no período` : '—';

  q(view, '[data-role="kpiSecondary"]').textContent = avgSecondary === null ? '—' : (cfg.secondaryIsPct ? `${(avgSecondary * 100).toFixed(1)}%` : nfInt.format(Math.round(avgSecondary)));
  q(view, '[data-role="kpiSecondaryLabel"]').textContent = cfg.secondaryLabel;

  q(view, '[data-role="kpiAvg"]').textContent = nfCompact.format(avgPrimary);
  q(view, '[data-role="kpiMonths"]').textContent = rows.length;

  drawChart('chart-' + key + '-primary', {
    type: 'line',
    data: {
      labels: rows.map(r => r.label),
      datasets: [{
        label: cfg.primaryLabel, data: rows.map(r => r.primary),
        borderColor: cfg.color, backgroundColor: hexAlpha(cfg.color, .12), fill: true, tension: .35
      }]
    },
    options: baseChartOptions({ y: { display: true } })
  });

  drawChart('chart-' + key + '-total', {
    type: 'bar',
    data: {
      labels: rows.map(r => r.label),
      datasets: [{ label: cfg.totalLabel, data: rows.map(r => r.total), backgroundColor: hexAlpha(cfg.color, .55), borderRadius: 6 }]
    },
    options: baseChartOptions({ y: { display: true } })
  });

  // seletor de mês
  const select = q(view, '[data-role="monthSelect"]');
  select.innerHTML = rows.map((r, i) => `<option value="${i}">${r.label}</option>`).join('');
  select.value = rows.length - 1;
  const paintMonth = () => paintMonthDetail(view, cfg, rows[Number(select.value)], rows[Number(select.value) - 1]);
  select.onchange = paintMonth;
  paintMonth();

  q(view, '[data-role="executive"]').innerHTML = buildPlatformExecutive(cfg, rows);
}

function platformTemplate(cfg) {
  return `
  <section class="hero">
    <div class="hero-main">
      <span data-role="heroLabel"></span>
      <strong data-role="total"></strong>
      <p data-role="totalSub"></p>
    </div>
    <div class="hero-main" style="justify-content:center">
      <span>Melhor mês</span>
      <strong data-role="bestMonth" style="color:${cfg.color}"></strong>
      <p data-role="bestSub"></p>
    </div>
  </section>
  <section class="kpis">
    <article><span data-role="kpiFollowersLabel"></span><strong data-role="kpiFollowers"></strong><small data-role="kpiFollowersSub"></small></article>
    <article><span data-role="kpiSecondaryLabel"></span><strong data-role="kpiSecondary"></strong><small>Média do período</small></article>
    <article><span>Média mensal de ${cfg.primaryLabel.toLowerCase()}</span><strong data-role="kpiAvg"></strong><small>Período completo</small></article>
    <article><span>Meses analisados</span><strong data-role="kpiMonths"></strong><small>Histórico disponível</small></article>
  </section>
  <section class="layout">
    <div class="panel wide">
      <div class="panel-head"><div><h2>${cfg.primaryLabel} por mês</h2><p>Evolução do alcance em ${cfg.label}</p></div></div>
      <div class="chart"><canvas id="chart-${cfg.key}-primary"></canvas></div>
    </div>
    <div class="panel">
      <div class="panel-head"><div><h2>${cfg.totalLabel}</h2><p>Crescimento da comunidade</p></div></div>
      <div class="chart"><canvas id="chart-${cfg.key}-total"></canvas></div>
    </div>
  </section>
  <section class="layout">
    <div class="panel">
      <div class="panel-head">
        <div><h2>Resumo mensal</h2><p>Escolha um mês para analisar</p></div>
        <select data-role="monthSelect"></select>
      </div>
      <div class="month-grid" data-role="monthGrid"></div>
    </div>
    <div class="panel">
      <div class="panel-head"><div><h2>Leitura executiva</h2><p>Texto pronto para reunião</p></div></div>
      <div class="executive" data-role="executive"></div>
    </div>
  </section>`;
}

function paintMonthDetail(view, cfg, row, prevRow) {
  if (!row) return;
  const grid = q(view, '[data-role="monthGrid"]');
  const cells = [
    [cfg.primaryLabel, nfInt.format(row.primary)],
    ['Variação', row.variation === null ? '—' : nfPct(row.variation)],
    [cfg.totalLabel, nfInt.format(row.total)],
    [cfg.newLabel || 'Novos', row.newF === null ? '—' : nfInt.format(row.newF)],
    [cfg.secondaryLabel, row.secondary === null ? '—' : (cfg.secondaryIsPct ? `${(row.secondary * 100).toFixed(1)}%` : nfInt.format(row.secondary))],
    ['Observação', row.obs || '—']
  ];
  grid.innerHTML = cells.map(([label, val]) => `<div class="cell"><span>${label}</span><strong>${val}</strong></div>`).join('');
}

function buildPlatformExecutive(cfg, rows) {
  const last = rows[rows.length - 1];
  const total = rows.reduce((s, r) => s + r.primary, 0);
  const trend = last.variation === null ? 'estabilidade' : (last.variation >= 0 ? `alta de ${nfPct(last.variation)}` : `queda de ${nfPct(last.variation)}`);
  return `<p>Em <b>${last.label}</b>, ${cfg.label} registrou <b>${nfInt.format(last.primary)}</b> ${cfg.unit}, uma ${trend} em relação ao mês anterior.</p>
  <p>${cfg.totalLabel} atuais: <b>${nfInt.format(last.total)}</b>. Acumulado do período analisado: <b>${nfInt.format(total)}</b> ${cfg.unit} em ${rows.length} meses.</p>
  ${last.obs ? `<p>${last.obs}</p>` : ''}`;
}

/* ---------------- histórico ---------------- */

function renderHistorico() {
  const filter = document.getElementById('histFilter');
  const paint = () => {
    const val = filter.value;
    const rowsHtml = [];
    Object.keys(PLATFORMS).forEach(key => {
      if (val !== 'all' && val !== key) return;
      const cfg = PLATFORMS[key];
      DATA[key].forEach(r => {
        rowsHtml.push(`<tr>
          <td><span class="badge ${key === 'instagram' ? 'ig' : key === 'youtube' ? 'yt' : 'sp'}">${cfg.label}</span></td>
          <td>${r.label}</td>
          <td>${nfInt.format(r.primary)}</td>
          <td class="${r.variation === null ? '' : (r.variation >= 0 ? 'up' : 'down')}">${r.variation === null ? '—' : nfPct(r.variation)}</td>
          <td>${nfInt.format(r.total)}</td>
          <td>${r.newF === null ? '—' : nfInt.format(r.newF)}</td>
        </tr>`);
      });
    });
    document.getElementById('histRows').innerHTML = rowsHtml.join('') || '<tr><td colspan="6" class="empty">Sem dados ainda.</td></tr>';
  };
  filter.onchange = paint;
  paint();
}

/* ---------------- charts helpers ---------------- */

function baseChartOptions(opts = {}) {
  return {
    responsive: true, maintainAspectRatio: false,
    interaction: { mode: 'index', intersect: false },
    plugins: {
      legend: { display: !!opts.legend, position: 'bottom', labels: { color: '#8aa093', boxWidth: 10, font: { size: 11 } } },
      tooltip: opts.tooltipSuffix ? { callbacks: { label: (ctx) => `${ctx.dataset.label}: ${ctx.parsed.y?.toFixed(1)} ${opts.tooltipSuffix}` } } : {}
    },
    scales: {
      x: { ticks: { color: '#5e746a', font: { size: 11 } }, grid: { color: 'rgba(255,255,255,.04)' } },
      y: { display: opts.y?.display !== false, ticks: { color: '#5e746a', font: { size: 11 } }, grid: { color: 'rgba(255,255,255,.04)' } }
    }
  };
}

function drawChart(canvasId, config) {
  const el = document.getElementById(canvasId);
  if (!el) return;
  if (CHARTS[canvasId]) CHARTS[canvasId].destroy();
  CHARTS[canvasId] = new Chart(el.getContext('2d'), config);
}

function hexAlpha(hex, alpha) {
  const r = parseInt(hex.slice(1, 3), 16), g = parseInt(hex.slice(3, 5), 16), b = parseInt(hex.slice(5, 7), 16);
  return `rgba(${r},${g},${b},${alpha})`;
}

/* ---------------- navegação ---------------- */

function setActiveView(hash) {
  const key = (hash || '#overview').replace('#', '');
  document.querySelectorAll('.view').forEach(v => v.style.display = 'none');
  const target = document.getElementById('view-' + key) || document.getElementById('view-overview');
  target.style.display = 'block';
  document.querySelectorAll('nav a').forEach(a => a.classList.toggle('active', a.dataset.p === key));
  const titles = { overview: 'Pelada Musical', instagram: 'Instagram', youtube: 'YouTube', spotify: 'Spotify', historico: 'Histórico completo' };
  document.getElementById('pageTitle').textContent = titles[key] || 'Pelada Musical';
}

window.addEventListener('hashchange', () => setActiveView(location.hash));

/* ---------------- upload manual ---------------- */

document.getElementById('file').addEventListener('change', (e) => {
  const f = e.target.files[0];
  if (!f) return;
  setStatus('loading', `Lendo ${f.name}...`);
  const reader = new FileReader();
  reader.onload = (evt) => {
    try {
      const wb = XLSX.read(evt.target.result, { type: 'array', cellDates: true });
      ingestWorkbook(wb);
    } catch (err) {
      setStatus('error', 'Não consegui ler esse arquivo.');
      console.error(err);
    }
  };
  reader.readAsArrayBuffer(f);
});

/* ---------------- exportar relatório do mês ---------------- */

function exportMonth() {
  const latest = {
    instagram: DATA.instagram[DATA.instagram.length - 1],
    youtube: DATA.youtube[DATA.youtube.length - 1],
    spotify: DATA.spotify[DATA.spotify.length - 1]
  };
  const anyMonth = latest.instagram || latest.youtube || latest.spotify;
  if (!anyMonth) { alert('Ainda não há dados para exportar.'); return; }

  document.getElementById('rGenerated').textContent = `Gerado em ${new Date().toLocaleDateString('pt-BR')}`;
  document.getElementById('rMonth').textContent = anyMonth.label;

  const grid = document.getElementById('rGrid');
  grid.innerHTML = Object.entries(PLATFORMS).map(([key, cfg]) => {
    const row = latest[key];
    return `<div><span>${cfg.label} — ${cfg.primaryLabel}</span><strong>${row ? nfInt.format(row.primary) : '—'}</strong></div>`;
  }).join('') + Object.entries(PLATFORMS).map(([key, cfg]) => {
    const row = latest[key];
    return `<div><span>${cfg.label} — ${cfg.totalLabel}</span><strong>${row ? nfInt.format(row.total) : '—'}</strong></div>`;
  }).join('');

  const text = document.getElementById('rText');
  text.innerHTML = Object.entries(PLATFORMS).map(([key, cfg]) => {
    const row = latest[key];
    if (!row) return `<h3>${cfg.label}</h3><p>Sem dados neste mês.</p>`;
    const trend = row.variation === null ? '' : (row.variation >= 0 ? ` (alta de ${nfPct(row.variation)})` : ` (queda de ${nfPct(row.variation)})`);
    return `<h3>${cfg.label}</h3><p>${nfInt.format(row.primary)} ${cfg.unit}${trend}. ${cfg.totalLabel}: ${nfInt.format(row.total)}. ${row.obs || ''}</p>`;
  }).join('');

  setTimeout(() => window.print(), 50);
}

/* ---------------- utils ---------------- */

function q(scope, sel) { return scope.querySelector(sel); }

/* ---------------- init ---------------- */

setActiveView(location.hash);
loadExcel(false);
setInterval(() => loadExcel(false), 5 * 60 * 1000);
