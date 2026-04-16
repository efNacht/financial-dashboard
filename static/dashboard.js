// ============================================================
// HELPERS
// ============================================================
// Clean, distinct palette for light background — max 5 on one chart
const COLORS5 = ['#4f46e5','#059669','#d97706','#dc2626','#7c3aed'];
// Extended palette for bar charts and cards
const COLORS = [
  '#4f46e5','#059669','#d97706','#dc2626','#7c3aed',
  '#0891b2','#c026d3','#ea580c','#2563eb','#65a30d',
  '#be123c','#0d9488','#9333ea','#b45309','#475569',
  '#6366f1','#10b981','#f59e0b','#ef4444'
];

// Light-theme chart defaults
const GRID_COLOR = 'rgba(0,0,0,0.06)';
const TICK_COLOR = '#6b7280';
const LEGEND_COLOR = '#374151';

function fmt(n, decimals=0) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  if (Math.abs(n) >= 1e6) return (n/1e6).toFixed(1) + ' млрд';
  if (Math.abs(n) >= 1e3) return (n/1e3).toFixed(1) + ' млн';
  return n.toFixed(decimals);
}

function fmtK(n) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  return new Intl.NumberFormat('ru-RU').format(Math.round(n));
}

function fmtPct(n) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  return n.toFixed(1) + '%';
}

// Get full display name: "Екатеринбург Яблоко (ЕЯ)"
function nameOf(c, k) {
  if (c.full_name && c.full_name !== k) return c.full_name;
  return k;
}
// Short label for chart axes: "Екатеринбург Яблоко"
function shortName(c, k) {
  if (c.full_name && c.full_name !== k) {
    // Remove legal form prefix
    return c.full_name.replace(/^(ООО|АО|ЗАО|ПАО)\s*["«]?/i, '').replace(/["»]$/,'');
  }
  return k;
}

function getYears() {
  const ys = new Set();
  Object.values(DATA).forEach(c => Object.keys(c.years_data).forEach(y => ys.add(+y)));
  return [...ys].sort();
}

function getLatestYear(company) {
  const years = Object.keys(company.years_data).map(Number).sort();
  return years[years.length - 1];
}

function getLatest(company) {
  const y = getLatestYear(company);
  return y ? company.years_data[y] : null;
}

function getAllCompanies() {
  return Object.entries(DATA).sort((a,b) => {
    const la = getLatest(a[1]), lb = getLatest(b[1]);
    return ((lb?.revenue)||0) - ((la?.revenue)||0);
  });
}

const chartInstances = {};
function destroyChart(id) {
  if (chartInstances[id]) { chartInstances[id].destroy(); delete chartInstances[id]; }
}
function makeChart(id, config) {
  destroyChart(id);
  const ctx = document.getElementById(id);
  if (!ctx) return null;
  const parent = ctx.parentElement;
  if (parent && !parent.querySelector('.chart-container')) {
    const wrapper = document.createElement('div');
    wrapper.className = 'chart-container';
    wrapper.style.position = 'relative';
    wrapper.style.height = parent.classList.contains('full') ? '380px' : '320px';
    parent.insertBefore(wrapper, ctx);
    wrapper.appendChild(ctx);
  }
  config.options = config.options || {};
  config.options.responsive = true;
  config.options.maintainAspectRatio = false;
  config.options.plugins = config.options.plugins || {};
  config.options.plugins.legend = config.options.plugins.legend || { labels: { color: LEGEND_COLOR, font: { size: 11 } } };
  if (!config.options.scales) {
    if (config.type === 'bar' || config.type === 'line') {
      config.options.scales = {
        x: { ticks: { color: TICK_COLOR, font: { size: 11 } }, grid: { color: GRID_COLOR } },
        y: { ticks: { color: TICK_COLOR, font: { size: 11 } }, grid: { color: GRID_COLOR } }
      };
    }
  }
  const chart = new Chart(ctx, config);
  chartInstances[id] = chart;
  return chart;
}

// Shared scale options builder
function scaleOpts(opts) {
  return {
    ticks: { color: TICK_COLOR, font: { size: 11 }, ...opts?.ticks },
    grid: { color: GRID_COLOR, ...opts?.grid },
    ...opts?.extra
  };
}

// ============================================================
// OVERVIEW SECTION
// ============================================================
function renderOverview() {
  const years = getYears();
  const companies = getAllCompanies();
  const latestYear = years[years.length - 1];

  let totalRev = 0, totalProfit = 0, avgMargin = 0, cnt = 0;
  companies.forEach(([k, c]) => {
    const d = c.years_data[latestYear];
    if (d) {
      totalRev += d.revenue || 0;
      totalProfit += d.net_profit || 0;
      if (d.gross_margin !== null) { avgMargin += d.gross_margin; cnt++; }
    }
  });
  avgMargin = cnt ? avgMargin / cnt : 0;

  let prevRev = 0, prevProfit = 0;
  const prevYear = years.length > 1 ? years[years.length - 2] : null;
  if (prevYear) {
    companies.forEach(([k, c]) => {
      const d = c.years_data[prevYear];
      if (d) { prevRev += d.revenue || 0; prevProfit += d.net_profit || 0; }
    });
  }

  const revDelta = prevRev ? ((totalRev - prevRev) / prevRev * 100).toFixed(1) : null;
  const profDelta = prevProfit ? ((totalProfit - prevProfit) / Math.abs(prevProfit) * 100).toFixed(1) : null;

  document.getElementById('overview-cards').innerHTML = `
    <div class="stat-card purple">
      <div class="label">Компаний в анализе</div>
      <div class="value">${companies.length}</div>
      <div class="delta" style="color:var(--text-secondary)">${years.length} лет данных (${years[0]}–${latestYear})</div>
    </div>
    <div class="stat-card green">
      <div class="label">Совокупная выручка ${latestYear}</div>
      <div class="value">${fmt(totalRev)}</div>
      <div class="delta ${revDelta > 0 ? 'positive' : 'negative'}">${revDelta ? (revDelta > 0 ? '+' : '') + revDelta + '% к ' + prevYear : ''}</div>
    </div>
    <div class="stat-card orange">
      <div class="label">Совокупная чистая прибыль ${latestYear}</div>
      <div class="value">${fmt(totalProfit)}</div>
      <div class="delta ${profDelta > 0 ? 'positive' : 'negative'}">${profDelta ? (profDelta > 0 ? '+' : '') + profDelta + '% к ' + prevYear : ''}</div>
    </div>
    <div class="stat-card blue">
      <div class="label">Средняя валовая маржа ${latestYear}</div>
      <div class="value">${avgMargin.toFixed(1)}%</div>
      <div class="delta" style="color:var(--text-secondary)">по ${cnt} компаниям</div>
    </div>
  `;

  // Stacked bar — top 8 companies + others
  const top8 = companies.slice(0, 8);
  const dsStacked = top8.map(([k, c], i) => ({
    label: shortName(c, k),
    data: years.map(y => c.years_data[y]?.revenue || 0),
    backgroundColor: COLORS[i],
    borderRadius: 2,
  }));
  // "Others" combined
  const othersDS = {
    label: 'Остальные',
    data: years.map(y => {
      let s = 0;
      companies.slice(8).forEach(([k,c]) => s += c.years_data[y]?.revenue || 0);
      return s;
    }),
    backgroundColor: '#d1d5db',
    borderRadius: 2,
  };
  dsStacked.push(othersDS);
  makeChart('chart-revenue-all', {
    type: 'bar',
    data: { labels: years, datasets: dsStacked },
    options: {
      plugins: { legend: { position: 'bottom', labels: { color: LEGEND_COLOR, font: { size: 11 }, boxWidth: 12, padding: 10 } } },
      scales: {
        x: { stacked: true, ticks: { color: TICK_COLOR }, grid: { color: GRID_COLOR } },
        y: { stacked: true, ticks: { color: TICK_COLOR, callback: v => fmt(v) }, grid: { color: GRID_COLOR } }
      }
    }
  });

  // Top revenue horizontal bar
  const topRev = companies.slice(0, 10).map(([k, c]) => ({ name: shortName(c, k), val: getLatest(c)?.revenue || 0 }));
  makeChart('chart-top-revenue', {
    type: 'bar',
    data: {
      labels: topRev.map(r => r.name),
      datasets: [{ data: topRev.map(r => r.val), backgroundColor: COLORS.slice(0, 10), borderRadius: 4 }]
    },
    options: {
      indexAxis: 'y',
      plugins: { legend: { display: false } },
      scales: {
        x: { ticks: { color: TICK_COLOR, callback: v => fmt(v) }, grid: { color: GRID_COLOR } },
        y: { ticks: { color: '#1a1d2e', font: { size: 12 } }, grid: { display: false } }
      }
    }
  });

  // Top profit
  const topProf = [...companies].sort((a,b) => (getLatest(b[1])?.net_profit||0) - (getLatest(a[1])?.net_profit||0)).slice(0, 10);
  makeChart('chart-top-profit', {
    type: 'bar',
    data: {
      labels: topProf.map(([k,c]) => shortName(c,k)),
      datasets: [{
        data: topProf.map(([k,c]) => getLatest(c)?.net_profit || 0),
        backgroundColor: topProf.map(([k,c]) => (getLatest(c)?.net_profit||0) >= 0 ? '#059669' : '#dc2626'),
        borderRadius: 4,
      }]
    },
    options: {
      indexAxis: 'y',
      plugins: { legend: { display: false } },
      scales: {
        x: { ticks: { color: TICK_COLOR, callback: v => fmt(v) }, grid: { color: GRID_COLOR } },
        y: { ticks: { color: '#1a1d2e', font: { size: 12 } }, grid: { display: false } }
      }
    }
  });

  // Market share doughnut
  const shareData = companies.filter(([k,c]) => (getLatest(c)?.revenue||0) > 0).slice(0, 8);
  const othersRev = totalRev - shareData.reduce((s,[k,c]) => s + (getLatest(c)?.revenue||0), 0);
  makeChart('chart-market-share', {
    type: 'doughnut',
    data: {
      labels: [...shareData.map(([k,c]) => shortName(c,k)), ...(othersRev > 0 ? ['Прочие'] : [])],
      datasets: [{
        data: [...shareData.map(([k,c]) => getLatest(c)?.revenue||0), ...(othersRev > 0 ? [othersRev] : [])],
        backgroundColor: [...COLORS.slice(0, shareData.length), '#d1d5db'],
        borderWidth: 2,
        borderColor: '#fff',
      }]
    },
    options: {
      plugins: {
        legend: { position: 'right', labels: { color: LEGEND_COLOR, font: { size: 11 }, boxWidth: 12, padding: 8 } },
      }
    }
  });

  // Total revenue trend
  const totalByYear = years.map(y => {
    let sum = 0;
    companies.forEach(([k,c]) => { sum += c.years_data[y]?.revenue || 0; });
    return sum;
  });
  makeChart('chart-total-revenue', {
    type: 'line',
    data: {
      labels: years,
      datasets: [{
        label: 'Совокупная выручка',
        data: totalByYear,
        borderColor: '#4f46e5',
        backgroundColor: 'rgba(79,70,229,0.08)',
        fill: true,
        tension: 0.3,
        pointRadius: 5,
        pointBackgroundColor: '#4f46e5',
        borderWidth: 2.5,
      }]
    },
    options: {
      plugins: { legend: { display: false } },
      scales: {
        y: { ticks: { color: TICK_COLOR, callback: v => fmt(v) }, grid: { color: GRID_COLOR } },
        x: { ticks: { color: TICK_COLOR }, grid: { color: GRID_COLOR } }
      }
    }
  });

  // ---- Summary table: Revenue & Profit 2023-2025 ----
  let sh = '<thead><tr><th>Юр. лицо</th><th>Выручка 2023</th><th>Выручка 2024</th><th>Выручка 2025</th><th>Чистая прибыль 2023</th><th>Чистая прибыль 2024</th><th>Чистая прибыль 2025</th><th>Приб. 24/23, %</th><th>Приб. 25/24, %</th><th>Приб. 25–24, руб.</th><th>Приб. за 2 года</th></tr></thead><tbody>';
  companies.forEach(([k, c]) => {
    const d23 = c.years_data[2023], d24 = c.years_data[2024], d25 = c.years_data[2025];
    const r23 = d23?.revenue || 0, r24 = d24?.revenue || 0, r25 = d25?.revenue || 0;
    const p23 = d23?.net_profit || 0, p24 = d24?.net_profit || 0, p25 = d25?.net_profit || 0;
    const g2423 = p23 ? ((p24 - p23) / Math.abs(p23) * 100) : null;
    const g2524 = p24 ? ((p25 - p24) / Math.abs(p24) * 100) : null;
    const diff = p25 - p24;
    const sum2y = p24 + p25;
    sh += `<tr><td>${nameOf(c, k)}</td>`;
    sh += `<td>${r23 ? fmtK(r23) : '—'}</td><td>${r24 ? fmtK(r24) : '—'}</td><td>${r25 ? fmtK(r25) : '—'}</td>`;
    sh += `<td class="${p23>=0?'positive':'negative'}">${p23 ? fmtK(p23) : '—'}</td>`;
    sh += `<td class="${p24>=0?'positive':'negative'}">${p24 ? fmtK(p24) : '—'}</td>`;
    sh += `<td class="${p25>=0?'positive':'negative'}">${p25 ? fmtK(p25) : '—'}</td>`;
    sh += `<td class="${(g2423||0)>=0?'positive':'negative'}">${g2423 !== null ? g2423.toFixed(0) + '%' : '—'}</td>`;
    sh += `<td class="${(g2524||0)>=0?'positive':'negative'}">${g2524 !== null ? g2524.toFixed(0) + '%' : '—'}</td>`;
    sh += `<td class="${diff>=0?'positive':'negative'}">${fmtK(diff)}</td>`;
    sh += `<td class="${sum2y>=0?'positive':'negative'}">${fmtK(sum2y)}</td>`;
    sh += '</tr>';
  });
  sh += '</tbody>';
  document.getElementById('table-summary').innerHTML = sh;
}

// ============================================================
// REVENUE SECTION
// ============================================================
function renderRevenue() {
  const years = getYears();
  const companies = getAllCompanies();

  // Grouped bar chart — top 8 companies, bars per year
  const top8 = companies.slice(0, 8);
  makeChart('chart-rev-dynamics', {
    type: 'bar',
    data: {
      labels: years,
      datasets: top8.map(([k, c], i) => ({
        label: shortName(c, k),
        data: years.map(y => c.years_data[y]?.revenue || 0),
        backgroundColor: COLORS[i % COLORS.length] + 'cc',
        borderRadius: 3,
      }))
    },
    options: {
      plugins: { legend: { position: 'bottom', labels: { color: LEGEND_COLOR, font: { size: 11 }, boxWidth: 12, padding: 10 } } },
      scales: {
        x: { ticks: { color: TICK_COLOR }, grid: { display: false } },
        y: { ticks: { color: TICK_COLOR, callback: v => fmt(v) }, grid: { color: GRID_COLOR } }
      }
    }
  });

  // Growth — horizontal bar, all companies, latest year vs prior
  const ly = years[years.length - 1];
  const py = years[years.length - 2];
  const growthData = companies.filter(([k,c]) => c.years_data[ly]?.revenue && c.years_data[py]?.revenue)
    .map(([k,c]) => ({ name: shortName(c,k), growth: ((c.years_data[ly].revenue - c.years_data[py].revenue) / c.years_data[py].revenue * 100) }))
    .sort((a,b) => b.growth - a.growth);
  makeChart('chart-rev-growth', {
    type: 'bar',
    data: {
      labels: growthData.map(d => d.name),
      datasets: [{ data: growthData.map(d => d.growth), backgroundColor: growthData.map(d => d.growth >= 0 ? '#059669cc' : '#dc2626cc'), borderRadius: 4 }]
    },
    options: {
      indexAxis: 'y',
      plugins: { legend: { display: false } },
      scales: {
        x: { ticks: { color: TICK_COLOR, callback: v => v + '%' }, grid: { color: GRID_COLOR } },
        y: { ticks: { color: '#1a1d2e', font: { size: 11 } }, grid: { display: false } }
      }
    }
  });

  // Cost structure (top 5)
  makeChart('chart-cost-structure', {
    type: 'bar',
    data: {
      labels: top5.map(([k,c]) => shortName(c,k)),
      datasets: [
        { label: 'Себестоимость', data: top5.map(([k,c]) => getLatest(c)?.cogs || 0), backgroundColor: '#ef4444', borderRadius: 3 },
        { label: 'Коммерческие', data: top5.map(([k,c]) => getLatest(c)?.commercial_expenses || 0), backgroundColor: '#f59e0b', borderRadius: 3 },
        { label: 'Управленческие', data: top5.map(([k,c]) => getLatest(c)?.admin_expenses || 0), backgroundColor: '#3b82f6', borderRadius: 3 },
      ]
    },
    options: {
      plugins: { legend: { position: 'bottom', labels: { color: LEGEND_COLOR, font: { size: 11 }, boxWidth: 12 } } },
      scales: {
        x: { stacked: true, ticks: { color: TICK_COLOR }, grid: { display: false } },
        y: { stacked: true, ticks: { color: TICK_COLOR, callback: v => fmt(v) }, grid: { color: GRID_COLOR } }
      }
    }
  });

  // Revenue table — full names
  let html = '<thead><tr><th>Компания</th>';
  years.forEach(y => html += `<th>${y}</th>`);
  html += '<th>CAGR</th></tr></thead><tbody>';
  companies.forEach(([k, c]) => {
    html += `<tr><td>${nameOf(c, k)}</td>`;
    years.forEach(y => {
      const v = c.years_data[y]?.revenue;
      html += `<td>${v ? fmtK(v) : '—'}</td>`;
    });
    const firstY = years.find(y => c.years_data[y]?.revenue > 0);
    const lastY = [...years].reverse().find(y => c.years_data[y]?.revenue > 0);
    if (firstY && lastY && firstY !== lastY) {
      const n = lastY - firstY;
      const cagr = (Math.pow(c.years_data[lastY].revenue / c.years_data[firstY].revenue, 1/n) - 1) * 100;
      html += `<td class="${cagr >= 0 ? 'positive' : 'negative'}">${cagr.toFixed(1)}%</td>`;
    } else html += '<td>—</td>';
    html += '</tr>';
  });
  html += '</tbody>';
  document.getElementById('table-revenue').innerHTML = html;
}

// ============================================================
// PROFITABILITY SECTION
// ============================================================
function renderProfitability() {
  const years = getYears();
  const companies = getAllCompanies();

  // Horizontal bar charts — all companies ranked by margin
  ['gross_margin', 'operating_margin', 'net_margin'].forEach((metric, idx) => {
    const ids = ['chart-gross-margin', 'chart-op-margin', 'chart-net-margin'];
    const accentColors = ['#4f46e5', '#059669', '#7c3aed'];
    const data = companies.filter(([k,c]) => getLatest(c)?.[metric] != null)
      .map(([k,c]) => ({ name: shortName(c,k), val: getLatest(c)[metric] }))
      .sort((a,b) => b.val - a.val);
    makeChart(ids[idx], {
      type: 'bar',
      data: {
        labels: data.map(d => d.name),
        datasets: [{ data: data.map(d => d.val), backgroundColor: data.map(d => d.val >= 0 ? accentColors[idx] + 'cc' : '#dc2626cc'), borderRadius: 4 }]
      },
      options: {
        indexAxis: 'y',
        plugins: { legend: { display: false } },
        scales: {
          x: { ticks: { color: TICK_COLOR, callback: v => v + '%' }, grid: { color: GRID_COLOR } },
          y: { ticks: { color: '#1a1d2e', font: { size: 11 } }, grid: { display: false } }
        }
      }
    });
  });

  // Net profit grouped bars — top 8
  const top8 = companies.slice(0, 8);
  makeChart('chart-net-profit-dyn', {
    type: 'bar',
    data: {
      labels: years,
      datasets: top8.map(([k,c], i) => ({
        label: shortName(c, k),
        data: years.map(y => c.years_data[y]?.net_profit || 0),
        backgroundColor: COLORS[i % COLORS.length] + 'cc',
        borderRadius: 3,
      }))
    },
    options: {
      plugins: { legend: { position: 'bottom', labels: { color: LEGEND_COLOR, font: { size: 11 }, boxWidth: 12 } } },
      scales: {
        x: { ticks: { color: TICK_COLOR }, grid: { display: false } },
        y: { ticks: { color: TICK_COLOR, callback: v => fmt(v) }, grid: { color: GRID_COLOR } }
      }
    }
  });

  // Waterfall for #1 company
  const [topK, topC] = companies[0];
  const topD = getLatest(topC);
  if (topD) {
    const wfLabels = ['Выручка', 'Себестоимость', 'Валовая приб.', 'Комм. расх.', 'Упр. расх.', 'Операц. приб.', 'Прочее', 'Чистая приб.'];
    const wfData = [topD.revenue, -topD.cogs, topD.gross_profit, -topD.commercial_expenses, -topD.admin_expenses, topD.operating_profit,
      (topD.interest_income - topD.interest_expense + topD.other_income - topD.other_expense), topD.net_profit];
    makeChart('chart-waterfall', {
      type: 'bar',
      data: {
        labels: wfLabels,
        datasets: [{
          data: wfData,
          backgroundColor: wfData.map(v => v >= 0 ? '#059669' : '#dc2626'),
          borderRadius: 4,
        }]
      },
      options: {
        plugins: { legend: { display: false },
          title: { display: true, text: nameOf(topC, topK), color: '#1a1d2e', font: { size: 13, weight: 600 } }
        },
        scales: {
          x: { ticks: { color: TICK_COLOR, font: { size: 10 } }, grid: { display: false } },
          y: { ticks: { color: TICK_COLOR, callback: v => fmt(v) }, grid: { color: GRID_COLOR } }
        }
      }
    });
  }

  // Margins table — full names
  let html = '<thead><tr><th>Компания</th><th>Год</th><th>Валовая</th><th>Операционная</th><th>Чистая</th><th>ROA</th><th>ROE</th></tr></thead><tbody>';
  companies.forEach(([k, c]) => {
    const ly = getLatestYear(c);
    const d = c.years_data[ly];
    if (d) {
      html += `<tr><td>${nameOf(c, k)}</td><td>${ly}</td>`;
      html += `<td class="${(d.gross_margin||0)>=0?'positive':'negative'}">${fmtPct(d.gross_margin)}</td>`;
      html += `<td class="${(d.operating_margin||0)>=0?'positive':'negative'}">${fmtPct(d.operating_margin)}</td>`;
      html += `<td class="${(d.net_margin||0)>=0?'positive':'negative'}">${fmtPct(d.net_margin)}</td>`;
      html += `<td class="${(d.roa||0)>=0?'positive':'negative'}">${fmtPct(d.roa)}</td>`;
      html += `<td class="${(d.roe||0)>=0?'positive':'negative'}">${fmtPct(d.roe)}</td>`;
      html += '</tr>';
    }
  });
  html += '</tbody>';
  document.getElementById('table-margins').innerHTML = html;
}

// ============================================================
// BALANCE SECTION
// ============================================================
function renderBalance() {
  const years = getYears();
  const companies = getAllCompanies();

  // Assets horizontal bar — all companies
  const assetsData = companies.filter(([k,c]) => (getLatest(c)?.total_assets||0) > 0);
  makeChart('chart-assets-bar', {
    type: 'bar',
    data: {
      labels: assetsData.map(([k,c]) => shortName(c,k)),
      datasets: [{ data: assetsData.map(([k,c]) => getLatest(c)?.total_assets||0), backgroundColor: COLORS, borderRadius: 4 }]
    },
    options: {
      indexAxis: 'y',
      plugins: { legend: { display: false } },
      scales: {
        x: { ticks: { color: TICK_COLOR, callback: v => fmt(v) }, grid: { color: GRID_COLOR } },
        y: { ticks: { color: '#1a1d2e', font: { size: 11 } }, grid: { display: false } }
      }
    }
  });

  // Asset structure (top 5)
  const top5 = companies.slice(0, 5);
  makeChart('chart-asset-structure', {
    type: 'bar',
    data: {
      labels: top5.map(([k,c]) => shortName(c,k)),
      datasets: [
        { label: 'Основные средства', data: top5.map(([k,c]) => getLatest(c)?.fixed_assets||0), backgroundColor: '#4f46e5', borderRadius: 2 },
        { label: 'Запасы', data: top5.map(([k,c]) => getLatest(c)?.inventories||0), backgroundColor: '#059669', borderRadius: 2 },
        { label: 'Дебиторская задолженность', data: top5.map(([k,c]) => getLatest(c)?.receivables||0), backgroundColor: '#d97706', borderRadius: 2 },
        { label: 'Денежные средства', data: top5.map(([k,c]) => getLatest(c)?.cash||0), backgroundColor: '#2563eb', borderRadius: 2 },
      ]
    },
    options: {
      plugins: { legend: { position: 'bottom', labels: { color: LEGEND_COLOR, font: { size: 11 }, boxWidth: 12 } } },
      scales: {
        x: { stacked: true, ticks: { color: TICK_COLOR }, grid: { display: false } },
        y: { stacked: true, ticks: { color: TICK_COLOR, callback: v => fmt(v) }, grid: { color: GRID_COLOR } }
      }
    }
  });

  // Equity vs Debt — top 8
  const top8 = companies.slice(0, 8);
  makeChart('chart-equity-debt', {
    type: 'bar',
    data: {
      labels: top8.map(([k,c]) => shortName(c,k)),
      datasets: [
        { label: 'Собственный капитал', data: top8.map(([k,c]) => getLatest(c)?.equity||0), backgroundColor: '#059669', borderRadius: 2 },
        { label: 'Долгосрочный долг', data: top8.map(([k,c]) => getLatest(c)?.lt_debt||0), backgroundColor: '#d97706', borderRadius: 2 },
        { label: 'Краткосрочный долг', data: top8.map(([k,c]) => getLatest(c)?.st_debt||0), backgroundColor: '#dc2626', borderRadius: 2 },
      ]
    },
    options: {
      plugins: { legend: { position: 'bottom', labels: { color: LEGEND_COLOR, font: { size: 11 }, boxWidth: 12 } } },
      scales: {
        x: { ticks: { color: TICK_COLOR, font: { size: 10 } }, grid: { display: false } },
        y: { ticks: { color: TICK_COLOR, callback: v => fmt(v) }, grid: { color: GRID_COLOR } }
      }
    }
  });

  // Cash — horizontal bar, top 8
  const top8cash = companies.slice(0, 8);
  const cashData = top8cash.filter(([k,c]) => (getLatest(c)?.cash||0) > 0)
    .map(([k,c]) => ({ name: shortName(c,k), val: getLatest(c)?.cash||0 }))
    .sort((a,b) => b.val - a.val);
  makeChart('chart-cash', {
    type: 'bar',
    data: {
      labels: cashData.map(d => d.name),
      datasets: [{ data: cashData.map(d => d.val), backgroundColor: '#2563ebcc', borderRadius: 4 }]
    },
    options: {
      indexAxis: 'y',
      plugins: { legend: { display: false } },
      scales: {
        x: { ticks: { color: TICK_COLOR, callback: v => fmt(v) }, grid: { color: GRID_COLOR } },
        y: { ticks: { color: '#1a1d2e', font: { size: 11 } }, grid: { display: false } }
      }
    }
  });

  // Balance table
  let html = '<thead><tr><th>Компания</th><th>Год</th><th>Активы</th><th>Внеоборотные</th><th>Оборотные</th><th>Капитал</th><th>Долгоср. долг</th><th>Краткоср. долг</th><th>Кредиторка</th></tr></thead><tbody>';
  companies.forEach(([k,c]) => {
    const ly = getLatestYear(c);
    const d = c.years_data[ly];
    if (d) {
      html += `<tr><td>${nameOf(c, k)}</td><td>${ly}</td><td>${fmtK(d.total_assets)}</td><td>${fmtK(d.non_current_assets)}</td><td>${fmtK(d.current_assets)}</td><td>${fmtK(d.equity)}</td><td>${fmtK(d.lt_debt)}</td><td>${fmtK(d.st_debt)}</td><td>${fmtK(d.payables)}</td></tr>`;
    }
  });
  html += '</tbody>';
  document.getElementById('table-balance').innerHTML = html;
}

// ============================================================
// RATIOS SECTION
// ============================================================
function renderRatios() {
  const years = getYears();
  const companies = getAllCompanies();

  const metrics = ['roa', 'roe', 'current_ratio', 'debt_equity'];
  const ids = ['chart-roa', 'chart-roe', 'chart-current-ratio', 'chart-de-ratio'];
  const suffixes = ['%', '%', '', ''];
  const barColors = ['#4f46e5', '#059669', '#d97706', '#7c3aed'];

  metrics.forEach((metric, idx) => {
    const data = companies.filter(([k,c]) => {
      const v = getLatest(c)?.[metric];
      return v !== null && v !== undefined && isFinite(v) && Math.abs(v) < 500;
    }).map(([k,c]) => ({ name: shortName(c,k), val: getLatest(c)[metric] }))
      .sort((a,b) => b.val - a.val);
    makeChart(ids[idx], {
      type: 'bar',
      data: {
        labels: data.map(d => d.name),
        datasets: [{ data: data.map(d => d.val), backgroundColor: data.map(d => d.val >= 0 ? barColors[idx] + 'cc' : '#dc2626cc'), borderRadius: 4 }]
      },
      options: {
        indexAxis: 'y',
        plugins: { legend: { display: false } },
        scales: {
          x: { ticks: { color: TICK_COLOR, callback: v => v + suffixes[idx] }, grid: { color: GRID_COLOR } },
          y: { ticks: { color: '#1a1d2e', font: { size: 11 } }, grid: { display: false } }
        }
      }
    });
  });

  // Ratios table
  const allC = getAllCompanies();
  let html = '<thead><tr><th>Компания</th><th>Год</th><th>Валовая маржа</th><th>Опер. маржа</th><th>Чистая маржа</th><th>ROA</th><th>ROE</th><th>Тек. ликвидность</th><th>Долг/Капитал</th></tr></thead><tbody>';
  allC.forEach(([k,c]) => {
    const ly = getLatestYear(c);
    const d = c.years_data[ly];
    if (d) {
      html += `<tr><td>${nameOf(c, k)}</td><td>${ly}</td><td>${fmtPct(d.gross_margin)}</td><td>${fmtPct(d.operating_margin)}</td><td>${fmtPct(d.net_margin)}</td><td>${fmtPct(d.roa)}</td><td>${fmtPct(d.roe)}</td><td>${d.current_ratio !== null ? d.current_ratio.toFixed(2) : '—'}</td><td>${d.debt_equity !== null ? d.debt_equity.toFixed(2) : '—'}</td></tr>`;
    }
  });
  html += '</tbody>';
  document.getElementById('table-ratios').innerHTML = html;
}

// ============================================================
// COMPARE SECTION
// ============================================================
function renderCompare() {
  const years = getYears();
  const companies = getAllCompanies();
  const sel1 = document.getElementById('cmp-company1');
  const sel2 = document.getElementById('cmp-company2');

  if (!sel1.options.length) {
    companies.forEach(([k,c]) => {
      sel1.add(new Option(nameOf(c, k), k));
      sel2.add(new Option(nameOf(c, k), k));
    });
    if (companies.length >= 2) sel2.selectedIndex = 1;
    sel1.onchange = sel2.onchange = document.getElementById('cmp-metric').onchange = updateCompare;
  }
  updateCompare();

  // Bubble chart
  const bubbleData = companies.filter(([k,c]) => getLatest(c)?.revenue > 0).map(([k,c], i) => {
    const d = getLatest(c);
    return {
      label: shortName(c, k),
      data: [{ x: d.revenue || 0, y: d.net_margin || 0, r: Math.max(4, Math.sqrt((d.total_assets||0) / 50000)) }],
      backgroundColor: COLORS[i % COLORS.length] + '88',
      borderColor: COLORS[i % COLORS.length],
      borderWidth: 1.5,
    };
  });
  makeChart('chart-bubble', {
    type: 'bubble',
    data: { datasets: bubbleData },
    options: {
      plugins: { legend: { position: 'bottom', labels: { color: LEGEND_COLOR, font: { size: 10 }, boxWidth: 10 } } },
      scales: {
        x: { title: { display: true, text: 'Выручка', color: TICK_COLOR }, ticks: { color: TICK_COLOR, callback: v => fmt(v) }, grid: { color: GRID_COLOR } },
        y: { title: { display: true, text: 'Чистая рентабельность (%)', color: TICK_COLOR }, ticks: { color: TICK_COLOR, callback: v => v + '%' }, grid: { color: GRID_COLOR } }
      }
    }
  });
}

function updateCompare() {
  const years = getYears();
  const k1 = document.getElementById('cmp-company1').value;
  const k2 = document.getElementById('cmp-company2').value;
  const metric = document.getElementById('cmp-metric').value;
  const c1 = DATA[k1], c2 = DATA[k2];
  if (!c1 || !c2) return;

  const metricNames = {revenue:'Выручка',net_profit:'Чистая прибыль',gross_margin:'Валовая рентабельность',operating_margin:'Операц. рентабельность',net_margin:'Чистая рентабельность',total_assets:'Активы',roa:'ROA',roe:'ROE'};
  document.getElementById('cmp-title').textContent = metricNames[metric] + ': ' + shortName(c1, k1) + ' vs ' + shortName(c2, k2);

  makeChart('chart-compare', {
    type: 'line',
    data: {
      labels: years,
      datasets: [
        { label: shortName(c1, k1), data: years.map(y => c1.years_data[y]?.[metric] ?? null), borderColor: '#4f46e5', backgroundColor: 'rgba(79,70,229,0.08)', fill: true, tension: 0.3, pointRadius: 5, borderWidth: 2.5 },
        { label: shortName(c2, k2), data: years.map(y => c2.years_data[y]?.[metric] ?? null), borderColor: '#059669', backgroundColor: 'rgba(5,150,105,0.08)', fill: true, tension: 0.3, pointRadius: 5, borderWidth: 2.5 },
      ]
    },
    options: {
      plugins: { legend: { labels: { color: LEGEND_COLOR, font: { size: 12 } } } },
      scales: {
        y: { ticks: { color: TICK_COLOR, callback: v => metric.includes('margin') || metric === 'roa' || metric === 'roe' ? v + '%' : fmt(v) }, grid: { color: GRID_COLOR } },
        x: { ticks: { color: TICK_COLOR }, grid: { color: GRID_COLOR } }
      }
    }
  });

  // Radar
  const d1 = getLatest(c1), d2 = getLatest(c2);
  if (d1 && d2) {
    const radarMetrics = ['gross_margin', 'operating_margin', 'net_margin', 'roa', 'roe'];
    const radarLabels = ['Вал. маржа', 'Опер. маржа', 'Чист. маржа', 'ROA', 'ROE'];
    makeChart('chart-radar', {
      type: 'radar',
      data: {
        labels: radarLabels,
        datasets: [
          { label: shortName(c1, k1), data: radarMetrics.map(m => d1[m] ?? 0), borderColor: '#4f46e5', backgroundColor: 'rgba(79,70,229,0.12)', pointBackgroundColor: '#4f46e5', borderWidth: 2 },
          { label: shortName(c2, k2), data: radarMetrics.map(m => d2[m] ?? 0), borderColor: '#059669', backgroundColor: 'rgba(5,150,105,0.12)', pointBackgroundColor: '#059669', borderWidth: 2 },
        ]
      },
      options: {
        scales: {
          r: { angleLines: { color: 'rgba(0,0,0,0.08)' }, grid: { color: 'rgba(0,0,0,0.08)' }, ticks: { color: TICK_COLOR, backdropColor: 'transparent' }, pointLabels: { color: '#1a1d2e', font: { size: 11 } } }
        },
        plugins: { legend: { labels: { color: LEGEND_COLOR } } }
      }
    });
  }
}

// ============================================================
// COMPANIES GRID
// ============================================================
function renderCompaniesGrid() {
  const companies = getAllCompanies();
  let html = '';
  companies.forEach(([k, c]) => {
    const d = getLatest(c);
    const ly = getLatestYear(c);
    html += `
      <div class="company-card" onclick="showCompanyDetail('${k}')">
        <div class="name">${nameOf(c, k)}</div>
        <div class="inn">ИНН: ${c.inn || '—'} | ${ly}</div>
        <div class="mini-stats">
          <div class="mini-stat"><div class="lbl">Выручка</div><div class="val">${fmt(d?.revenue)}</div></div>
          <div class="mini-stat"><div class="lbl">Чистая прибыль</div><div class="val" style="color:${(d?.net_profit||0)>=0?'var(--green)':'var(--red)'}">${fmt(d?.net_profit)}</div></div>
          <div class="mini-stat"><div class="lbl">Активы</div><div class="val">${fmt(d?.total_assets)}</div></div>
          <div class="mini-stat"><div class="lbl">Вал. маржа</div><div class="val">${fmtPct(d?.gross_margin)}</div></div>
        </div>
      </div>
    `;
  });
  document.getElementById('company-cards-grid').innerHTML = html;
}

function showCompanyDetail(key) {
  document.getElementById('detail-company').value = key;
  activateSection('details');
}

// ============================================================
// DETAIL SECTION
// ============================================================
function renderDetail() {
  const sel = document.getElementById('detail-company');
  const companies = getAllCompanies();

  if (!sel.options.length) {
    companies.forEach(([k,c]) => sel.add(new Option(nameOf(c, k), k)));
    sel.onchange = renderDetail;
  }

  const k = sel.value;
  const c = DATA[k];
  if (!c) return;

  const years = Object.keys(c.years_data).map(Number).sort();
  const ly = years[years.length - 1];
  const d = c.years_data[ly];
  if (!d) return;

  const py = years.length > 1 ? years[years.length - 2] : null;
  const pd = py ? c.years_data[py] : null;
  const revGrowth = pd && pd.revenue ? ((d.revenue - pd.revenue) / pd.revenue * 100).toFixed(1) : null;

  document.getElementById('detail-cards').innerHTML = `
    <div class="stat-card purple"><div class="label">Выручка ${ly}</div><div class="value">${fmt(d.revenue)}</div>
      <div class="delta ${revGrowth > 0 ? 'positive' : 'negative'}">${revGrowth ? (revGrowth > 0 ? '+' : '') + revGrowth + '% к ' + py : ''}</div></div>
    <div class="stat-card green"><div class="label">Чистая прибыль</div><div class="value" style="color:${d.net_profit>=0?'var(--green)':'var(--red)'}">${fmt(d.net_profit)}</div></div>
    <div class="stat-card orange"><div class="label">Активы</div><div class="value">${fmt(d.total_assets)}</div></div>
    <div class="stat-card blue"><div class="label">Валовая маржа</div><div class="value">${fmtPct(d.gross_margin)}</div></div>
  `;

  // P&L dynamics
  makeChart('chart-detail-pl', {
    type: 'bar',
    data: {
      labels: years,
      datasets: [
        { label: 'Выручка', data: years.map(y => c.years_data[y]?.revenue||0), backgroundColor: '#4f46e5', borderRadius: 4 },
        { label: 'Валовая прибыль', data: years.map(y => c.years_data[y]?.gross_profit||0), backgroundColor: '#059669', borderRadius: 4 },
        { label: 'Чистая прибыль', data: years.map(y => c.years_data[y]?.net_profit||0), backgroundColor: '#d97706', borderRadius: 4 },
      ]
    },
    options: {
      plugins: { legend: { position: 'bottom', labels: { color: LEGEND_COLOR, font: { size: 11 }, boxWidth: 12 } } },
      scales: {
        x: { ticks: { color: TICK_COLOR }, grid: { display: false } },
        y: { ticks: { color: TICK_COLOR, callback: v => fmt(v) }, grid: { color: GRID_COLOR } }
      }
    }
  });

  // Balance structure
  makeChart('chart-detail-balance', {
    type: 'bar',
    data: {
      labels: years,
      datasets: [
        { label: 'Основные средства', data: years.map(y => c.years_data[y]?.fixed_assets||0), backgroundColor: '#4f46e5', borderRadius: 2 },
        { label: 'Запасы', data: years.map(y => c.years_data[y]?.inventories||0), backgroundColor: '#059669', borderRadius: 2 },
        { label: 'Дебиторская задолж.', data: years.map(y => c.years_data[y]?.receivables||0), backgroundColor: '#d97706', borderRadius: 2 },
        { label: 'Денежные средства', data: years.map(y => c.years_data[y]?.cash||0), backgroundColor: '#2563eb', borderRadius: 2 },
      ]
    },
    options: {
      plugins: { legend: { position: 'bottom', labels: { color: LEGEND_COLOR, font: { size: 11 }, boxWidth: 12 } } },
      scales: {
        x: { stacked: true, ticks: { color: TICK_COLOR }, grid: { display: false } },
        y: { stacked: true, ticks: { color: TICK_COLOR, callback: v => fmt(v) }, grid: { color: GRID_COLOR } }
      }
    }
  });

  // Margins
  makeChart('chart-detail-margins', {
    type: 'line',
    data: {
      labels: years,
      datasets: [
        { label: 'Валовая', data: years.map(y => c.years_data[y]?.gross_margin??null), borderColor: '#4f46e5', tension: 0.3, pointRadius: 5, borderWidth: 2.5, fill: false },
        { label: 'Операционная', data: years.map(y => c.years_data[y]?.operating_margin??null), borderColor: '#059669', tension: 0.3, pointRadius: 5, borderWidth: 2.5, fill: false },
        { label: 'Чистая', data: years.map(y => c.years_data[y]?.net_margin??null), borderColor: '#d97706', tension: 0.3, pointRadius: 5, borderWidth: 2.5, fill: false },
      ]
    },
    options: {
      plugins: { legend: { position: 'bottom', labels: { color: LEGEND_COLOR, font: { size: 11 }, boxWidth: 14 } } },
      scales: {
        y: { ticks: { color: TICK_COLOR, callback: v => v + '%' }, grid: { color: GRID_COLOR } },
        x: { ticks: { color: TICK_COLOR }, grid: { color: GRID_COLOR } }
      }
    }
  });

  // Cost structure
  makeChart('chart-detail-costs', {
    type: 'bar',
    data: {
      labels: years,
      datasets: [
        { label: 'Себестоимость', data: years.map(y => c.years_data[y]?.cogs||0), backgroundColor: '#ef4444', borderRadius: 2 },
        { label: 'Коммерческие', data: years.map(y => c.years_data[y]?.commercial_expenses||0), backgroundColor: '#f59e0b', borderRadius: 2 },
        { label: 'Управленческие', data: years.map(y => c.years_data[y]?.admin_expenses||0), backgroundColor: '#3b82f6', borderRadius: 2 },
      ]
    },
    options: {
      plugins: { legend: { position: 'bottom', labels: { color: LEGEND_COLOR, font: { size: 11 }, boxWidth: 12 } } },
      scales: {
        x: { stacked: true, ticks: { color: TICK_COLOR }, grid: { display: false } },
        y: { stacked: true, ticks: { color: TICK_COLOR, callback: v => fmt(v) }, grid: { color: GRID_COLOR } }
      }
    }
  });

  // Detail P&L table
  const plRows = [
    ['Выручка', 'revenue'], ['Себестоимость', 'cogs'], ['Валовая прибыль', 'gross_profit'],
    ['Коммерческие расходы', 'commercial_expenses'], ['Управленческие расходы', 'admin_expenses'],
    ['Прибыль от продаж', 'operating_profit'], ['Проценты к получению', 'interest_income'],
    ['Проценты к уплате', 'interest_expense'], ['Прочие доходы', 'other_income'],
    ['Прочие расходы', 'other_expense'], ['Прибыль до налогообложения', 'profit_before_tax'],
    ['Чистая прибыль', 'net_profit'],
  ];
  let plHtml = '<thead><tr><th>Показатель</th>' + years.map(y => `<th>${y}</th>`).join('') + '</tr></thead><tbody>';
  plRows.forEach(([label, key]) => {
    plHtml += `<tr><td>${label}</td>`;
    years.forEach(y => {
      const v = c.years_data[y]?.[key];
      plHtml += `<td>${v !== undefined ? fmtK(v) : '—'}</td>`;
    });
    plHtml += '</tr>';
  });
  plHtml += '</tbody>';
  document.getElementById('table-detail-pl').innerHTML = plHtml;

  // Detail balance table
  const balRows = [
    ['Активы всего', 'total_assets'], ['Внеоборотные', 'non_current_assets'], ['Основные средства', 'fixed_assets'],
    ['Оборотные', 'current_assets'], ['Запасы', 'inventories'], ['Дебиторская', 'receivables'],
    ['Денежные средства', 'cash'], ['Собственный капитал', 'equity'],
    ['Долгоср. долг', 'lt_debt'], ['Краткоср. долг', 'st_debt'], ['Кредиторская задолж.', 'payables'],
  ];
  let balHtml = '<thead><tr><th>Показатель</th>' + years.map(y => `<th>${y}</th>`).join('') + '</tr></thead><tbody>';
  balRows.forEach(([label, key]) => {
    balHtml += `<tr><td>${label}</td>`;
    years.forEach(y => {
      const v = c.years_data[y]?.[key];
      balHtml += `<td>${v !== undefined ? fmtK(v) : '—'}</td>`;
    });
    balHtml += '</tr>';
  });
  balHtml += '</tbody>';
  document.getElementById('table-detail-balance').innerHTML = balHtml;
}

// ============================================================
// INIT
// ============================================================
const rendered = new Set();
const renderMap = {
  overview: renderOverview,
  revenue: renderRevenue,
  profitability: renderProfitability,
  balance: renderBalance,
  ratios: renderRatios,
  compare: renderCompare,
  companies: renderCompaniesGrid,
  details: renderDetail,
};

function activateSection(name) {
  document.querySelectorAll('.nav-tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.section').forEach(s => s.classList.remove('active'));
  const tab = document.querySelector(`[data-section="${name}"]`);
  if (tab) tab.classList.add('active');
  const sec = document.getElementById('sec-' + name);
  if (sec) sec.classList.add('active');
  if (renderMap[name]) renderMap[name]();
  rendered.add(name);
  window.scrollTo(0, 0);
}

document.querySelectorAll('.nav-tab').forEach(tab => {
  tab.addEventListener('click', () => activateSection(tab.dataset.section));
});

activateSection('overview');

