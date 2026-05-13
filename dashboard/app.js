const data = window.REPORT_DASHBOARD_DATA;

const state = {
  sourceLabel: 'all',
  utmSource: 'all',
  utmMedium: 'all',
  utmCampaign: 'all',
  dateFrom: data?.filters?.minDate || '',
  dateTo: data?.filters?.maxDate || '',
  detailDate: 'latest',
  search: '',
};

const els = {
  siteHeader: document.querySelector('.site-header'),
  navLinks: [...document.querySelectorAll('.site-nav a')],
  tableLink: document.getElementById('table-link'),
  kcReportLink: document.getElementById('kc-report-link'),
  heroSources: document.getElementById('hero-sources'),
  heroCombinations: document.getElementById('hero-combinations'),
  heroRecords: document.getElementById('hero-records'),
  heroApproved: document.getElementById('hero-approved'),
  heroMeetings: document.getElementById('hero-meetings'),
  heroReservations: document.getElementById('hero-reservations'),
  heroClosed: document.getElementById('hero-closed'),
  sourceLabel: document.getElementById('filter-source-label'),
  utmSource: document.getElementById('filter-utm-source'),
  utmMedium: document.getElementById('filter-utm-medium'),
  utmCampaign: document.getElementById('filter-utm-campaign'),
  dateFrom: document.getElementById('filter-date-from'),
  dateTo: document.getElementById('filter-date-to'),
  search: document.getElementById('filter-search'),
  reset: document.getElementById('reset-filters'),
  activeFilters: document.getElementById('active-filters'),
  selectionSummary: document.getElementById('selection-summary'),
  sourceSummaryBody: document.getElementById('source-summary-body'),
  detailCaption: document.getElementById('detail-caption'),
  detailDate: document.getElementById('detail-date-select'),
  detailBody: document.getElementById('detail-body'),
  exportCsv: document.getElementById('export-csv'),
  kpiRecords: document.getElementById('kpi-records'),
  kpiApproved: document.getElementById('kpi-approved'),
  kpiMeetings: document.getElementById('kpi-meetings'),
  kpiReservations: document.getElementById('kpi-reservations'),
  kpiClosed: document.getElementById('kpi-closed'),
  kpiCr: document.getElementById('kpi-cr'),
};

const chartPalette = {
  cyan: '#2f8cff',
  emerald: '#5aa68f',
  amber: '#efbd55',
  coral: '#d66b62',
  lime: '#66a05e',
  violet: '#7b7af0',
  ink: '#161414',
};

const numberFormatter = new Intl.NumberFormat('ru-RU');
const percentFormatter = new Intl.NumberFormat('ru-RU', {
  style: 'percent',
  minimumFractionDigits: 2,
  maximumFractionDigits: 2,
});
const dateFormatter = new Intl.DateTimeFormat('ru-RU', {
  day: '2-digit',
  month: '2-digit',
  year: 'numeric',
});
const monthFormatter = new Intl.DateTimeFormat('ru-RU', {
  month: 'long',
  year: 'numeric',
});

let dailyChart;
let sourceChart;
let funnelChart;
let performanceChart;

function formatNumber(value) {
  return numberFormatter.format(Number(value || 0));
}

function formatPercent(value) {
  return percentFormatter.format(Number(value || 0));
}

function formatDate(value) {
  if (!value) return '—';
  return dateFormatter.format(new Date(`${value}T00:00:00`));
}

function formatMonth(value) {
  if (!value) return '—';
  const label = monthFormatter.format(new Date(`${value}-01T00:00:00`));
  return label.slice(0, 1).toUpperCase() + label.slice(1);
}

function csvEscape(value) {
  const text = String(value ?? '');
  if (/[;"\r\n]/.test(text)) return `"${text.replace(/"/g, '""')}"`;
  return text;
}

function setText(element, value) {
  if (element) element.textContent = value;
}

function populateSelect(select, options, allLabel) {
  select.innerHTML = '';
  const allOption = document.createElement('option');
  allOption.value = 'all';
  allOption.textContent = allLabel;
  select.append(allOption);

  for (const option of options) {
    const element = document.createElement('option');
    element.value = String(option);
    element.textContent = String(option);
    select.append(element);
  }
}

function normalizeSearch(value) {
  return String(value || '').trim().toLowerCase();
}

function uniqueDates(rows) {
  return [...new Set(rows.map((row) => row.uploadDate).filter(Boolean))].sort((a, b) => b.localeCompare(a));
}

function filteredRows() {
  const query = normalizeSearch(state.search);
  return data.baseRows.filter((row) => {
    if (state.sourceLabel !== 'all' && row.sourceLabel !== state.sourceLabel) return false;
    if (state.utmSource !== 'all' && row.utmSource !== state.utmSource) return false;
    if (state.utmMedium !== 'all' && row.utmMedium !== state.utmMedium) return false;
    if (state.utmCampaign !== 'all' && row.utmCampaign !== state.utmCampaign) return false;
    if (state.dateFrom && row.uploadDate < state.dateFrom) return false;
    if (state.dateTo && row.uploadDate > state.dateTo) return false;
    if (!query) return true;

    return [
      row.sourceLabel,
      row.utmSource,
      row.utmMedium,
      row.utmCampaign,
    ].some((field) => normalizeSearch(field).includes(query));
  });
}

function summarizeRows(rows) {
  return rows.reduce((acc, row) => {
    acc.records += Number(row.records || 0);
    acc.approvedMortgage += Number(row.approvedMortgage || 0);
    acc.meetingShow += Number(row.meetingShow || 0);
    acc.reservation += Number(row.reservation || 0);
    acc.closedDeals += Number(row.closedDeals || 0);
    return acc;
  }, {
    records: 0,
    approvedMortgage: 0,
    meetingShow: 0,
    reservation: 0,
    closedDeals: 0,
  });
}

function summarizeByDate(rows) {
  const groups = new Map();

  for (const row of rows) {
    const key = row.uploadDate;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(row);
  }

  return [...groups.entries()]
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([date, items]) => ({
      date,
      ...summarizeRows(items),
    }));
}

function summarizeBySource(rows) {
  const groups = new Map();

  for (const row of rows) {
    const key = row.sourceLabel || 'Без источника';
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(row);
  }

  return [...groups.entries()]
    .map(([source, items]) => ({
      source,
      ...summarizeRows(items),
    }))
    .sort((a, b) => b.records - a.records);
}

function summarizeSourceTable(rows) {
  const groups = new Map();

  for (const row of rows) {
    const period = formatMonth(row.month);
    const source = row.sourceLabel || 'Без источника';
    const key = `${period}__${source}`;

    if (!groups.has(key)) {
      groups.set(key, {
        month: row.month,
        period,
        source,
        records: 0,
        approvedMortgage: 0,
        meetingShow: 0,
        reservation: 0,
        closedDeals: 0,
      });
    }

    const group = groups.get(key);
    group.records += Number(row.records || 0);
    group.approvedMortgage += Number(row.approvedMortgage || 0);
    group.meetingShow += Number(row.meetingShow || 0);
    group.reservation += Number(row.reservation || 0);
    group.closedDeals += Number(row.closedDeals || 0);
  }

  return [...groups.values()].sort((a, b) => `${a.month}_${a.source}`.localeCompare(`${b.month}_${b.source}`));
}

function uniqueCombinations(rows) {
  return new Set(
    rows.map((row) => `${row.utmSource}__${row.utmMedium}__${row.utmCampaign}`)
  ).size;
}

function renderHero(rows) {
  const summary = summarizeRows(rows);
  const uniqueSources = new Set(rows.map((row) => row.sourceLabel).filter(Boolean)).size;

  setText(els.heroSources, formatNumber(uniqueSources));
  setText(els.heroCombinations, formatNumber(uniqueCombinations(rows)));
  setText(els.heroRecords, formatNumber(summary.records));
  setText(els.heroApproved, formatNumber(summary.approvedMortgage));
  setText(els.heroMeetings, formatNumber(summary.meetingShow));
  setText(els.heroReservations, formatNumber(summary.reservation));
  setText(els.heroClosed, formatNumber(summary.closedDeals));
}

function renderKpis(rows) {
  const summary = summarizeRows(rows);
  els.kpiRecords.textContent = formatNumber(summary.records);
  els.kpiApproved.textContent = formatNumber(summary.approvedMortgage);
  els.kpiMeetings.textContent = formatNumber(summary.meetingShow);
  els.kpiReservations.textContent = formatNumber(summary.reservation);
  els.kpiClosed.textContent = formatNumber(summary.closedDeals);
  els.kpiCr.textContent = formatPercent(summary.records > 0 ? summary.closedDeals / summary.records : 0);
}

function renderActiveState(rows) {
  const chips = [];
  if (state.sourceLabel !== 'all') chips.push(`Источник: ${state.sourceLabel}`);
  if (state.utmSource !== 'all') chips.push(`utm source: ${state.utmSource}`);
  if (state.utmMedium !== 'all') chips.push(`utm medium: ${state.utmMedium}`);
  if (state.utmCampaign !== 'all') chips.push(`utm campaign: ${state.utmCampaign}`);
  if (state.dateFrom) chips.push(`От: ${formatDate(state.dateFrom)}`);
  if (state.dateTo) chips.push(`До: ${formatDate(state.dateTo)}`);
  if (normalizeSearch(state.search)) chips.push(`Поиск: ${state.search.trim()}`);

  els.activeFilters.innerHTML = chips.length > 0
    ? chips.map((chip) => `<span class="chip">${chip}</span>`).join('')
    : '<span class="chip">Все данные</span>';

  const summary = summarizeRows(rows);
  const uniqueSources = new Set(rows.map((row) => row.sourceLabel).filter(Boolean)).size;
  els.selectionSummary.textContent = `Строк: ${formatNumber(rows.length)} · Источников: ${formatNumber(uniqueSources)} · Лидов: ${formatNumber(summary.records)}`;
}

function populateDetailDateSelect(rows) {
  const dates = uniqueDates(rows);
  const previous = state.detailDate;
  els.detailDate.innerHTML = '';

  const options = [
    { value: 'latest', label: 'Последняя дата' },
    { value: 'all', label: 'Все даты' },
    ...dates.map((date) => ({ value: date, label: formatDate(date) })),
  ];

  for (const option of options) {
    const element = document.createElement('option');
    element.value = option.value;
    element.textContent = option.label;
    els.detailDate.append(element);
  }

  if (previous === 'all' || previous === 'latest' || dates.includes(previous)) {
    state.detailDate = previous;
  } else {
    state.detailDate = dates[0] || 'latest';
  }

  els.detailDate.value = state.detailDate;
}

function detailRowsForView(rows) {
  const dates = uniqueDates(rows);
  if (state.detailDate === 'all') {
    return { rows, mode: 'all', date: '' };
  }

  const selectedDate = state.detailDate === 'latest'
    ? (dates[0] || '')
    : (dates.includes(state.detailDate) ? state.detailDate : (dates[0] || ''));

  return {
    rows: selectedDate ? rows.filter((row) => row.uploadDate === selectedDate) : rows,
    mode: 'single',
    date: selectedDate,
  };
}

function renderDetailCaption(detailView) {
  if (detailView.rows.length === 0) {
    els.detailCaption.textContent = 'Нет строк по текущим фильтрам';
    return;
  }

  if (detailView.mode === 'all') {
    els.detailCaption.textContent = `Все даты · ${formatNumber(detailView.rows.length)} строк`;
    return;
  }

  els.detailCaption.textContent = `${formatDate(detailView.date)} · ${formatNumber(detailView.rows.length)} строк`;
}

function renderSourceSummaryTable(rows) {
  const summaryRows = summarizeSourceTable(rows);
  if (summaryRows.length === 0) {
    els.sourceSummaryBody.innerHTML = '<tr class="empty-row"><td colspan="7">Нет данных</td></tr>';
    return;
  }

  els.sourceSummaryBody.innerHTML = summaryRows
    .map((row) => `
      <tr>
        <td>${row.period}</td>
        <td>${row.source}</td>
        <td>${formatNumber(row.records)}</td>
        <td>${formatNumber(row.approvedMortgage)}</td>
        <td>${formatNumber(row.meetingShow)}</td>
        <td>${formatNumber(row.reservation)}</td>
        <td>${formatNumber(row.closedDeals)}</td>
      </tr>
    `)
    .join('');
}

function renderDetailTable(rows) {
  if (rows.length === 0) {
    els.detailBody.innerHTML = '<tr class="empty-row"><td colspan="11">Нет данных</td></tr>';
    return;
  }

  els.detailBody.innerHTML = rows
    .map((row) => {
      const cr = Number(row.records || 0) > 0 ? Number(row.closedDeals || 0) / Number(row.records || 0) : 0;
      return `
        <tr>
          <td>${formatDate(row.uploadDate)}</td>
          <td>${row.sourceLabel || '—'}</td>
          <td>${row.utmSource || '—'}</td>
          <td>${row.utmMedium || '—'}</td>
          <td>${row.utmCampaign || '—'}</td>
          <td>${formatNumber(row.records)}</td>
          <td>${formatNumber(row.approvedMortgage)}</td>
          <td>${formatNumber(row.meetingShow)}</td>
          <td>${formatNumber(row.reservation)}</td>
          <td>${formatNumber(row.closedDeals)}</td>
          <td>${formatPercent(cr)}</td>
        </tr>
      `;
    })
    .join('');
}

function ensureCharts() {
  if (!dailyChart) {
    dailyChart = new Chart(document.getElementById('daily-chart'), {
      type: 'bar',
      data: { labels: [], datasets: [] },
      options: {
        maintainAspectRatio: false,
        responsive: true,
        interaction: { mode: 'index', intersect: false },
        scales: {
          x: { grid: { display: false } },
          y: { beginAtZero: true, ticks: { callback: (value) => formatNumber(value) } },
          y1: {
            beginAtZero: true,
            position: 'right',
            grid: { drawOnChartArea: false },
            ticks: { callback: (value) => formatNumber(value) },
          },
        },
        plugins: {
          legend: { position: 'bottom' },
          tooltip: {
            callbacks: {
              label(context) {
                return `${context.dataset.label}: ${formatNumber(context.parsed.y)}`;
              },
            },
          },
        },
      },
    });
  }

  if (!sourceChart) {
    sourceChart = new Chart(document.getElementById('source-chart'), {
      type: 'doughnut',
      data: { labels: [], datasets: [] },
      options: {
        maintainAspectRatio: false,
        plugins: {
          legend: { position: 'bottom' },
          tooltip: {
            callbacks: {
              label(context) {
                return `${context.label}: ${formatNumber(context.parsed)}`;
              },
            },
          },
        },
      },
    });
  }

  if (!funnelChart) {
    funnelChart = new Chart(document.getElementById('funnel-chart'), {
      type: 'bar',
      data: { labels: [], datasets: [] },
      options: {
        indexAxis: 'y',
        maintainAspectRatio: false,
        responsive: true,
        scales: {
          x: { beginAtZero: true, ticks: { callback: (value) => formatNumber(value) } },
          y: { grid: { display: false } },
        },
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label(context) {
                return formatNumber(context.parsed.x);
              },
            },
          },
        },
      },
    });
  }

  if (!performanceChart) {
    performanceChart = new Chart(document.getElementById('performance-chart'), {
      type: 'bar',
      data: { labels: [], datasets: [] },
      options: {
        maintainAspectRatio: false,
        responsive: true,
        interaction: { mode: 'index', intersect: false },
        scales: {
          x: { grid: { display: false } },
          y: { beginAtZero: true, ticks: { callback: (value) => formatNumber(value) } },
        },
        plugins: {
          legend: { position: 'bottom' },
          tooltip: {
            callbacks: {
              label(context) {
                return `${context.dataset.label}: ${formatNumber(context.parsed.y)}`;
              },
            },
          },
        },
      },
    });
  }
}

function renderCharts(rows) {
  ensureCharts();

  const dailyRows = summarizeByDate(rows);
  dailyChart.data.labels = dailyRows.map((row) => formatDate(row.date));
  dailyChart.data.datasets = [
    {
      type: 'bar',
      label: 'Лиды',
      data: dailyRows.map((row) => row.records),
      backgroundColor: `${chartPalette.cyan}B3`,
      borderColor: chartPalette.cyan,
      borderWidth: 1,
      borderRadius: 4,
      yAxisID: 'y',
    },
    {
      type: 'line',
      label: 'Бронь',
      data: dailyRows.map((row) => row.reservation),
      borderColor: chartPalette.amber,
      backgroundColor: chartPalette.amber,
      tension: 0.25,
      pointRadius: 3,
      yAxisID: 'y1',
    },
    {
      type: 'line',
      label: 'Сделки',
      data: dailyRows.map((row) => row.closedDeals),
      borderColor: chartPalette.coral,
      backgroundColor: chartPalette.coral,
      tension: 0.25,
      pointRadius: 3,
      yAxisID: 'y1',
    },
  ];
  dailyChart.update();

  const sourceRows = summarizeBySource(rows);
  sourceChart.data.labels = sourceRows.map((row) => row.source);
  sourceChart.data.datasets = [{
    data: sourceRows.map((row) => row.records),
    backgroundColor: [
      chartPalette.cyan,
      chartPalette.emerald,
      chartPalette.amber,
      chartPalette.coral,
      chartPalette.lime,
      chartPalette.violet,
    ],
    borderWidth: 0,
  }];
  sourceChart.update();

  const summary = summarizeRows(rows);
  funnelChart.data.labels = ['Лиды', 'Ипотека', 'Встреча/показ', 'Бронь', 'Сделки'];
  funnelChart.data.datasets = [{
    data: [
      summary.records,
      summary.approvedMortgage,
      summary.meetingShow,
      summary.reservation,
      summary.closedDeals,
    ],
    backgroundColor: [
      chartPalette.cyan,
      chartPalette.emerald,
      chartPalette.amber,
      chartPalette.violet,
      chartPalette.coral,
    ],
    borderRadius: 6,
  }];
  funnelChart.update();

  const performanceRows = sourceRows.slice(0, 6);
  performanceChart.data.labels = performanceRows.map((row) => row.source);
  performanceChart.data.datasets = [
    {
      label: 'Ипотека',
      data: performanceRows.map((row) => row.approvedMortgage),
      backgroundColor: chartPalette.emerald,
      borderRadius: 4,
    },
    {
      label: 'Встреча/показ',
      data: performanceRows.map((row) => row.meetingShow),
      backgroundColor: chartPalette.amber,
      borderRadius: 4,
    },
    {
      label: 'Бронь',
      data: performanceRows.map((row) => row.reservation),
      backgroundColor: chartPalette.violet,
      borderRadius: 4,
    },
    {
      label: 'Сделки',
      data: performanceRows.map((row) => row.closedDeals),
      backgroundColor: chartPalette.coral,
      borderRadius: 4,
    },
  ];
  performanceChart.update();
}

function exportCsv(rows) {
  const header = ['Дата', 'Источник', 'utm source', 'utm medium', 'utm campaign', 'Лиды', 'Одобрена ипотека', 'Встреча/показ', 'Бронь', 'Сделки', 'CR'];
  const body = rows.map((row) => {
    const cr = Number(row.records || 0) > 0 ? Number(row.closedDeals || 0) / Number(row.records || 0) : 0;
    return [
      row.uploadDate,
      row.sourceLabel,
      row.utmSource,
      row.utmMedium,
      row.utmCampaign,
      row.records,
      row.approvedMortgage,
      row.meetingShow,
      row.reservation,
      row.closedDeals,
      cr,
    ];
  });
  const csv = [header, ...body].map((line) => line.map(csvEscape).join(';')).join('\r\n');
  const blob = new Blob([`\uFEFF${csv}`], { type: 'text/csv;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = 'op-dashboard-export.csv';
  link.click();
  URL.revokeObjectURL(url);
}

function render() {
  const rows = filteredRows();
  populateDetailDateSelect(rows);
  const detailView = detailRowsForView(rows);
  renderHero(rows);
  renderKpis(rows);
  renderActiveState(rows);
  renderSourceSummaryTable(rows);
  renderDetailCaption(detailView);
  renderDetailTable(detailView.rows);
  renderCharts(rows);
}

function bindHeaderState() {
  if (!els.siteHeader || els.navLinks.length === 0) return;

  const sections = els.navLinks
    .map((link) => document.querySelector(link.getAttribute('href')))
    .filter(Boolean);

  const syncHeader = () => {
    els.siteHeader.classList.toggle('is-scrolled', window.scrollY > 16);

    const checkpoint = window.scrollY + 130;
    let activeSection = sections[0];
    for (const section of sections) {
      if (section.offsetTop <= checkpoint) activeSection = section;
    }

    for (const link of els.navLinks) {
      const target = link.getAttribute('href');
      link.classList.toggle('is-active', activeSection && `#${activeSection.id}` === target);
    }
  };

  syncHeader();
  window.addEventListener('scroll', syncHeader, { passive: true });
}

function bindControls() {
  els.sourceLabel.addEventListener('change', () => {
    state.sourceLabel = els.sourceLabel.value;
    render();
  });
  els.utmSource.addEventListener('change', () => {
    state.utmSource = els.utmSource.value;
    render();
  });
  els.utmMedium.addEventListener('change', () => {
    state.utmMedium = els.utmMedium.value;
    render();
  });
  els.utmCampaign.addEventListener('change', () => {
    state.utmCampaign = els.utmCampaign.value;
    render();
  });
  els.dateFrom.addEventListener('change', () => {
    state.dateFrom = els.dateFrom.value;
    render();
  });
  els.dateTo.addEventListener('change', () => {
    state.dateTo = els.dateTo.value;
    render();
  });
  els.search.addEventListener('input', () => {
    state.search = els.search.value;
    render();
  });
  els.detailDate.addEventListener('change', () => {
    state.detailDate = els.detailDate.value;
    render();
  });
  els.reset.addEventListener('click', () => {
    state.sourceLabel = 'all';
    state.utmSource = 'all';
    state.utmMedium = 'all';
    state.utmCampaign = 'all';
    state.dateFrom = data.filters.minDate || '';
    state.dateTo = data.filters.maxDate || '';
    state.detailDate = 'latest';
    state.search = '';

    els.sourceLabel.value = 'all';
    els.utmSource.value = 'all';
    els.utmMedium.value = 'all';
    els.utmCampaign.value = 'all';
    els.dateFrom.value = state.dateFrom;
    els.dateTo.value = state.dateTo;
    els.detailDate.value = state.detailDate;
    els.search.value = '';
    render();
  });
  els.exportCsv.addEventListener('click', () => exportCsv(detailRowsForView(filteredRows()).rows));
}

function init() {
  if (!data) {
    document.body.innerHTML = '<main class="page-shell"><section class="panel"><div class="panel-head"><h2>Нет данных</h2></div></section></main>';
    return;
  }

  if (els.tableLink && data.report?.tableUrl) {
    els.tableLink.href = data.report.tableUrl;
  }
  if (els.kcReportLink && data.report?.kcDashboardUrl) {
    els.kcReportLink.href = data.report.kcDashboardUrl;
  }

  populateSelect(els.sourceLabel, data.filters.sourceLabels || [], 'Все источники');
  populateSelect(els.utmSource, data.filters.utmSources || [], 'Все utm source');
  populateSelect(els.utmMedium, data.filters.utmMediums || [], 'Все utm medium');
  populateSelect(els.utmCampaign, data.filters.utmCampaigns || [], 'Все utm campaign');

  els.dateFrom.value = state.dateFrom;
  els.dateTo.value = state.dateTo;
  els.dateFrom.min = data.filters.minDate || '';
  els.dateFrom.max = data.filters.maxDate || '';
  els.dateTo.min = data.filters.minDate || '';
  els.dateTo.max = data.filters.maxDate || '';

  bindControls();
  bindHeaderState();
  render();
}

init();
