let data = null;

const state = {
  sourceLabel: 'all',
  dateFrom: '',
  dateTo: '',
  detailDate: 'latest',
  search: '',
};

const els = {
  siteHeader: document.querySelector('.site-header'),
  navLinks: [...document.querySelectorAll('.site-nav a')],
  tableLink: document.getElementById('table-link'),
  kcReportLink: document.getElementById('kc-report-link'),
  heroSources: document.getElementById('hero-sources'),
  heroRecords: document.getElementById('hero-records'),
  heroDeals: document.getElementById('hero-deals'),
  heroApproved: document.getElementById('hero-approved'),
  heroMeetings: document.getElementById('hero-meetings'),
  heroReservations: document.getElementById('hero-reservations'),
  heroClosed: document.getElementById('hero-closed'),
  sourceLabel: document.getElementById('filter-source-label'),
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
  kpiDeals: document.getElementById('kpi-deals'),
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

function clampDate(value, minValue, maxValue) {
  if (!value) return '';
  if (minValue && value < minValue) return minValue;
  if (maxValue && value > maxValue) return maxValue;
  return value;
}

function localIsoDate(year, monthIndex, day) {
  return `${year}-${String(monthIndex + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
}

function resolveDefaultDateRange(filters = {}) {
  const minDate = filters.minDate || '';
  const maxDate = filters.maxDate || '';
  if (!maxDate) {
    return { dateFrom: minDate, dateTo: maxDate };
  }

  const anchor = new Date(`${maxDate}T00:00:00`);
  const monthStart = localIsoDate(anchor.getFullYear(), anchor.getMonth(), 1);
  return {
    dateFrom: clampDate(monthStart, minDate, maxDate),
    dateTo: maxDate,
  };
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
    if (state.dateFrom && row.uploadDate < state.dateFrom) return false;
    if (state.dateTo && row.uploadDate > state.dateTo) return false;
    if (!query) return true;

    return normalizeSearch(row.sourceLabel).includes(query);
  });
}

function rawMeetingDailyTotals() {
  return Array.isArray(data?.rawMeetingDailyTotals) ? data.rawMeetingDailyTotals : [];
}

function shouldUseRawMeetingTotals() {
  return state.sourceLabel === 'all' && !normalizeSearch(state.search);
}

function sumRawMeetingTotalsForState() {
  return rawMeetingDailyTotals().reduce((acc, row) => {
    if (state.dateFrom && row.uploadDate < state.dateFrom) return acc;
    if (state.dateTo && row.uploadDate > state.dateTo) return acc;
    return acc + Number(row.meetingShow || 0);
  }, 0);
}

function summarizeRowsBase(rows) {
  return rows.reduce((acc, row) => {
    acc.records += Number(row.records || 0);
    acc.deals += Number(row.deals || 0);
    acc.approvedMortgage += Number(row.approvedMortgage || 0);
    acc.meetingShow += Number(row.meetingShow || 0);
    acc.reservation += Number(row.reservation || 0);
    acc.sales += Number((row.sales ?? row.closedDeals) || 0);
    return acc;
  }, {
    records: 0,
    deals: 0,
    approvedMortgage: 0,
    meetingShow: 0,
    reservation: 0,
    sales: 0,
  });
}

function summarizeRows(rows) {
  const summary = summarizeRowsBase(rows);
  if (shouldUseRawMeetingTotals()) {
    summary.meetingShow = sumRawMeetingTotalsForState();
  }
  return summary;
}

function summarizeByDate(rows) {
  const groups = new Map();

  for (const row of rows) {
    const key = row.uploadDate;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(row);
  }

  const summaryRows = [...groups.entries()]
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([date, items]) => ({
      date,
      ...summarizeRowsBase(items),
    }));

  if (!shouldUseRawMeetingTotals()) {
    return summaryRows;
  }

  const summaryMap = new Map(summaryRows.map((row) => [row.date, row]));
  for (const row of rawMeetingDailyTotals()) {
    if (state.dateFrom && row.uploadDate < state.dateFrom) continue;
    if (state.dateTo && row.uploadDate > state.dateTo) continue;

    if (!summaryMap.has(row.uploadDate)) {
      summaryMap.set(row.uploadDate, {
        date: row.uploadDate,
        records: 0,
        deals: 0,
        approvedMortgage: 0,
        meetingShow: 0,
        reservation: 0,
        sales: 0,
      });
    }

    summaryMap.get(row.uploadDate).meetingShow = Number(row.meetingShow || 0);
  }

  return [...summaryMap.values()].sort((a, b) => a.date.localeCompare(b.date));
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
        deals: 0,
        approvedMortgage: 0,
        meetingShow: 0,
        reservation: 0,
        sales: 0,
      });
    }

    const group = groups.get(key);
    group.records += Number(row.records || 0);
    group.deals += Number(row.deals || 0);
    group.approvedMortgage += Number(row.approvedMortgage || 0);
    group.meetingShow += Number(row.meetingShow || 0);
    group.reservation += Number(row.reservation || 0);
    group.sales += Number((row.sales ?? row.closedDeals) || 0);
  }

  return [...groups.values()].sort((a, b) => `${a.month}_${a.source}`.localeCompare(`${b.month}_${b.source}`));
}

function renderHero(rows) {
  const summary = summarizeRows(rows);
  const uniqueSources = new Set(rows.map((row) => row.sourceLabel).filter(Boolean)).size;

  setText(els.heroSources, formatNumber(uniqueSources));
  setText(els.heroRecords, formatNumber(summary.records));
  setText(els.heroDeals, formatNumber(summary.deals));
  setText(els.heroApproved, formatNumber(summary.approvedMortgage));
  setText(els.heroMeetings, formatNumber(summary.meetingShow));
  setText(els.heroReservations, formatNumber(summary.reservation));
  setText(els.heroClosed, formatNumber(summary.sales));
}

function renderKpis(rows) {
  const summary = summarizeRows(rows);
  els.kpiRecords.textContent = formatNumber(summary.records);
  els.kpiDeals.textContent = formatNumber(summary.deals);
  els.kpiApproved.textContent = formatNumber(summary.approvedMortgage);
  els.kpiMeetings.textContent = formatNumber(summary.meetingShow);
  els.kpiReservations.textContent = formatNumber(summary.reservation);
  els.kpiClosed.textContent = formatNumber(summary.sales);
  els.kpiCr.textContent = formatPercent(summary.deals > 0 ? summary.sales / summary.deals : 0);
}

function renderActiveState(rows) {
  const chips = [];
  if (state.sourceLabel !== 'all') chips.push(`Источник: ${state.sourceLabel}`);
  if (state.dateFrom) chips.push(`От: ${formatDate(state.dateFrom)}`);
  if (state.dateTo) chips.push(`До: ${formatDate(state.dateTo)}`);
  if (normalizeSearch(state.search)) chips.push(`Поиск: ${state.search.trim()}`);

  els.activeFilters.innerHTML = chips.length > 0
    ? chips.map((chip) => `<span class="chip">${chip}</span>`).join('')
    : '<span class="chip">Все данные</span>';

  const summary = summarizeRows(rows);
  const uniqueSources = new Set(rows.map((row) => row.sourceLabel).filter(Boolean)).size;
  els.selectionSummary.textContent = `Строк: ${formatNumber(rows.length)} · Источников: ${formatNumber(uniqueSources)} · Лидов: ${formatNumber(summary.records)} · Сделок: ${formatNumber(summary.deals)} · Продаж: ${formatNumber(summary.sales)}`;
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
    els.sourceSummaryBody.innerHTML = '<tr class="empty-row"><td colspan="9">Нет данных</td></tr>';
    return;
  }

  els.sourceSummaryBody.innerHTML = summaryRows
    .map((row) => `
      <tr>
        <td>${row.period}</td>
        <td>${row.source}</td>
        <td>${formatNumber(row.records)}</td>
        <td>${formatNumber(row.deals)}</td>
        <td>${formatPercent(Number(row.deals || 0) > 0 ? Number(row.sales || 0) / Number(row.deals || 0) : 0)}</td>
        <td>${formatNumber(row.approvedMortgage)}</td>
        <td>${formatNumber(row.meetingShow)}</td>
        <td>${formatNumber(row.reservation)}</td>
        <td>${formatNumber(row.sales)}</td>
      </tr>
    `)
    .join('');
}

function renderDetailTable(rows) {
  if (rows.length === 0) {
    els.detailBody.innerHTML = '<tr class="empty-row"><td colspan="12">Нет данных</td></tr>';
    return;
  }

  els.detailBody.innerHTML = rows
    .map((row) => {
      const sales = Number((row.sales ?? row.closedDeals) || 0);
      const saleConversion = Number(row.deals || 0) > 0 ? sales / Number(row.deals || 0) : 0;
      return `
        <tr>
          <td>${formatDate(row.uploadDate)}</td>
          <td>${row.sourceLabel || '—'}</td>
          <td>${row.utmSource || '—'}</td>
          <td>${row.utmMedium || '—'}</td>
          <td>${row.utmCampaign || '—'}</td>
          <td>${formatNumber(row.records)}</td>
          <td>${formatNumber(row.deals)}</td>
          <td>${formatPercent(saleConversion)}</td>
          <td>${formatNumber(row.approvedMortgage)}</td>
          <td>${formatNumber(row.meetingShow)}</td>
          <td>${formatNumber(row.reservation)}</td>
          <td>${formatNumber(sales)}</td>
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
      label: 'Сделки',
      data: dailyRows.map((row) => row.deals),
      borderColor: chartPalette.violet,
      backgroundColor: chartPalette.violet,
      tension: 0.25,
      pointRadius: 3,
      yAxisID: 'y1',
    },
    {
      type: 'line',
      label: 'Продажи',
      data: dailyRows.map((row) => (row.sales ?? row.closedDeals)),
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
  funnelChart.data.labels = ['Лиды', 'Сделки', 'Ипотека', 'Встреча/показ', 'Бронь', 'Продажи'];
  funnelChart.data.datasets = [{
    data: [
      summary.records,
      summary.deals,
      summary.approvedMortgage,
      summary.meetingShow,
      summary.reservation,
      summary.sales,
    ],
    backgroundColor: [
      chartPalette.cyan,
      chartPalette.violet,
      chartPalette.emerald,
      chartPalette.amber,
      chartPalette.lime,
      chartPalette.coral,
    ],
    borderRadius: 6,
  }];
  funnelChart.update();

  const performanceRows = sourceRows.slice(0, 6);
  performanceChart.data.labels = performanceRows.map((row) => row.source);
  performanceChart.data.datasets = [
    {
      label: 'Сделки',
      data: performanceRows.map((row) => row.deals),
      backgroundColor: chartPalette.violet,
      borderRadius: 4,
    },
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
      label: 'Продажи',
      data: performanceRows.map((row) => (row.sales ?? row.closedDeals)),
      backgroundColor: chartPalette.coral,
      borderRadius: 4,
    },
  ];
  performanceChart.update();
}

function exportCsv(rows) {
  const header = ['Дата', 'Источник', 'utm source', 'utm medium', 'utm campaign', 'Лиды', 'Сделки', 'Конверсия в продажу', 'Одобрена ипотека', 'Встреча/показ', 'Бронь', 'Продажи'];
  const body = rows.map((row) => {
    const sales = Number((row.sales ?? row.closedDeals) || 0);
    const saleConversion = Number(row.deals || 0) > 0 ? sales / Number(row.deals || 0) : 0;
    return [
      row.uploadDate,
      row.sourceLabel,
      row.utmSource,
      row.utmMedium,
      row.utmCampaign,
      row.records,
      row.deals,
      saleConversion,
      row.approvedMortgage,
      row.meetingShow,
      row.reservation,
      sales,
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
    const defaultRange = resolveDefaultDateRange(data.filters);
    state.sourceLabel = 'all';
    state.dateFrom = defaultRange.dateFrom;
    state.dateTo = defaultRange.dateTo;
    state.detailDate = 'latest';
    state.search = '';

    els.sourceLabel.value = 'all';
    els.dateFrom.value = state.dateFrom;
    els.dateTo.value = state.dateTo;
    els.detailDate.value = state.detailDate;
    els.search.value = '';
    render();
  });
  els.exportCsv.addEventListener('click', () => exportCsv(detailRowsForView(filteredRows()).rows));
}

function init() {
  const defaultRange = resolveDefaultDateRange(data?.filters);
  state.dateFrom = defaultRange.dateFrom;
  state.dateTo = defaultRange.dateTo;

  if (els.tableLink && data.report?.tableUrl) {
    els.tableLink.href = data.report.tableUrl;
  }
  if (els.kcReportLink && data.report?.kcDashboardUrl) {
    els.kcReportLink.href = data.report.kcDashboardUrl;
  }

  populateSelect(els.sourceLabel, data.filters.sourceLabels || [], 'Все источники');

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

async function loadData() {
  try {
    const response = await fetch('./data/report-data.json', { cache: 'no-store' });
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }
    return await response.json();
  } catch (error) {
    return window.REPORT_DASHBOARD_DATA || null;
  }
}

async function bootstrap() {
  data = await loadData();
  if (!data) {
    document.body.innerHTML = '<main class="page-shell"><section class="panel"><div class="panel-head"><h2>Нет данных</h2><p>Файл дашборда пока не сгенерирован.</p></div></section></main>';
    return;
  }
  init();
}

bootstrap();
