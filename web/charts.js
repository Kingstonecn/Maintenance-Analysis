/* ===================================================================
   charts.js — Chart.js 封装
   提供统一的图表创建/销毁/导出接口
   =================================================================== */
'use strict';

const Charts = {
  _instances: new Map(),

  _colors: ['#1d4e89', '#b1281e', '#2a7a7a', '#a98437', '#3e6b3a', '#6b6258', '#3a6ba0', '#c9483c', '#3a9999', '#c9a64e',
    '#5a3a8a', '#8a5a3a', '#3a8a5a', '#8a3a6a', '#5a5a3a', '#3a5a8a', '#7a3a3a', '#3a7a5a', '#6a5a3a', '#5a7a8a'],

  _fontFamily: "'JetBrains Mono', monospace",

  _baseOpts: {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: { labels: { font: { family: "'Noto Serif SC', serif", size: 12 } } },
      tooltip: { titleFont: { family: "'Noto Serif SC', serif" }, bodyFont: { family: "'JetBrains Mono', monospace" } }
    },
    scales: {
      x: { ticks: { font: { family: "'JetBrains Mono', monospace", size: 10 }, maxRotation: 45, minRotation: 0 } },
      y: { ticks: { font: { family: "'JetBrains Mono', monospace", size: 10 } } }
    }
  },

  _mergeOpts(extra) {
    return Charts._deepMerge(JSON.parse(JSON.stringify(Charts._baseOpts)), extra || {});
  },

  _deepMerge(a, b) {
    for (const k in b) {
      if (b[k] && typeof b[k] === 'object' && !Array.isArray(b[k])) {
        a[k] = Charts._deepMerge(a[k] || {}, b[k]);
      } else { a[k] = b[k]; }
    }
    return a;
  },

  _getCtx(canvasId) {
    const canvas = document.getElementById(canvasId);
    if (!canvas) return null;
    const ctx = canvas.getContext('2d');
    return ctx;
  },

  _destroy(canvasId) {
    if (Charts._instances.has(canvasId)) {
      Charts._instances.get(canvasId).destroy();
      Charts._instances.delete(canvasId);
    }
  },

  bar(canvasId, labels, datasets, opts) {
    Charts._destroy(canvasId);
    const ctx = Charts._getCtx(canvasId);
    if (!ctx) return null;
    const config = {
      type: 'bar',
      data: { labels, datasets: datasets.map((d, i) => ({
        ...d,
        backgroundColor: d.backgroundColor || Charts._colors[i % Charts._colors.length],
        borderColor: d.borderColor || Charts._colors[i % Charts._colors.length],
        borderWidth: d.borderWidth || 1
      }))},
      options: Charts._mergeOpts(opts)
    };
    const chart = new Chart(ctx, config);
    Charts._instances.set(canvasId, chart);
    return chart;
  },

  horizontalBar(canvasId, labels, datasets, opts) {
    return Charts.bar(canvasId, labels, datasets, { ...opts, indexAxis: 'y' });
  },

  line(canvasId, labels, datasets, opts) {
    Charts._destroy(canvasId);
    const ctx = Charts._getCtx(canvasId);
    if (!ctx) return null;
    const config = {
      type: 'line',
      data: { labels, datasets: datasets.map((d, i) => ({
        ...d,
        borderColor: d.borderColor || Charts._colors[i % Charts._colors.length],
        backgroundColor: d.backgroundColor || Charts._colors[i % Charts._colors.length] + '20',
        borderWidth: d.borderWidth || 2,
        tension: d.tension !== undefined ? d.tension : 0.3,
        fill: d.fill !== undefined ? d.fill : false,
        pointRadius: d.pointRadius !== undefined ? d.pointRadius : 3
      }))},
      options: Charts._mergeOpts(opts)
    };
    const chart = new Chart(ctx, config);
    Charts._instances.set(canvasId, chart);
    return chart;
  },

  pie(canvasId, labels, data, opts) {
    Charts._destroy(canvasId);
    const ctx = Charts._getCtx(canvasId);
    if (!ctx) return null;
    const config = {
      type: 'pie',
      data: { labels, datasets: [{ data, backgroundColor: labels.map((_, i) => Charts._colors[i % Charts._colors.length]), borderWidth: 1, borderColor: '#f3ecdc' }] },
      options: Charts._mergeOpts({ ...opts, scales: {} })
    };
    const chart = new Chart(ctx, config);
    Charts._instances.set(canvasId, chart);
    return chart;
  },

  doughnut(canvasId, labels, data, opts) {
    Charts._destroy(canvasId);
    const ctx = Charts._getCtx(canvasId);
    if (!ctx) return null;
    const config = {
      type: 'doughnut',
      data: { labels, datasets: [{ data, backgroundColor: labels.map((_, i) => Charts._colors[i % Charts._colors.length]), borderWidth: 1, borderColor: '#f3ecdc' }] },
      options: Charts._mergeOpts({ ...opts, scales: {} })
    };
    const chart = new Chart(ctx, config);
    Charts._instances.set(canvasId, chart);
    return chart;
  },

  scatter(canvasId, datasets, opts) {
    Charts._destroy(canvasId);
    const ctx = Charts._getCtx(canvasId);
    if (!ctx) return null;
    const config = {
      type: 'scatter',
      data: { datasets: datasets.map((d, i) => ({
        ...d,
        backgroundColor: d.backgroundColor || Charts._colors[i % Charts._colors.length] + '80',
        borderColor: d.borderColor || Charts._colors[i % Charts._colors.length],
        pointRadius: d.pointRadius || 4
      }))},
      options: Charts._mergeOpts(opts)
    };
    const chart = new Chart(ctx, config);
    Charts._instances.set(canvasId, chart);
    return chart;
  },

  radar(canvasId, labels, datasets, opts) {
    Charts._destroy(canvasId);
    const ctx = Charts._getCtx(canvasId);
    if (!ctx) return null;
    const config = {
      type: 'radar',
      data: { labels, datasets: datasets.map((d, i) => ({
        ...d,
        borderColor: d.borderColor || Charts._colors[i % Charts._colors.length],
        backgroundColor: d.backgroundColor || Charts._colors[i % Charts._colors.length] + '30',
        borderWidth: 2
      }))},
      options: Charts._mergeOpts({ ...opts, scales: { r: { beginAtZero: true, ticks: { font: { family: "'JetBrains Mono', monospace", size: 9 } } } } })
    };
    const chart = new Chart(ctx, config);
    Charts._instances.set(canvasId, chart);
    return chart;
  },

  toImage(canvasId) {
    const chart = Charts._instances.get(canvasId);
    if (!chart) return null;
    return chart.toBase64Image('image/png', 1);
  },

  downloadImage(canvasId, filename) {
    const url = Charts.toImage(canvasId);
    if (!url) return;
    const a = document.createElement('a');
    a.href = url;
    a.download = filename + '.png';
    document.body.appendChild(a);
    a.click();
    setTimeout(() => a.remove(), 1000);
  },

  destroyAll() {
    Charts._instances.forEach(c => c.destroy());
    Charts._instances.clear();
  }
};
