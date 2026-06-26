'use strict';

/**
 * Chart element class
 */
class ChartElement {
  constructor(type = 'bar', data = []) {
    this.type = 'chart';
    this.chartType = type;
    this.data = data;
    this.position = { x: 0, y: 0 };
    this.size = { width: 300, height: 200 };
  }

  /**
   * Set the position of the chart element
   * @param {number} x - X coordinate
   * @param {number} y - Y coordinate
   */
  setPosition(x, y) {
    this.position.x = x;
    this.position.y = y;
  }

  /**
   * Set the size of the chart element
   * @param {number} width - Width
   * @param {number} height - Height
   */
  setSize(width, height) {
    this.size.width = width;
    this.size.height = height;
  }

  /**
   * Set chart data
   * @param {Array} data - Chart data
   */
  setData(data) {
    this.data = data;
  }
}

function normalizeChartData(data = []) {
  const labels = [];
  const values = [];

  if (!Array.isArray(data)) {
    return { labels, values };
  }

  data.forEach(item => {
    if (item === null || item === undefined) return;

    // Array-pair format: ["Jan", 10]
    if (Array.isArray(item)) {
      const [label, value] = item;
      labels.push(label !== undefined ? label : '');
      values.push(typeof value === 'number' ? value : Number(value) || 0);
      return;
    }

    // Object format: { category: "Jan", value: 10 }
    if (typeof item === 'object') {
      labels.push(item.category !== undefined ? item.category : '');
      values.push(typeof item.value === 'number' ? item.value : Number(item.value) || 0);
    }
  });

  return { labels, values };
}

function buildChartOptions(options = {}) {
  const {
    x = 0.5,
    y = 1,
    w = 8,
    h = 4.5,
    title,
    showLegend = true,
    showValue = false,
    showCategoryName = true,
    showPercentage = true
  } = options;

  return {
    x,
    y,
    w,
    h,
    title,
    showTitle: !!title,
    titleAlign: title ? 'center' : undefined,
    titleFontSize: title ? 18 : undefined,
    titleColor: title ? '#000000' : undefined,
    titleBold: title ? true : undefined,
    showLegend,
    legendPos: showLegend ? 'r' : undefined,
    showValue,
    showLabel: showCategoryName,
    showPercent: showPercentage
  };
}

function createChartSeries(data = [], seriesName = 'Series 1') {
  const { labels, values } = normalizeChartData(data);

  return [
    {
      name: seriesName,
      labels,
      values
    }
  ];
}

/**
 * Builds category labels with each value's percentage share appended,
 * e.g. "A" -> "A (40.0%)". Used so the legend (which pie charts draw
 * from category labels) shows percentage composition per slice.
 * @param {string[]} labels
 * @param {number[]} values
 * @returns {string[]}
 */
function buildPercentageLabels(labels, values) {
  const total = values.reduce((sum, v) => sum + v, 0);

  return labels.map((label, i) => {
    const value = values[i];
    const pct = total > 0 ? ((value / total) * 100).toFixed(1) : '0.0';
    return `${label} (${pct}%)`;
  });
}

function addBarChart(slide, data, options = {}) {
  const chartData = createChartSeries(data, options.title || 'Series 1');
  const chartOptions = buildChartOptions(options);
  return slide.addChart('bar', chartData, chartOptions);
}

function addLineChart(slide, data, options = {}) {
  const chartData = createChartSeries(data, options.title || 'Series 1');
  const chartOptions = buildChartOptions(options);
  return slide.addChart('line', chartData, chartOptions);
}

function addPieChart(slide, data, options = {}) {
  const { labels, values } = normalizeChartData(data);
  const percentageLabels = buildPercentageLabels(labels, values);

  const chartData = [
    {
      name: options.title || 'Series 1',
      labels: percentageLabels,
      values
    }
  ];

  const chartOptions = buildChartOptions(options);
  return slide.addChart('pie', chartData, chartOptions);
}

module.exports = {
  ChartElement,
  addBarChart,
  addLineChart,
  addPieChart
};