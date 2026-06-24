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
    if (!item || typeof item !== 'object') return;

    labels.push(item.category !== undefined ? item.category : '');
    values.push(typeof item.value === 'number' ? item.value : Number(item.value) || 0);
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
    showLegend = false,
    showValue = false,
    showCategoryName = false
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
    showLabel: showCategoryName
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
  const chartData = createChartSeries(data, options.title || 'Series 1');
  const chartOptions = buildChartOptions(options);
  return slide.addChart('pie', chartData, chartOptions);
}

module.exports = {
  ChartElement,
  addBarChart,
  addLineChart,
  addPieChart
};