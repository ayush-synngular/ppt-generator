'use strict';

const { addText } = require('../elements/text');
const { addBarChart, addLineChart, addPieChart } = require('../elements/chart');

/**
 * Render a chart card with title and chart content.
 * @param {Object} slide - The slide instance.
 * @param {Object} config - Chart card configuration.
 * @param {string} config.title - Card title.
 * @param {string} config.chartType - Chart type: bar, line, pie.
 * @param {Array} config.data - Chart data.
 * @param {number} config.x - X coordinate.
 * @param {number} config.y - Y coordinate.
 * @param {number} config.w - Width.
 * @param {number} config.h - Height.
 * @param {Object} theme - Theme instance.
 * @returns {Object[]} Added objects.
 */
function renderChartCard(slide, config = {}, theme) {
  const {
    title = '',
    chartType = 'bar',
    data = [],
    x = 0.5,
    y = 0.5,
    w = 8,
    h = 5
  } = config;

  const backgroundColor = theme && typeof theme.getColor === 'function'
    ? theme.getColor('secondaryColor') || '#F9FAFB'
    : '#F9FAFB';

  const borderColor = theme && typeof theme.getColor === 'function'
    ? theme.getColor('primaryColor') || '#1F4E79'
    : '#1F4E79';

  const titleColor = theme && typeof theme.getColor === 'function'
    ? theme.getColor('primaryColor') || '#1F4E79'
    : '#1F4E79';

  const titleFont = theme && typeof theme.getFont === 'function'
    ? theme.getFont('titleFont') || 'Arial Black'
    : 'Arial Black';

  const titleFontSize = theme && typeof theme.getFontSize === 'function'
    ? theme.getFontSize('titleFontSize') || 24
    : 24;

  const card = slide.addShape('rect', {
    x,
    y,
    w,
    h,
    fill: { color: backgroundColor },
    line: { color: borderColor }
  });

  const titleText = addText(slide, title, {
    x: x + 0.2,
    y: y + 0.2,
    w: w - 0.4,
    h: 0.8,
    fontFace: titleFont,
    fontSize: titleFontSize,
    color: titleColor,
    bold: true,
    align: 'left'
  });

  const chartOptions = {
    x: x + 0.35,
    y: y + 1.2,
    w: w - 0.7,
    h: h - 1.5,
    showLegend: true,
    showValue: true,
    showCategoryName: true
  };

  let chart;
  if (chartType === 'pie') {
    chart = addPieChart(slide, data, chartOptions);
  } else if (chartType === 'line') {
    chart = addLineChart(slide, data, chartOptions);
  } else {
    chart = addBarChart(slide, data, chartOptions);
  }

  return [card, titleText, chart];
}

module.exports = {
  renderChartCard
};
