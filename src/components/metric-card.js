'use strict';

const { addText } = require('../elements/text');

/**
 * Render a metric card with title and value.
 * @param {Object} slide - The slide instance.
 * @param {Object} config - Metric card configuration.
 * @param {string} config.title - Card title.
 * @param {string|number} config.value - Metric value.
 * @param {number} config.x - X coordinate.
 * @param {number} config.y - Y coordinate.
 * @param {number} config.w - Width.
 * @param {number} config.h - Height.
 * @param {Object} theme - Theme instance.
 * @returns {Object[]} Added objects.
 */
function renderMetricCard(slide, config = {}, theme) {
  const {
    title = '',
    value = '',
    x = 0.5,
    y = 1,
    w = 4,
    h = 2.5
  } = config;

  const backgroundColor = theme && typeof theme.getColor === 'function'
    ? theme.getColor('secondaryColor') || '#F2F2F2'
    : '#F2F2F2';

  const textColor = theme && typeof theme.getColor === 'function'
    ? theme.getColor('textColor') || '#333333'
    : '#333333';

  const titleFont = theme && typeof theme.getFont === 'function'
    ? theme.getFont('bodyFont') || 'Arial'
    : 'Arial';

  const titleFontSize = theme && typeof theme.getFontSize === 'function'
    ? Math.max(theme.getFontSize('bodyFontSize') - 2, 14)
    : 16;

  const valueFontSize = theme && typeof theme.getFontSize === 'function'
    ? Math.max(theme.getFontSize('titleFontSize'), 24)
    : 24;

  const card = slide.addShape('rect', {
    x,
    y,
    w,
    h,
    fill: { color: backgroundColor },
    line: { color: theme && typeof theme.getColor === 'function' ? theme.getColor('primaryColor') || '#1F4E79' : '#1F4E79' }
  });

  const titleText = addText(slide, title, {
    x: x + 0.2,
    y: y + 0.2,
    w: w - 0.4,
    h: 0.75,
    fontFace: titleFont,
    fontSize: titleFontSize,
    color: textColor,
    bold: true,
    align: 'left'
  });

  const valueText = addText(slide, String(value), {
    x: x + 0.2,
    y: y + 0.95,
    w: w - 0.4,
    h: 1.2,
    fontFace: theme && typeof theme.getFont === 'function' ? theme.getFont('titleFont') || 'Arial Black' : 'Arial Black',
    fontSize: valueFontSize,
    color: textColor,
    bold: true,
    align: 'left'
  });

  return [card, titleText, valueText];
}

module.exports = {
  renderMetricCard
};
