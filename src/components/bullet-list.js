'use strict';

const { addText } = require('../elements/text');

/**
 * Render a bullet list on a slide.
 * @param {Object} slide - The slide instance.
 * @param {Object} config - Bullet list configuration.
 * @param {Array<string>} config.items - List items.
 * @param {number} config.x - X coordinate.
 * @param {number} config.y - Y coordinate.
 * @param {number} config.w - Width.
 * @param {number} config.h - Height.
 * @param {Object} theme - Theme instance.
 * @returns {Object} Added text object.
 */
function renderBulletList(slide, config = {}, theme) {
  const {
    items = [],
    x = 0.5,
    y = 1.75,
    w = 8,
    h = 4
  } = config;

  const fontFace = theme && typeof theme.getFont === 'function'
    ? theme.getFont('bodyFont') || 'Arial'
    : 'Arial';

  const fontSize = theme && typeof theme.getFontSize === 'function'
    ? theme.getFontSize('bodyFontSize') || 18
    : 18;

  const color = theme && typeof theme.getColor === 'function'
    ? theme.getColor('textColor') || '#333333'
    : '#333333';

  const text = Array.isArray(items) ? items.map(item => `• ${item}`).join('\n') : String(items);

  return addText(slide, text, {
    x,
    y,
    w,
    h,
    fontFace,
    fontSize,
    color,
    align: 'left'
  });
}

module.exports = {
  renderBulletList
};
