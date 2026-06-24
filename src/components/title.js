'use strict';

const { addText } = require('../elements/text');

/**
 * Render a styled title block on a slide.
 * @param {Object} slide - The slide instance.
 * @param {Object} config - Title configuration.
 * @param {string} config.text - Title text.
 * @param {number} config.x - X coordinate.
 * @param {number} config.y - Y coordinate.
 * @param {number} config.w - Width.
 * @param {number} config.h - Height.
 * @param {Object} theme - Theme instance.
 * @returns {Object} Added text object.
 */
function renderTitle(slide, config = {}, theme) {
  const {
    text = '',
    x = 0.5,
    y = 0.5,
    w = 8,
    h = 1
  } = config;

  const fontFace = theme && typeof theme.getFont === 'function'
    ? theme.getFont('titleFont') || 'Arial Black'
    : 'Arial Black';

  const fontSize = theme && typeof theme.getFontSize === 'function'
    ? theme.getFontSize('titleFontSize') || 28
    : 28;

  const color = theme && typeof theme.getColor === 'function'
    ? theme.getColor('primaryColor') || '#1F4E79'
    : '#1F4E79';

  return addText(slide, text, {
    x,
    y,
    w,
    h,
    fontFace,
    fontSize,
    color,
    bold: true,
    align: 'center'
  });
}

module.exports = {
  renderTitle
};
