'use strict';

const { addText } = require('../elements/text');

function getThemeValue(theme, getter, fallback) {
  if (theme && typeof theme[getter] === 'function') {
    const value = theme[getter](fallback.key);
    return value !== undefined ? value : fallback.default;
  }

  return fallback.default;
}

function formatList(items) {
  if (!Array.isArray(items)) {
    return String(items || '');
  }

  return items.map(item => `• ${item}`).join('\n');
}

/**
 * Create a comparison slide with left and right item lists.
 * @param {Object} presentation - PPT generator presentation.
 * @param {Object} options - Layout settings.
 * @param {string} options.title - Slide title.
 * @param {string} options.leftTitle - Left column heading.
 * @param {string} options.rightTitle - Right column heading.
 * @param {Array<string>} options.leftItems - Left list items.
 * @param {Array<string>} options.rightItems - Right list items.
 * @param {Object} options.theme - Theme instance for styling.
 * @returns {Object} The created slide.
 */
function createComparisonSlide(presentation, options = {}) {
  const {
    title = 'Comparison Layout',
    leftTitle = 'Left',
    rightTitle = 'Right',
    leftItems = [],
    rightItems = [],
    theme
  } = options;

  const slide = presentation.addSlide();

  const titleFont = getThemeValue(theme, 'getFont', { key: 'titleFont', default: 'Arial Black' });
  const titleFontSize = getThemeValue(theme, 'getFontSize', { key: 'titleFontSize', default: 30 });
  const titleColor = getThemeValue(theme, 'getColor', { key: 'primaryColor', default: '#1F4E79' });
  const headingFontSize = getThemeValue(theme, 'getFontSize', { key: 'titleFontSize', default: 24 });
  const headingColor = getThemeValue(theme, 'getColor', { key: 'accentColor', default: '#FFC000' });
  const bodyFont = getThemeValue(theme, 'getFont', { key: 'bodyFont', default: 'Arial' });
  const bodyFontSize = getThemeValue(theme, 'getFontSize', { key: 'bodyFontSize', default: 18 });
  const textColor = getThemeValue(theme, 'getColor', { key: 'textColor', default: '#333333' });

  addText(slide, title, {
    x: 0.5,
    y: 0.5,
    w: 9,
    h: 1,
    fontFace: titleFont,
    fontSize: titleFontSize,
    color: titleColor,
    bold: true,
    align: 'center'
  });

  addText(slide, leftTitle, {
    x: 0.5,
    y: 1.75,
    w: 4.25,
    h: 0.5,
    fontFace: titleFont,
    fontSize: headingFontSize,
    color: headingColor,
    bold: true,
    align: 'left'
  });

  addText(slide, rightTitle, {
    x: 4.75,
    y: 1.75,
    w: 4.25,
    h: 0.5,
    fontFace: titleFont,
    fontSize: headingFontSize,
    color: headingColor,
    bold: true,
    align: 'left'
  });

  addText(slide, formatList(leftItems), {
    x: 0.5,
    y: 2.4,
    w: 4.25,
    h: 4.75,
    fontFace: bodyFont,
    fontSize: bodyFontSize,
    color: textColor,
    align: 'left'
  });

  addText(slide, formatList(rightItems), {
    x: 4.75,
    y: 2.4,
    w: 4.25,
    h: 4.75,
    fontFace: bodyFont,
    fontSize: bodyFontSize,
    color: textColor,
    align: 'left'
  });

  return slide;
}

/**
 * Comparison layout compatibility class.
 */
class ComparisonLayout {
  constructor() {
    this.name = 'comparison';
    this.elements = [];
  }

  applyTo(slide) {
    slide.setLayout(this);
  }
}

module.exports = {
  ComparisonLayout,
  createComparisonSlide
};
