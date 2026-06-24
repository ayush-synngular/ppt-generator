'use strict';

const { addText } = require('../elements/text');

function getThemeValue(theme, getter, fallback) {
  if (theme && typeof theme[getter] === 'function') {
    const value = theme[getter](fallback.key);
    return value !== undefined ? value : fallback.default;
  }

  return fallback.default;
}

function normalizeContent(content) {
  if (typeof content === 'string') {
    return content;
  }

  if (Array.isArray(content)) {
    return content.join('\n\n');
  }

  return content !== undefined ? String(content) : '';
}

/**
 * Create a two-column slide with a title and separate left/right content blocks.
 * @param {Object} presentation - PPT generator presentation.
 * @param {Object} options - Layout settings.
 * @param {string} options.title - Slide title.
 * @param {string|Array} options.leftContent - Left panel content.
 * @param {string|Array} options.rightContent - Right panel content.
 * @param {Object} options.theme - Theme instance for styling.
 * @returns {Object} The created slide.
 */
function createTwoColumnSlide(presentation, options = {}) {
  const {
    title = 'Two Column Layout',
    leftContent = '',
    rightContent = '',
    theme
  } = options;

  const slide = presentation.addSlide();

  const titleFont = getThemeValue(theme, 'getFont', { key: 'titleFont', default: 'Arial Black' });
  const titleFontSize = getThemeValue(theme, 'getFontSize', { key: 'titleFontSize', default: 30 });
  const titleColor = getThemeValue(theme, 'getColor', { key: 'primaryColor', default: '#1F4E79' });
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

  addText(slide, normalizeContent(leftContent), {
    x: 0.5,
    y: 1.75,
    w: 4.25,
    h: 5,
    fontFace: bodyFont,
    fontSize: bodyFontSize,
    color: textColor,
    align: 'left'
  });

  addText(slide, normalizeContent(rightContent), {
    x: 4.75,
    y: 1.75,
    w: 4.25,
    h: 5,
    fontFace: bodyFont,
    fontSize: bodyFontSize,
    color: textColor,
    align: 'left'
  });

  return slide;
}

/**
 * Two column layout compatibility class.
 */
class TwoColumnLayout {
  constructor() {
    this.name = 'two-column';
    this.columns = 2;
    this.elements = [];
  }

  applyTo(slide) {
    slide.setLayout(this);
  }
}

module.exports = {
  TwoColumnLayout,
  createTwoColumnSlide
};
