'use strict';

const { addText } = require('../elements/text');

function getThemeValue(theme, getter, fallback) {
  if (theme && typeof theme[getter] === 'function') {
    const value = theme[getter](fallback.key);
    return value !== undefined ? value : fallback.default;
  }

  return fallback.default;
}

/**
 * Create a title slide using the provided presentation and theme.
 * @param {Object} presentation - PPT generator presentation.
 * @param {Object} options - Layout options.
 * @param {string} options.title - Slide title text.
 * @param {string} options.subtitle - Slide subtitle text.
 * @param {Object} options.theme - Theme instance for styling.
 * @returns {Object} The created slide.
 */
function createTitleSlide(presentation, options = {}) {
  const {
    title = 'Title Slide',
    subtitle = '',
    theme
  } = options;

  const slide = presentation.addSlide();

  const titleFont = getThemeValue(theme, 'getFont', { key: 'titleFont', default: 'Arial Black' });
  const titleFontSize = getThemeValue(theme, 'getFontSize', { key: 'titleFontSize', default: 36 });
  const titleColor = getThemeValue(theme, 'getColor', { key: 'primaryColor', default: '#1F4E79' });
  const subtitleFont = getThemeValue(theme, 'getFont', { key: 'bodyFont', default: 'Arial' });
  const subtitleFontSize = getThemeValue(theme, 'getFontSize', { key: 'bodyFontSize', default: 18 });
  const subtitleColor = getThemeValue(theme, 'getColor', { key: 'secondaryColor', default: '#7A7A7A' });

  if (theme && typeof theme.getColor === 'function') {
    const backgroundColor = theme.getColor('backgroundColor');
    if (backgroundColor) {
      slide.background = { color: backgroundColor };
    }
  }

  addText(slide, title, {
    x: 0.5,
    y: 1.5,
    w: 9,
    h: 1.25,
    fontFace: titleFont,
    fontSize: titleFontSize,
    color: titleColor,
    bold: true,
    align: 'center'
  });

  if (subtitle) {
    addText(slide, subtitle, {
      x: 1,
      y: 3,
      w: 8,
      h: 1,
      fontFace: subtitleFont,
      fontSize: subtitleFontSize,
      color: subtitleColor,
      align: 'center'
    });
  }

  return slide;
}

/**
 * Title slide layout compatibility class.
 */
class TitleSlideLayout {
  constructor() {
    this.name = 'title-slide';
    this.elements = [];
  }

  applyTo(slide) {
    slide.setLayout(this);
  }
}

module.exports = {
  TitleSlideLayout,
  createTitleSlide
};
