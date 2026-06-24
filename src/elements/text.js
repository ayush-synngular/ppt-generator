'use strict';

/**
 * Add text to a slide
 * @param {Object} slide - The slide to add text to
 * @param {string} text - The text content
 * @param {Object} options - Text formatting options
 * @returns {Object} The added text object
 */
function addText(slide, text, options = {}) {
  // Extract options with defaults
  const {
    x = 0,
    y = 0,
    w = undefined,
    h = undefined,
    fontFace = 'Arial',
    fontSize = 12,
    color = '#000000',
    bold = false,
    italic = false,
    underline = false,
    align = 'left',
    valign = 'top',
    margin = 0
  } = options;

  // Create text object with all specified properties
  const textOptions = {
    text: text,
    x: x,
    y: y,
    w: w,
    h: h,
    fontFace: fontFace,
    fontSize: fontSize,
    color: color,
    bold: bold,
    italic: italic,
    underline: underline,
    align: align,
    valign: valign,
    margin: margin
  };

  // Add text to slide using pptxgenjs
  return slide.addText(text, textOptions);
}

module.exports = {
  addText
};
