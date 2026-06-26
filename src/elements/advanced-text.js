'use strict';

/**
 * Advanced text formatting functions for PowerPoint presentations
 */

/**
 * Apply strikethrough formatting to text
 * @param {Object} slide - The slide to add text to
 * @param {string} text - The text content
 * @param {Object} options - Text formatting options
 * @returns {Object} The added text object with strikethrough
 */
function addStrikethroughText(slide, text, options = {}) {
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
    strike = true, // Add strikethrough
    align = 'left',
    valign = 'top',
    margin = 0
  } = options;

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
    strike: strike,
    align: align,
    valign: valign,
    margin: margin
  };

  return slide.addText(text, textOptions);
}

/**
 * Add hyperlink to text
 * @param {Object} slide - The slide to add text to
 * @param {string} text - The text content
 * @param {string} url - The URL to link to
 * @param {Object} options - Text formatting options
 * @returns {Object} The added text object with hyperlink
 */
function addHyperlinkText(slide, text, url, options = {}) {
  const {
    x = 0,
    y = 0,
    w = undefined,
    h = undefined,
    fontFace = 'Arial',
    fontSize = 12,
    color = '#0000FF', // Default blue for links
    bold = false,
    italic = false,
    underline = true, // Underline for links
    align = 'left',
    valign = 'top',
    margin = 0
  } = options;

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
    margin: margin,
    hyperlink: url // Add hyperlink property
  };

  return slide.addText(text, textOptions);
}

/**
 * Add rotated text
 * @param {Object} slide - The slide to add text to
 * @param {string} text - The text content
 * @param {number} rotation - Rotation angle in degrees
 * @param {Object} options - Text formatting options
 * @returns {Object} The added text object with rotation
 */
function addRotatedText(slide, text, rotation, options = {}) {
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
    margin: margin,
    rotation: rotation // Add rotation property
  };

  return slide.addText(text, textOptions);
}

/**
 * Add vertical text
 * @param {Object} slide - The slide to add text to
 * @param {string} text - The text content
 * @param {Object} options - Text formatting options
 * @returns {Object} The added text object with vertical orientation
 */
function addVerticalText(slide, text, options = {}) {
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

  // Vertical text is typically achieved by rotating 90 degrees
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
    margin: margin,
    rotation: 90 // Add 90-degree rotation for vertical text
  };

  return slide.addText(text, textOptions);
}

/**
 * Add text with autofit
 * @param {Object} slide - The slide to add text to
 * @param {string} text - The text content
 * @param {Object} options - Text formatting options
 * @returns {Object} The added text object with autofit
 */
function addAutofitText(slide, text, options = {}) {
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
    margin = 0,
    autofit = true // Enable autofit
  } = options;

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
    margin: margin,
    autofit: autofit
  };

  return slide.addText(text, textOptions);
}

/**
 * Add text with word wrapping
 * @param {Object} slide - The slide to add text to
 * @param {string} text - The text content
 * @param {Object} options - Text formatting options
 * @returns {Object} The added text object with word wrapping
 */
function addWrappedText(slide, text, options = {}) {
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
    margin = 0,
    wrap = true // Enable word wrapping
  } = options;

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
    margin: margin,
    wrap: wrap
  };

  return slide.addText(text, textOptions);
}

/**
 * Add paragraph with formatting
 * @param {Object} slide - The slide to add text to
 * @param {string} text - The text content
 * @param {Object} options - Text formatting options
 * @returns {Object} The added text object with paragraph formatting
 */
function addParagraphText(slide, text, options = {}) {
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
    margin = 0,
    spacingBefore = 0,
    spacingAfter = 0,
    lineSpacing = 1.0
  } = options;

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
    margin: margin,
    spacingBefore: spacingBefore,
    spacingAfter: spacingAfter,
    lineSpacing: lineSpacing
  };

  return slide.addText(text, textOptions);
}

/**
 * Add bullet list
 * @param {Object} slide - The slide to add text to
 * @param {Array} items - Array of bullet items
 * @param {Object} options - Text formatting options
 * @returns {Object} The added text object with bullet list
 */
function addBulletList(slide, items, options = {}) {
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
    margin = 0,
    bulletChar = '•',
    bulletIndent = 0.25
  } = options;

  // Add each bullet item separately
  const results = [];
  let currentY = y;
  
  items.forEach((item, index) => {
    const bulletText = `${bulletChar} ${item}`;
    const textOptions = {
      text: bulletText,
      x: x + (index > 0 ? bulletIndent : 0),
      y: currentY,
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

    const result = slide.addText(bulletText, textOptions);
    results.push(result);
    currentY += 0.5; // Adjust spacing between items
  });

  return results;
}

/**
 * Add numbered list
 * @param {Object} slide - The slide to add text to
 * @param {Array} items - Array of list items
 * @param {Object} options - Text formatting options
 * @returns {Object} The added text object with numbered list
 */
function addNumberedList(slide, items, options = {}) {
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
    margin = 0,
    startNumber = 1,
    bulletIndent = 0.25
  } = options;

  // Add each numbered item separately
  const results = [];
  let currentY = y;
  
  items.forEach((item, index) => {
    const number = startNumber + index;
    const numberedText = `${number}. ${item}`;
    const textOptions = {
      text: numberedText,
      x: x + (index > 0 ? bulletIndent : 0),
      y: currentY,
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

    const result = slide.addText(numberedText, textOptions);
    results.push(result);
    currentY += 0.5; // Adjust spacing between items
  });

  return results;
}

/**
 * Add multi-level bullet list
 * @param {Object} slide - The slide to add text to
 * @param {Array} items - Array of bullet items with nesting levels
 * @param {Object} options - Text formatting options
 * @returns {Object} The added text object with multi-level bullet list
 */
function addMultiLevelBulletList(slide, items, options = {}) {
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
    margin = 0,
    bulletChars = ['•', '◦', '▪'],
    bulletIndent = 0.25
  } = options;

  // Add each item with appropriate indentation based on level
  const results = [];
  let currentY = y;
  
  items.forEach((item, index) => {
    const level = item.level || 0;
    const bulletChar = bulletChars[level % bulletChars.length];
    const indent = level * bulletIndent;
    const bulletText = `${bulletChar} ${item.text}`;
    
    const textOptions = {
      text: bulletText,
      x: x + indent,
      y: currentY,
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

    const result = slide.addText(bulletText, textOptions);
    results.push(result);
    currentY += 0.5; // Adjust spacing between items
  });

  return results;
}

/**
 * Search text in a presentation
 * @param {Object} presentation - The presentation object
 * @param {string} searchText - Text to search for
 * @returns {Array} Array of matching text objects
 */
function searchText(presentation, searchText) {
  const results = [];
  
  // Iterate through slides and text elements
  presentation.slides.forEach(slide => {
    if (slide.texts) {
      slide.texts.forEach(textObj => {
        if (textObj.text && textObj.text.includes(searchText)) {
          results.push({
            slide: slide.slideNumber,
            text: textObj.text,
            position: {
              x: textObj.x,
              y: textObj.y
            }
          });
        }
      });
    }
  });
  
  return results;
}

/**
 * Replace text using regex in a presentation
 * @param {Object} presentation - The presentation object
 * @param {RegExp} searchPattern - Regex pattern to search for
 * @param {string} replacement - Replacement text
 * @returns {number} Number of replacements made
 */
function replaceTextWithRegex(presentation, searchPattern, replacement) {
  let count = 0;
  
  // Iterate through slides and text elements
  presentation.slides.forEach(slide => {
    if (slide.texts) {
      slide.texts.forEach(textObj => {
        if (textObj.text && searchPattern.test(textObj.text)) {
          textObj.text = textObj.text.replace(searchPattern, replacement);
          count++;
        }
      });
    }
  });
  
  return count;
}

module.exports = {
  addStrikethroughText,
  addHyperlinkText,
  addRotatedText,
  addVerticalText,
  addAutofitText,
  addWrappedText,
  addParagraphText,
  addBulletList,
  addNumberedList,
  addMultiLevelBulletList,
  searchText,
  replaceTextWithRegex
};