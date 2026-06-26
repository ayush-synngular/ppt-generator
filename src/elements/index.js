'use strict';

// Re-export all text-related elements including advanced text functions
const { addText } = require('./text');
const { 
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
} = require('./advanced-text');

module.exports = {
  addText,
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