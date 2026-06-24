'use strict';

const { Presentation, Theme, addText } = require('../src/index');

async function testTheme() {
  try {
    const theme = new Theme({
      primaryColor: '#1E70BF',
      secondaryColor: '#7F8C8D',
      accentColor: '#F39C12',
      backgroundColor: '#FFFFFF',
      textColor: '#1F2937',
      titleFont: 'Calibri Light',
      bodyFont: 'Calibri',
      titleFontSize: 34,
      bodyFontSize: 18,
      slideMargin: 0.75
    });

    console.log('Theme values:');
    console.log('primaryColor:', theme.getColor('primaryColor'));
    console.log('secondaryColor:', theme.getColor('secondaryColor'));
    console.log('accentColor:', theme.getColor('accentColor'));
    console.log('backgroundColor:', theme.getColor('backgroundColor'));
    console.log('textColor:', theme.getColor('textColor'));
    console.log('titleFont:', theme.getFont('titleFont'));
    console.log('bodyFont:', theme.getFont('bodyFont'));
    console.log('titleFontSize:', theme.getFontSize('titleFontSize'));
    console.log('bodyFontSize:', theme.getFontSize('bodyFontSize'));
    console.log('slideMargin:', theme.getSlideMargin());

    const ppt = new Presentation('Theme Test Presentation');
    const slide = ppt.addSlide();

    const margin = theme.getSlideMargin();

    addText(slide, 'Theme Engine Test', {
      x: margin,
      y: margin,
      w: 9 - margin * 2,
      h: 1,
      fontFace: theme.getFont('titleFont'),
      fontSize: theme.getFontSize('titleFontSize'),
      color: theme.getColor('primaryColor'),
      bold: true,
      align: 'center'
    });

    addText(slide, 'This presentation demonstrates applying theme settings to text and layout. The Theme engine provides colors, fonts, sizes, and margins.', {
      x: margin,
      y: margin + 1.2,
      w: 9 - margin * 2,
      h: 3,
      fontFace: theme.getFont('bodyFont'),
      fontSize: theme.getFontSize('bodyFontSize'),
      color: theme.getColor('textColor'),
      align: 'left'
    });

    await ppt.save('test-theme.pptx');
    console.log('Presentation saved successfully as test-theme.pptx');
  } catch (error) {
    console.error('Error creating theme presentation:', error);
  }
}

testTheme();
