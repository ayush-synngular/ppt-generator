'use strict';

const { Presentation, Theme, createTitleSlide, createTwoColumnSlide, createComparisonSlide } = require('../src/index');

async function testLayouts() {
  try {
    const theme = new Theme({
      primaryColor: '#1E70BF',
      secondaryColor: '#7F8C8D',
      accentColor: '#F39C12',
      backgroundColor: '#FFFFFF',
      textColor: '#1F2937',
      titleFont: 'Calibri Light',
      bodyFont: 'Calibri',
      titleFontSize: 32,
      bodyFontSize: 18,
      slideMargin: 0.75
    });

    const ppt = new Presentation('Layout Test Presentation');

    createTitleSlide(ppt, {
      title: 'Title Slide',
      subtitle: 'Welcome to the layout system demo',
      theme
    });

    createTwoColumnSlide(ppt, {
      title: 'Two Column Layout',
      leftContent: ['This is the left column.', 'It supports multiple paragraphs.'],
      rightContent: ['This is the right column.', 'It also supports multiple paragraphs.'],
      theme
    });

    createComparisonSlide(ppt, {
      title: 'Comparison Layout',
      leftTitle: 'Pros',
      rightTitle: 'Cons',
      leftItems: ['Easy to read', 'Consistent styling', 'Simple layout'],
      rightItems: ['May be text-heavy', 'Limited visuals'],
      theme
    });

    await ppt.save('test-layouts.pptx');
    console.log('Presentation saved successfully as test-layouts.pptx');
  } catch (error) {
    console.error('Error creating layout presentation:', error);
  }
}

testLayouts();
