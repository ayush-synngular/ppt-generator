'use strict';

const { Presentation, Theme, SlideBuilder } = require('../src/index');

async function testSlideBuilder() {
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

    const ppt = new Presentation('Slide Builder Demo');
    const builder = new SlideBuilder(ppt, theme);

    const config = {
      slides: [
        {
          type: 'title-slide',
          title: 'Welcome',
          subtitle: 'JSON-driven Slide Builder Demo'
        },
        {
          type: 'two-column',
          title: 'Overview',
          leftContent: ['Point 1', 'Point 2'],
          rightContent: ['Point A', 'Point B']
        },
        {
          type: 'comparison',
          title: 'Comparison',
          leftTitle: 'Pros',
          rightTitle: 'Cons',
          leftItems: ['Simple JSON input', 'Reusable layout helpers'],
          rightItems: ['Less visual control', 'Requires valid config']
        }
      ]
    };

    builder.build(config);

    await ppt.save('test-slide-builder.pptx');
    console.log('Presentation saved successfully as test-slide-builder.pptx');
  } catch (error) {
    console.error('Error creating slide builder presentation:', error);
  }
}

testSlideBuilder();
