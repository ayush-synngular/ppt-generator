'use strict';

const {
  Presentation,
  Theme,
  TemplateEngine
} = require('../src/index');
const ExecutiveSummaryTemplate = require('../src/templates/ExecutiveSummaryTemplate');

async function testDataBinding() {
  try {
    const theme = new Theme({
      primaryColor: '#0B5394',
      secondaryColor: '#D9E1F2',
      accentColor: '#F1C232',
      backgroundColor: '#FFFFFF',
      textColor: '#1F2937',
      titleFont: 'Calibri Light',
      bodyFont: 'Calibri',
      titleFontSize: 32,
      bodyFontSize: 18,
      slideMargin: 0.75
    });

    const ppt = new Presentation('Data Binding Demo');
    const engine = new TemplateEngine(ppt, theme);

    engine.registerTemplate('executive-summary', new ExecutiveSummaryTemplate(ppt, theme));

    const data = {
      company: 'ABC Corp',
      revenue: 1200000,
      title: '{{company}} QBR',
      subtitle: 'Revenue: ${{revenue}}',
      summary: ['{{company}} experienced strong demand.', 'Revenue grew to ${{revenue}} this quarter.'],
      metrics: [
        { title: 'Revenue', value: '${{revenue}}' },
        { title: 'Customers', value: '{{metrics[0].value}}' }
      ],
      recommendations: ['Continue expansion', 'Optimize process']
    };

    engine.build('executive-summary', data);

    await ppt.save('test-data-binding.pptx');
    console.log('Presentation saved successfully as test-data-binding.pptx');
  } catch (error) {
    console.error('Error creating data binding presentation:', error);
  }
}

testDataBinding();
