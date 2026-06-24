'use strict';

const { Presentation, Theme, renderTitle, renderBulletList, renderMetricCard, renderChartCard } = require('../src/index');

async function testComponents() {
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

    const ppt = new Presentation('Component System Demo');

    const slide1 = ppt.addSlide();
    renderTitle(slide1, {
      text: 'Component System',
      x: 0.5,
      y: 0.5,
      w: 9,
      h: 1
    }, theme);

    renderBulletList(slide1, {
      items: ['Reusable components', 'Theme-aware styling', 'Layout integration'],
      x: 0.5,
      y: 1.75,
      w: 8,
      h: 4
    }, theme);

    const slide2 = ppt.addSlide();
    renderMetricCard(slide2, {
      title: 'Revenue',
      value: '$42.7K',
      x: 0.5,
      y: 0.75,
      w: 4,
      h: 2.5
    }, theme);
    renderMetricCard(slide2, {
      title: 'Growth',
      value: '18%',
      x: 4.75,
      y: 0.75,
      w: 4,
      h: 2.5
    }, theme);
    renderMetricCard(slide2, {
      title: 'Customers',
      value: '1.2K',
      x: 0.5,
      y: 3.5,
      w: 4,
      h: 2.5
    }, theme);
    renderMetricCard(slide2, {
      title: 'Retention',
      value: '89%',
      x: 4.75,
      y: 3.5,
      w: 4,
      h: 2.5
    }, theme);

    const slide3 = ppt.addSlide();
    renderChartCard(slide3, {
      title: 'Sales Performance',
      chartType: 'bar',
      data: [
        { category: 'Jan', value: 10 },
        { category: 'Feb', value: 20 },
        { category: 'Mar', value: 25 },
        { category: 'Apr', value: 35 }
      ],
      x: 0.5,
      y: 0.5,
      w: 8,
      h: 5
    }, theme);

    await ppt.save('test-components.pptx');
    console.log('Presentation saved successfully as test-components.pptx');
  } catch (error) {
    console.error('Error creating components presentation:', error);
  }
}

testComponents();
