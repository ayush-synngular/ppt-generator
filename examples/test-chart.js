'use strict';

const { Presentation, addBarChart, addLineChart, addPieChart } = require('../src/index');

async function testChart() {
  try {
    const ppt = new Presentation('Chart Test Presentation');

    const barSlide = ppt.addSlide();
    barSlide.addText('Bar Chart Test', {
      x: 0.5,
      y: 0.5,
      w: 9,
      h: 1,
      fontFace: 'Arial',
      fontSize: 28,
      bold: true,
      align: 'center'
    });

    addBarChart(barSlide, [
      { category: 'Q1', value: 12000 },
      { category: 'Q2', value: 18000 },
      { category: 'Q3', value: 25000 },
      { category: 'Q4', value: 30000 }
    ], {
      x: 0.5,
      y: 1.5,
      w: 9,
      h: 5,
      title: 'Quarterly Revenue',
      showLegend: false,
      showValue: true,
      showCategoryName: true
    });

    const lineSlide = ppt.addSlide();
    lineSlide.addText('Line Chart Test', {
      x: 0.5,
      y: 0.5,
      w: 9,
      h: 1,
      fontFace: 'Arial',
      fontSize: 28,
      bold: true,
      align: 'center'
    });

    addLineChart(lineSlide, [
      { category: 'Jan', value: 10 },
      { category: 'Feb', value: 25 },
      { category: 'Mar', value: 35 },
      { category: 'Apr', value: 45 },
      { category: 'May', value: 55 },
      { category: 'Jun', value: 60 }
    ], {
      x: 0.5,
      y: 1.5,
      w: 9,
      h: 5,
      title: 'Monthly Growth',
      showLegend: false,
      showValue: true,
      showCategoryName: true
    });

    const pieSlide = ppt.addSlide();
    pieSlide.addText('Pie Chart Test', {
      x: 0.5,
      y: 0.5,
      w: 9,
      h: 1,
      fontFace: 'Arial',
      fontSize: 28,
      bold: true,
      align: 'center'
    });

    addPieChart(pieSlide, [
      { category: 'Product A', value: 40 },
      { category: 'Product B', value: 30 },
      { category: 'Product C', value: 20 },
      { category: 'Product D', value: 10 }
    ], {
      x: 1.5,
      y: 1.75,
      w: 7,
      h: 4.5,
      title: 'Market Share',
      showLegend: true,
      showValue: true,
      showCategoryName: true
    });

    await ppt.save('test-chart.pptx');
    console.log('Presentation saved successfully as test-chart.pptx');
  } catch (error) {
    console.error('Error creating chart presentation:', error);
  }
}

// Run the test
testChart();
