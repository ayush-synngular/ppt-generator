# Chart Functionality Test Commands

This document lists the test commands that would be used to manually test the new chart functionalities.

## 1. Area Chart Test Command
```bash
node src/index.js create chart --type area --data '[{"category":"Q1","value":12000},{"category":"Q2","value":18000},{"category":"Q3","value":25000}]' --title "Area Chart Test" --output area-chart.pptx
```
This would be the command to test area charts after implementing the `addAreaChart` function. Data should be passed as a JSON string via the `--data` flag.

## 2. Donut Chart Test Command
```bash
node examples/test-chart.js
```
This would be extended to include a test case for donut charts after implementing the `addDonutChart` function.

## 3. Scatter Chart Test Command
```bash
node examples/test-chart.js
```
This would be extended to include a test case for scatter charts after implementing the `addScatterChart` function.

## 4. Bubble Chart Test Command
```bash
node examples/test-chart.js
```
This would be extended to include a test case for bubble charts after implementing the `addBubbleChart` function.

## 5. Histogram Test Command
```bash
node examples/test-chart.js
```
This would be extended to include a test case for histograms after implementing the `addHistogram` function.

## 6. Waterfall Chart Test Command
```bash
node examples/test-chart.js
```
This would be extended to include a test case for waterfall charts after implementing the `addWaterfallChart` function.

## 7. Funnel Chart Test Command
```bash
node examples/test-chart.js
```
This would be extended to include a test case for funnel charts after implementing the `addFunnelChart` function.

## 8. Radar Chart Test Command
```bash
node examples/test-chart.js
```
This would be extended to include a test case for radar charts after implementing the `addRadarChart` function.

## 9. Treemap Test Command
```bash
node examples/test-chart.js
```
This would be extended to include a test case for treemaps after implementing the `addTreemap` function.

## 10. Combo Charts Test Command
```bash
node examples/test-chart.js
```
This would be extended to include a test case for combo charts after implementing the `addComboChart` function.

## 11. Secondary Axis Test Command
```bash
node examples/test-chart.js
```
This would be extended to include a test case for charts with secondary axes after implementing the `addChartWithSecondaryAxis` function.

## 12. Chart Updates Test Command
```bash
node examples/test-chart.js
```
This would be extended to include a test case for updating existing charts after implementing the `updateChart` function.

## General Testing Approach

To test these chart functionalities, you would typically:

1. Create a test presentation with existing chart examples
2. Add one of the new chart types to the presentation
3. Save the presentation and verify the results visually
4. Test edge cases like invalid data formats, empty datasets, etc.
5. Test with various data configurations and chart options

## Sample Test Structure

For testing purposes, you could create a new test file like `test-charts-advanced.js`:

```javascript
const { Presentation } = require('../src/index');
const { addBarChart, addLineChart, addPieChart } = require('../src/elements/chart');

async function testAdvancedCharts() {
  try {
    const ppt = new Presentation('Advanced Charts Test');
    const slide = ppt.addSlide();
    
    // Create a basic bar chart first (existing functionality)
    addBarChart(slide, [
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
    
    // Test new chart functionality (these would be implemented)
    // addAreaChart(slide, data, options);
    // addDonutChart(slide, data, options);
    // addScatterChart(slide, data, options);
    // etc.
    
    await ppt.save('test-advanced-charts.pptx');
    console.log('Advanced charts test saved successfully');
  } catch (error) {
    console.error('Error in advanced charts test:', error);
  }
}

testAdvancedCharts();
```

Each new function would be tested individually with appropriate parameters and validation, similar to how existing charts are tested.