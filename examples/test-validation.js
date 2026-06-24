'use strict';

const fs = require('fs');
const ValidationEngine = require('../src/core/ValidationEngine');

const engine = new ValidationEngine();

const validConfig = {
  slides: [
    {
      type: 'title',
      components: [
        { type: 'title', text: 'Welcome' }
      ]
    }
  ]
};

const invalidSlideTypeConfig = {
  slides: [
    {
      type: 'invalid-slide',
      components: []
    }
  ]
};

const invalidCoordinatesConfig = {
  slides: [
    {
      type: 'title',
      bounds: { x: -1, y: 0, w: 10, h: 5 },
      components: [
        { type: 'text', text: 'Coordinates', bounds: { x: 0, y: 0, w: 0, h: 2 } }
      ]
    }
  ]
};

const missingImagePathConfig = {
  slides: [
    {
      type: 'image',
      components: [
        { type: 'image', image: { path: '' } }
      ]
    }
  ]
};

const invalidChartTypeConfig = {
  slides: [
    {
      type: 'chart',
      components: [
        { type: 'chart', chart: { type: 'scatter', data: [] } }
      ]
    }
  ]
};

const cases = [
  { name: 'Valid configuration', config: validConfig },
  { name: 'Invalid slide type', config: invalidSlideTypeConfig },
  { name: 'Invalid coordinates', config: invalidCoordinatesConfig },
  { name: 'Missing image path', config: missingImagePathConfig },
  { name: 'Invalid chart type', config: invalidChartTypeConfig }
];

function runValidation() {
  const report = {
    valid: true,
    errors: [],
    warnings: []
  };

  cases.forEach(({ name, config }) => {
    const result = engine.validatePresentation(config);
    console.log(`\n=== ${name} ===`);
    console.log(`valid: ${result.valid}`);
    console.log('errors:', result.errors);
    console.log('warnings:', result.warnings);

    if (!result.valid) {
      report.valid = false;
    }
    report.errors.push({ name, errors: result.errors });
    report.warnings.push({ name, warnings: result.warnings });
  });

  fs.writeFileSync('validation-report.json', JSON.stringify(report, null, 2), 'utf8');
  console.log('\nWrote validation-report.json');
}

runValidation();
