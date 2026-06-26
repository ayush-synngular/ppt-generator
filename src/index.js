#!/usr/bin/env node
'use strict';

// Core classes
const Presentation = require('./core/Presentation');
const Slide = require('./core/Slide');
const Theme = require('./core/Theme');
const SlideBuilder = require('./core/SlideBuilder');
const TemplateEngine = require('./core/TemplateEngine');
const ValidationEngine = require('./core/ValidationEngine');
const DataBinder = require('./core/DataBinder');
const ExportUtils = require('./utils/export');

// Layouts
const { createTitleSlide } = require('./layouts/title-slide');
const { createTwoColumnSlide } = require('./layouts/two-column');
const { createComparisonSlide } = require('./layouts/comparison');

// Elements - Import all from index.js
const {
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
} = require('./elements/index');

// Individual element imports
const { addTable } = require('./elements/table');
const { addBarChart, addLineChart, addPieChart } = require('./elements/chart');
const { addIcon } = require('./elements/icon');
const { addImage } = require('./elements/image');
const { addRectangle, addCircle, addLine, addArrow } = require('./elements/shape');

// Components
const { renderTitle } = require('./components/title');
const { renderBulletList } = require('./components/bullet-list');
const { renderMetricCard } = require('./components/metric-card');
const { renderChartCard } = require('./components/chart-card');

// Templates
const ExecutiveSummaryTemplate = require('./templates/ExecutiveSummaryTemplate');
const ProjectStatusTemplate = require('./templates/ProjectStatusTemplate');
const QuarterlyBusinessReviewTemplate = require('./templates/QuarterlyBusinessReviewTemplate');

// Utilities
const { hexToRgb } = require('./utils/color');
const { calculateSize } = require('./utils/sizing');
const { calculatePosition } = require('./utils/positioning');

// Constants
const fonts = require('./constants/fonts');
const themes = require('./constants/themes');

// Auto Layout Engine
const AutoLayoutEngine = require('./core/AutoLayoutEngine');

const fs = require('fs');

/**
 * Parse command line arguments
 * @param {string[]} args - Command line arguments
 * @returns {Object} Parsed arguments
 */
function parseArgs(args) {
  const parsed = {
    flags: {},
    positional: []
  };

  for (let i = 0; i < args.length; i++) {
    if (args[i].startsWith('--')) {
      const key = args[i].substring(2);
      const value = args[i + 1];
      if (value && !value.startsWith('--')) {
        parsed.flags[key] = value;
        i++; // Skip next argument
      } else {
        parsed.flags[key] = true;
      }
    } else {
      parsed.positional.push(args[i]);
    }
  }

  return parsed;
}

/**
 * Validate required parameters
 * @param {Object} params - Parameters to validate
 * @param {string[]} required - Required parameter names
 * @returns {boolean} True if all required parameters are present
 */
function validateRequiredParams(params, required) {
  const missing = required.filter(param => !params[param]);
  if (missing.length > 0) {
    console.error(`❌ Missing required parameters: ${missing.join(', ')}`);
    return false;
  }
  return true;
}

function printHelp() {
  console.log('================================================================');
  console.log('🤖 PPT GENERATOR DYNAMIC COMPONENT CLI DIRECTORY');
  console.log('================================================================');
  console.log('Available Commands:');
  console.log('  demo --data <file.json> [--output <file.pptx>]');
  console.log('  inspect <file.pptx>');
  console.log('  create <component-type> [--flags]');
  console.log('  replace "old" "new" <file.pptx> [--output <file.pptx>]');
  console.log('  generate --data <schema.json> [--output <file.pptx>]');
  console.log('  template <template-type> --data <data.json> [--output <file.pptx>]');
  console.log('  bind --template <template.json> --data <data.json> [--output <bound.json>]');
  console.log('  export --input <file.pptx> --format json|xml [--output <file.ext>]');
  console.log('  help');
  console.log('');
  console.log('Component Types:');
  console.log('  text, table, chart, icon, image, shape');
  console.log('');
  console.log('Template Types:');
  console.log('  executive-summary, project-status, quarterly-business-review');
  console.log('');
  console.log('Examples:');
  console.log('  node src/index.js create text --title "Hello World" --content "This is content" --output hello.pptx');
  console.log('  node src/index.js create chart --type bar --data \'[["Jan", 10], ["Feb", 20]]\' --output chart.pptx');
  console.log('  node src/index.js template executive-summary --data data.json --output summary.pptx');
  console.log('================================================================');
}

/**
 * --- Command handlers ---
 * Each handler is isolated so the CLI dispatch table stays simple to read
 * and so each command can also be invoked programmatically if needed.
 */

function runDemo(flags, outputName) {
  if (!validateRequiredParams(flags, ['data'])) return;

  try {
    const deckData = JSON.parse(fs.readFileSync(flags.data, 'utf8'));
    console.log('🚀 Generating showcase demo layout presentation...');

    const presentation = new Presentation('Showcase Demo');
    const activeTheme = new Theme(deckData.theme);

    const dynamicBuilder = new SlideBuilder(presentation, activeTheme);
    dynamicBuilder.build(deckData);

    presentation.save(outputName)
      .then(() => console.log(`🎉 Showcase slide deck compiled successfully: ${outputName}`))
      .catch(err => console.error(`❌ Demo error: ${err.message}`));
  } catch (e) {
    console.error(`❌ Failed to process demo template data file: ${e.message}`);
  }
}

function runInspect(positional) {
  const filePath = positional[1];
  if (!filePath) {
    console.error('❌ Please specify a file path to inspect.');
    process.exit(1);
  }
  console.log(`🔍 Inspecting target presentation file: ${filePath}`);
  Presentation.open(filePath)
    .then((presInstance) => {
      console.log(`Presentation File Base: ${presInstance.title}`);
      console.log(`Slide count: ${presInstance.slides.length}\n`);
      console.log('Slide titles:');
      presInstance.slides.forEach(s => console.log(`  ${s.slideNumber || '?'}: ${s.title}`));
    })
    .catch(err => console.error(`❌ Error inspecting presentation: ${err.message}`));
}

function runCreate(flags, positional, outputName) {
  const subCommand = positional[1];
  if (!subCommand) {
    console.error('❌ Missing component sub-type parameter (text, table, chart, icon, image, shape).');
    process.exit(1);
  }

  const presentation = new Presentation();
  const slide = presentation.addSlide();

  // Text component
  if (subCommand === 'text') {
    const title = flags.title || 'Header Text';
    const content = flags.content || 'Body Content Block';
    const x = parseFloat(flags.x) || 0.5;
    const y = parseFloat(flags.y) || 0.5;
    const fontSize = parseInt(flags.fontSize) || 24;
    const bold = flags.bold === 'true';

    addText(slide, title, { x, y, fontSize, bold, fontFace: 'Arial' });
    addText(slide, content, { x, y: y + 1, fontSize: fontSize - 10 });
  }
  // Table component
  else if (subCommand === 'table') {
    if (!flags.rows) {
      console.error('❌ Missing data array row matrix: --rows');
      process.exit(1);
    }
    try {
      const rows = JSON.parse(flags.rows);
      const x = parseFloat(flags.x) || 0.5;
      const y = parseFloat(flags.y) || 1.25;
      const w = parseFloat(flags.w) || 9;

      addTable(slide, rows, { x, y, w, boldHeader: true });
    } catch (e) {
      console.error('❌ Invalid JSON matrix syntax provided.');
    }
  }
  // Chart component
  else if (subCommand === 'chart') {
    if (!flags.data) {
      console.error('❌ Missing data array matrix: --data');
      process.exit(1);
    }
    try {
      const rawChartData = JSON.parse(flags.data);
      const type = flags.type || 'bar';
      const title = flags.title || 'Chart Title';
      const x = parseFloat(flags.x) || 0.5;
      const y = parseFloat(flags.y) || 0.5;
      const w = parseFloat(flags.w) || 8;
      const h = parseFloat(flags.h) || 4;

      if (type === 'line') addLineChart(slide, rawChartData, { title, x, y, w, h });
      else if (type === 'pie') addPieChart(slide, rawChartData, { title, x, y, w, h });
      else addBarChart(slide, rawChartData, { title, x, y, w, h });
    } catch (e) {
      console.error('❌ Invalid JSON chart structure.');
    }
  }
  // Icon component
  else if (subCommand === 'icon') {
    const name = flags.name || 'star';
    const x = parseFloat(flags.x) || 4.25;
    const y = parseFloat(flags.y) || 2.0;

    addIcon(slide, name, { x, y });
  }
  // Image component
  else if (subCommand === 'image') {
    if (!flags.path) {
      console.error('❌ Missing image location: --path');
      process.exit(1);
    }
    const path = flags.path;
    const x = parseFloat(flags.x) || 1;
    const y = parseFloat(flags.y) || 1;
    const w = parseFloat(flags.w) || 5;
    const h = parseFloat(flags.h) || 4;

    addImage(slide, path, { x, y, w, h });
  }
  // Shape component
  else if (subCommand === 'shape') {
    const shapeType = flags.type || 'rect';
    const x = parseFloat(flags.x) || 1;
    const y = parseFloat(flags.y) || 1;
    const w = parseFloat(flags.w) || 3;
    const h = parseFloat(flags.h) || 2;
    const fillColor = flags.color || '0078D4';

    const opts = { x, y, w, h, fillColor };
    if (shapeType === 'circle') addCircle(slide, opts);
    else if (shapeType === 'line') addLine(slide, { x1: x, y1: y, x2: x + w, y2: y });
    else if (shapeType === 'arrow') addArrow(slide, { x1: x, y1: y, x2: x + w, y2: y });
    else addRectangle(slide, opts);
  }

  presentation.save(outputName)
    .then(() => console.log(`🎉 Component generation successful: ${outputName}`))
    .catch(err => console.error(`❌ Creation failure: ${err.message}`));
}

function runReplace(positional, outputName) {
  const matchPattern = positional[1];
  const replacementText = positional[2];
  const targetSourceFile = positional[3];

  if (!matchPattern || !replacementText || !targetSourceFile) {
    console.error('Usage: node src/index.js replace "Old" "New" file.pptx --output target.pptx');
    process.exit(1);
  }

  Presentation.open(targetSourceFile)
    .then((openedPres) => {
      openedPres.replaceText(matchPattern, replacementText);
      return openedPres.save(outputName);
    })
    .then(() => console.log(`🎉 In-place replacement successful. Saved to: ${outputName}`))
    .catch(err => console.error(`❌ Replace failure: ${err.message}`));
}

function runGenerate(flags, outputName) {
  if (!validateRequiredParams(flags, ['data'])) return;

  try {
    const schemaFilePath = flags.data;
    const configMap = JSON.parse(fs.readFileSync(schemaFilePath, 'utf8'));
    ValidationEngine.validatePresentationConfig(configMap);

    const presentation = new Presentation();
    const appTheme = new Theme(configMap.theme);
    const dynamicBuilder = new SlideBuilder(presentation, appTheme);

    dynamicBuilder.build(configMap);
    presentation.save(outputName)
      .then(() => console.log(`🎉 Automated schema compilation completed successfully. File: ${outputName}`))
      .catch(err => console.error(`❌ Batch loop error: ${err.message}`));
  } catch (err) {
    console.error(`❌ Workspace schema parsing validation exception: ${err.message}`);
  }
}

function runTemplate(flags, positional, outputName) {
  const templateType = positional[1];
  const dataFile = flags.data;

  if (!templateType) {
    console.error('❌ Missing template type. Use: template executive-summary|project-status|quarterly-business-review');
    process.exit(1);
  }

  if (!dataFile) {
    console.error('❌ Missing data file. Use: --data <file.json>');
    process.exit(1);
  }

  try {
    const data = JSON.parse(fs.readFileSync(dataFile, 'utf8'));
    const presentation = new Presentation();
    const theme = new Theme();

    let template;
    switch (templateType) {
      case 'executive-summary':
        template = new ExecutiveSummaryTemplate(presentation, theme);
        break;
      case 'project-status':
        template = new ProjectStatusTemplate(presentation, theme);
        break;
      case 'quarterly-business-review':
        template = new QuarterlyBusinessReviewTemplate(presentation, theme);
        break;
      default:
        console.error('❌ Unsupported template type. Available: executive-summary, project-status, quarterly-business-review');
        process.exit(1);
    }

    template.build(data);
    presentation.save(outputName)
      .then(() => console.log(`🎉 Template generation successful: ${outputName}`))
      .catch(err => console.error(`❌ Template generation error: ${err.message}`));
  } catch (err) {
    console.error(`❌ Template processing error: ${err.message}`);
  }
}

function runBind(flags) {
  const templateFile = flags.template;
  const dataFile = flags.data;

  if (!templateFile || !dataFile) {
    console.error('❌ Missing template or data file. Use: --template <file.json> --data <data.json>');
    process.exit(1);
  }

  try {
    const template = JSON.parse(fs.readFileSync(templateFile, 'utf8'));
    const data = JSON.parse(fs.readFileSync(dataFile, 'utf8'));
    const binder = new DataBinder();
    const boundTemplate = binder.bind(template, data);

    // Save the bound template to a file or display it
    const outputFile = flags.output || 'bound-template.json';
    fs.writeFileSync(outputFile, JSON.stringify(boundTemplate, null, 2));
    console.log(`🎉 Data binding successful. Output saved to: ${outputFile}`);
  } catch (err) {
    console.error(`❌ Data binding error: ${err.message}`);
  }
}

function runExport(flags) {
  const inputFile = flags.input;
  const format = flags.format || 'json';

  if (!inputFile) {
    console.error('❌ Missing input file. Use: --input <file.pptx>');
    process.exit(1);
  }

  Presentation.open(inputFile)
    .then(presentation => {
      let output;
      if (format === 'json') {
        output = ExportUtils.toJson(presentation);
      } else if (format === 'xml') {
        output = ExportUtils.toXml(presentation);
      } else {
        console.error('❌ Unsupported format. Use: json or xml');
        process.exit(1);
        return;
      }

      const outputFile = flags.output || `exported-${Date.now()}.${format}`;
      fs.writeFileSync(outputFile, output);
      console.log(`🎉 Export successful. Output saved to: ${outputFile}`);
    })
    .catch(err => console.error(`❌ Export error: ${err.message}`));
}

/**
 * Main execution function
 * This function handles CLI commands and can be called programmatically
 */
function main() {
  const args = process.argv.slice(2);
  const parsed = parseArgs(args);
  const { flags, positional } = parsed;

  // Default output name
  const outputName = flags.output || 'output.pptx';

  // Handle different commands
  const command = positional[0];

  switch (command) {
    case 'demo':
      runDemo(flags, outputName);
      break;

    case 'inspect':
      runInspect(positional);
      break;

    case 'create':
      runCreate(flags, positional, outputName);
      break;

    case 'replace':
      runReplace(positional, outputName);
      break;

    case 'generate':
      runGenerate(flags, outputName);
      break;

    case 'template':
      runTemplate(flags, positional, outputName);
      break;

    case 'bind':
      runBind(flags);
      break;

    case 'export':
      runExport(flags);
      break;

    case 'help':
    default:
      printHelp();
  }
}

// Execute main function if called directly from the terminal
// (e.g. `node src/index.js <command>` or `./src/index.js <command>` if chmod +x)
if (require.main === module) {
  main();
}

// Export all functionality so this module can also be required programmatically
module.exports = {
  // Core classes
  Presentation,
  Slide,
  Theme,
  SlideBuilder,
  TemplateEngine,
  ValidationEngine,
  DataBinder,
  ExportUtils,

  // Layouts
  createTitleSlide,
  createTwoColumnSlide,
  createComparisonSlide,

  // Elements
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
  replaceTextWithRegex,
  addTable,
  addBarChart,
  addLineChart,
  addPieChart,
  addIcon,
  addImage,
  addRectangle,
  addCircle,
  addLine,
  addArrow,

  // Components
  renderTitle,
  renderBulletList,
  renderMetricCard,
  renderChartCard,

  // Templates
  ExecutiveSummaryTemplate,
  ProjectStatusTemplate,
  QuarterlyBusinessReviewTemplate,

  // Utilities
  hexToRgb,
  calculateSize,
  calculatePosition,

  // Constants
  fonts,
  themes,

  // Auto Layout Engine
  AutoLayoutEngine,

  // Helper functions
  parseArgs,
  validateRequiredParams,

  // Main function
  main
};