'use strict';

/**
 * Main entry point for the PPT Generator
 */

const path = require('path');
const fs = require('fs');
const pkg = require('../package.json');

// Core modules
const Presentation = require('./core/Presentation');
const Slide = require('./core/Slide');
const Theme = require('./core/Theme');
const SlideBuilder = require('./core/SlideBuilder');
const TemplateEngine = require('./core/TemplateEngine');
const DataBinder = require('./core/DataBinder');
const AutoLayoutEngine = require('./core/AutoLayoutEngine');
const ValidationEngine = require('./core/ValidationEngine');

// Element modules
const { addText } = require('./elements/text');
const { addImage } = require('./elements/image');
const { addTable } = require('./elements/table');
const { ChartElement, addBarChart, addLineChart, addPieChart } = require('./elements/chart');
const { addRectangle, addCircle, addLine, addArrow } = require('./elements/shape');
const { addIcon, IconElement } = require('./elements/icon');

// Component modules
const { renderTitle } = require('./components/title');
const { renderBulletList } = require('./components/bullet-list');
const { renderMetricCard } = require('./components/metric-card');
const { renderChartCard } = require('./components/chart-card');

// Layout modules
const { TitleSlideLayout, createTitleSlide } = require('./layouts/title-slide');
const { TwoColumnLayout, createTwoColumnSlide } = require('./layouts/two-column');
const { ComparisonLayout, createComparisonSlide } = require('./layouts/comparison');

// Import modifiers
const PptModifier = require('./import/PptModifier');

// Utility modules
const ColorUtils = require('./utils/color');
const SizingUtils = require('./utils/sizing');
const PositioningUtils = require('./utils/positioning');
const ExportUtils = require('./utils/export');

// Constants
const Fonts = require('./constants/fonts');
const Themes = require('./constants/themes');

// ---------------------------------------------------------------------------
// Utility namespaces exposed to the CLI.
// Each entry maps a CLI-friendly namespace name to the underlying module.
// This is used both for the dedicated "util" command and for generic
// function introspection/help.
// ---------------------------------------------------------------------------
const UTIL_NAMESPACES = {
  color: ColorUtils,
  sizing: SizingUtils,
  positioning: PositioningUtils,
  export: ExportUtils
};

// ---------------------------------------------------------------------------
// CLI argument coercion helpers
// ---------------------------------------------------------------------------

/**
 * Attempt to coerce a raw CLI string argument into a more useful JS value.
 * - "true"/"false" -> boolean
 * - numeric strings -> number
 * - JSON-looking strings ({...} or [...]) -> parsed JSON
 * - everything else -> left as string
 */
function coerceArg(raw) {
  if (raw === undefined) return undefined;

  if (raw === 'true') return true;
  if (raw === 'false') return false;
  if (raw === 'null') return null;
  if (raw === 'undefined') return undefined;

  if (/^-?\d+(\.\d+)?$/.test(raw)) {
    return Number(raw);
  }

  const trimmed = raw.trim();
  if (
    (trimmed.startsWith('{') && trimmed.endsWith('}')) ||
    (trimmed.startsWith('[') && trimmed.endsWith(']'))
  ) {
    try {
      return JSON.parse(trimmed);
    } catch (e) {
      // fall through and return raw string if JSON parsing fails
    }
  }

  return raw;
}

function coerceArgs(rawArgs) {
  return rawArgs.map(coerceArg);
}

/**
 * Pretty-print a function's return value to stdout.
 */
function printResult(result) {
  if (result === undefined) {
    console.log('(no return value)');
    return;
  }
  if (typeof result === 'object') {
    console.log(JSON.stringify(result, null, 2));
  } else {
    console.log(result);
  }
}

// CLI functionality
async function handleCLI() {
  const args = process.argv.slice(2);

  if (args.length === 0) {
    showHelp();
    return;
  }

  const command = args[0];

  switch (command) {
    case 'help':
    case '--help':
    case '-h':
      showHelp();
      break;

    case 'version':
    case '--version':
    case '-v':
      console.log(`ppt-generator v${pkg.version}`);
      break;

    case 'generate':
      await generatePresentation(args.slice(1));
      break;

    case 'create':
      await createNewPresentation(args.slice(1));
      break;

    case 'demo':
      await createDemoPresentation();
      break;

    case 'open':
      await openPresentation(args.slice(1));
      break;

    case 'replace':
      await replaceTextInPresentation(args.slice(1));
      break;

    case 'add':
      await addContentToPresentation(args.slice(1));
      break;

    case 'util':
      await runUtilCommand(args.slice(1));
      break;

    case 'list':
      await listAvailable(args.slice(1));
      break;

    case 'call':
      await callFunction(args.slice(1));
      break;

    default:
      console.log(`Unknown command: ${command}`);
      console.log('Use "help" to see available commands');
      process.exitCode = 1;
      break;
  }
}

function showHelp() {
  console.log(`
PPT Generator CLI

Usage: node src/index.js [command] [options]

Commands:
  generate <input-file>          Generate presentation from input file
  create <template>              Create new presentation with template
  open <pptx-file>               Open a PPTX file and display slide count
  replace <old> <new> [input]    Replace text in a PPTX and save output
  add <subcommand> [options]     Add content to existing presentation
  util <namespace> <fn> [args]   Call a utility function (color, sizing, positioning, export)
  list [namespace]               List available namespaces / functions / classes
  call <namespace> <fn> [args]   Generically call any exported function by namespace.fn
  help, -h, --help               Show this help message
  version, -v, --version         Show version information

Templates:
  title-slide
  two-column
  comparison
  text
  table
  chart
  shapes
  icons
  image
  all

Utility namespaces (for "util" / "call"):
  color        ColorUtils       e.g. util color hexToRgb "#FF0000"
  sizing       SizingUtils      e.g. util sizing inchesToEmu 1.5
  positioning  PositioningUtils e.g. util positioning center 10 7.5 2 1
  export       ExportUtils      e.g. util export validatePath ./out.pptx

Examples:
  node src/index.js generate ./examples/test-presentation.js
  node src/index.js create title-slide
  node src/index.js create text
  node src/index.js create table
  node src/index.js create chart
  node src/index.js create shapes
  node src/index.js create icons
  node src/index.js create image
  node src/index.js create all
  node src/index.js demo                       # saves to demo-presentation.pptx
  node src/index.js open ./examples/test-presentation.pptx
  node src/index.js replace "Hello" "World" ./examples/test-presentation.pptx
  node src/index.js util color hexToRgb "#2E74B5"
  node src/index.js util sizing inchesToEmu 1.5
  node src/index.js list
  node src/index.js list util
  node src/index.js call util.color hexToRgb "#FFFFFF"
`);
}

async function generatePresentation(args) {
  if (args.length === 0) {
    console.log('Error: Please provide an input file for generation');
    console.log('Usage: node src/index.js generate <input-file>');
    process.exitCode = 2;
    return;
  }

  const inputFile = path.resolve(process.cwd(), args[0]);

  try {
    const inputModule = require(inputFile);

    if (typeof inputModule === 'function') {
      await inputModule();
    } else if (inputModule && typeof inputModule.default === 'function') {
      await inputModule.default();
    } else {
      console.log('Input file does not contain a valid function to execute');
      process.exitCode = 1;
    }
  } catch (error) {
    console.error('Error generating presentation:', error.message);
    process.exitCode = 1;
  }
}

async function createNewPresentation(args, customOutputFileName = null) {
  if (args.length === 0) {
    console.log('Error: Please provide a template name');
    console.log('Usage: node src/index.js create <template>');
    console.log('Available templates: title-slide, two-column, comparison, text, table, chart, shapes, icons, image, all');
    process.exitCode = 2;
    return;
  }

  const template = args[0];
  const outputFileName = customOutputFileName || path.resolve(process.cwd(), `generated-presentation-${Date.now()}.pptx`);

  try {
    const presentation = new Presentation(`Generated Presentation - ${template}`);

    switch (template) {
      case 'title-slide': {
        createTitleSlide(presentation, {
          title: 'Welcome to My Presentation',
          subtitle: 'Generated with PPT Generator CLI'
        });
        break;
      }

      case 'two-column': {
        createTwoColumnSlide(presentation, {
          title: 'Two Column Layout',
          leftContent: 'Left column content',
          rightContent: 'Right column content'
        });
        break;
      }

      case 'comparison': {
        createComparisonSlide(presentation, {
          title: 'Comparison View',
          leftTitle: 'Left Side',
          rightTitle: 'Right Side',
          leftItems: ['Feature A', 'Feature B'],
          rightItems: ['Option 1', 'Option 2']
        });
        break;
      }

      case 'text': {
        const slide = presentation.addSlide();
        addText(slide, 'Text Slide Demo', {
          x: 0.5,
          y: 0.5,
          w: 9,
          h: 1,
          fontFace: 'Arial Black',
          fontSize: 32,
          color: '#2E74B5',
          bold: true,
          align: 'center'
        });
        addText(slide, 'This slide demonstrates text support in the CLI. You can render paragraphs, headings, and formatted content directly from the command line.', {
          x: 0.75,
          y: 1.75,
          w: 8.5,
          h: 4,
          fontFace: 'Arial',
          fontSize: 18,
          color: '#333333',
          align: 'left'
        });
        break;
      }

      case 'table': {
        const slide = presentation.addSlide();
        addText(slide, 'Table Slide Demo', {
          x: 0.5,
          y: 0.5,
          w: 9,
          h: 1,
          fontFace: 'Arial Black',
          fontSize: 32,
          color: '#2E74B5',
          bold: true,
          align: 'center'
        });
        const rows = [
          ['Name', 'Department', 'Status'],
          ['John Smith', 'Engineering', 'Active'],
          ['Jane Doe', 'Marketing', 'Pending'],
          ['Alex Lee', 'Sales', 'Complete']
        ];
        addTable(slide, rows, {
          x: 0.5,
          y: 1.75,
          w: 9,
          h: 3,
          borderColor: '000000',
          borderWidth: 1,
          fontSize: 14,
          color: '000000',
          fillColor: 'F2F2F2',
          boldHeader: true,
          autoFit: true
        });
        break;
      }

      case 'chart': {
        const slide = presentation.addSlide();
        addText(slide, 'Chart Slide Demo', {
          x: 0.5,
          y: 0.5,
          w: 9,
          h: 1,
          fontFace: 'Arial Black',
          fontSize: 32,
          color: '#2E74B5',
          bold: true,
          align: 'center'
        });
        addBarChart(slide, [
          { category: 'Q1', value: 20 },
          { category: 'Q2', value: 35 },
          { category: 'Q3', value: 40 },
          { category: 'Q4', value: 50 }
        ], {
          x: 0.5,
          y: 1.75,
          w: 9,
          h: 4.5,
          title: 'Quarterly Growth',
          showLegend: false,
          showValue: true,
          showCategoryName: true
        });
        break;
      }

      case 'shapes': {
        const slide = presentation.addSlide();
        addText(slide, 'Shape Slide Demo', {
          x: 0.5,
          y: 0.5,
          w: 9,
          h: 1,
          fontFace: 'Arial Black',
          fontSize: 32,
          color: '#2E74B5',
          bold: true,
          align: 'center'
        });
        addRectangle(slide, {
          x: 0.5,
          y: 1.75,
          w: 2.5,
          h: 1.5,
          fillColor: '92D050',
          lineColor: '000000',
          lineWidth: 2
        });
        addCircle(slide, {
          x: 3.5,
          y: 1.75,
          w: 1.5,
          h: 1.5,
          fillColor: 'FF9900',
          lineColor: '000000',
          lineWidth: 2
        });
        addLine(slide, {
          x1: 0.5,
          y1: 4.25,
          x2: 4.5,
          y2: 4.25,
          lineColor: '0000FF',
          lineWidth: 3
        });
        addArrow(slide, {
          x1: 5,
          y1: 4.25,
          x2: 8.5,
          y2: 4.25,
          lineColor: 'FF0000',
          lineWidth: 3
        });
        break;
      }

      case 'icons': {
        const slide = presentation.addSlide();
        addText(slide, 'Icon Slide Demo', {
          x: 0.5,
          y: 0.5,
          w: 9,
          h: 1,
          fontFace: 'Arial Black',
          fontSize: 32,
          color: '#2E74B5',
          bold: true,
          align: 'center'
        });
        addIcon(slide, 'star', { x: 1.5, y: 1.75, color: '#F4B400' });
        addIcon(slide, 'heart', { x: 3.5, y: 1.75, color: '#D93025' });
        addIcon(slide, 'check', { x: 5.5, y: 1.75, color: '#0F9D58' });
        addIcon(slide, 'warning', { x: 7.5, y: 1.75, color: '#FBBC05' });
        addText(slide, 'CLI icon placeholders use Unicode symbols to represent icon assets.', {
          x: 0.5,
          y: 4.25,
          w: 9,
          h: 1,
          fontFace: 'Arial',
          fontSize: 16,
          color: '#333333',
          align: 'center'
        });
        break;
      }

      case 'image': {
        const slide = presentation.addSlide();
        addText(slide, 'Image Slide Demo', {
          x: 0.5,
          y: 0.5,
          w: 9,
          h: 1,
          fontFace: 'Arial Black',
          fontSize: 32,
          color: '#2E74B5',
          bold: true,
          align: 'center'
        });
        addImage(slide, 'assets/sample.png', {
          x: 1,
          y: 1.75,
          w: 6,
          h: 4,
          transparency: 0,
          rotation: 0
        });
        addText(slide, 'Image support uses a local asset path relative to the project root.', {
          x: 0.5,
          y: 6,
          w: 9,
          h: 0.75,
          fontFace: 'Arial',
          fontSize: 14,
          color: '#333333',
          align: 'center'
        });
        break;
      }

      case 'all': {
        createTitleSlide(presentation, {
          title: 'Full Elements Demo',
          subtitle: 'Text, table, charts, shapes, and icons'
        });

        const textSlide = presentation.addSlide();
        addText(textSlide, 'Text Demo', { x: 0.5, y: 0.5, w: 9, h: 1, fontFace: 'Arial Black', fontSize: 30, bold: true, align: 'center' });
        addText(textSlide, 'This demo includes all supported CLI elements in one presentation.', { x: 0.5, y: 1.75, w: 9, h: 3.5, fontFace: 'Arial', fontSize: 18, align: 'left' });

        const tableSlide = presentation.addSlide();
        addText(tableSlide, 'Table Demo', { x: 0.5, y: 0.5, w: 9, h: 1, fontFace: 'Arial Black', fontSize: 30, bold: true, align: 'center' });
        addTable(tableSlide, [['Name', 'Role'], ['Sara', 'Developer'], ['Tom', 'Designer']], { x: 1, y: 1.75, w: 8, h: 2.5, boldHeader: true, autoFit: true });

        const chartSlide = presentation.addSlide();
        addText(chartSlide, 'Chart Demo', { x: 0.5, y: 0.5, w: 9, h: 1, fontFace: 'Arial Black', fontSize: 30, bold: true, align: 'center' });
        addBarChart(chartSlide, [{ category: 'A', value: 30 }, { category: 'B', value: 55 }, { category: 'C', value: 25 }], { x: 0.5, y: 1.75, w: 9, h: 4.5, title: 'Sales', showValue: true, showCategoryName: true });

        const shapeSlide = presentation.addSlide();
        addText(shapeSlide, 'Shape Demo', { x: 0.5, y: 0.5, w: 9, h: 1, fontFace: 'Arial Black', fontSize: 30, bold: true, align: 'center' });
        addRectangle(shapeSlide, { x: 0.5, y: 1.75, w: 2.5, h: 1.5, fillColor: '92D050', lineColor: '000000', lineWidth: 2 });
        addCircle(shapeSlide, { x: 3.5, y: 1.75, w: 1.5, h: 1.5, fillColor: 'FF9900', lineColor: '000000', lineWidth: 2 });
        addArrow(shapeSlide, { x1: 5, y1: 3.25, x2: 8.5, y2: 3.25, lineColor: 'FF0000', lineWidth: 3 });

        const iconSlide = presentation.addSlide();
        addText(iconSlide, 'Icon Demo', { x: 0.5, y: 0.5, w: 9, h: 1, fontFace: 'Arial Black', fontSize: 30, bold: true, align: 'center' });
        addIcon(iconSlide, 'info', { x: 2, y: 1.75, color: '#1E88E5' });
        addIcon(iconSlide, 'check', { x: 4, y: 1.75, color: '#0F9D58' });
        addIcon(iconSlide, 'warning', { x: 6, y: 1.75, color: '#FBBC05' });
        addIcon(iconSlide, 'heart', { x: 8, y: 1.75, color: '#D93025' });
        break;
      }

      default:
        console.log(`Unknown template: ${template}`);
        console.log('Available templates: title-slide, two-column, comparison, text, table, chart, shapes, icons, all');
        process.exitCode = 2;
        return;
    }

    await presentation.save(outputFileName);
    console.log(`Presentation generated successfully: ${outputFileName}`);
  } catch (error) {
    console.error('Error creating presentation:', error.message);
    process.exitCode = 1;
  }
}

async function createDemoPresentation() {
  await createNewPresentation(['all'], path.resolve(process.cwd(), 'demo-presentation.pptx'));
}

async function openPresentation(args) {
  if (args.length === 0) {
    console.log('Error: Please provide a PPTX file path');
    console.log('Usage: node src/index.js open <pptx-file>');
    process.exitCode = 2;
    return;
  }

  const inputFile = path.resolve(process.cwd(), args[0]);
  if (!fs.existsSync(inputFile)) {
    console.error('File not found:', inputFile);
    process.exitCode = 2;
    return;
  }

  try {
    const presentation = await Presentation.open(inputFile);
    const slideCount = presentation._importModel?.slides?.length ?? 0;
    console.log(`Opened ${inputFile}`);
    console.log(`Slide count: ${slideCount}`);
  } catch (error) {
    console.error('Error opening presentation:', error.message);
    process.exitCode = 1;
  }
}

async function replaceTextInPresentation(args) {
  const [oldText, newText, inputArg] = args;
  if (!oldText || !newText) {
    console.log('Error: replace requires <old> <new> [input-file]');
    console.log('Usage: node src/index.js replace <old> <new> [input-file]');
    process.exitCode = 2;
    return;
  }

  const inputFile = path.resolve(process.cwd(), inputArg || 'modified-presentation-input.pptx');
  if (!fs.existsSync(inputFile)) {
    console.error('File not found:', inputFile);
    process.exitCode = 2;
    return;
  }

  const outputFile = path.resolve(process.cwd(), `modified-presentation-${Date.now()}.pptx`);

  try {
    const presentation = await Presentation.open(inputFile);
    presentation.replaceText(oldText, newText);
    await presentation.save(outputFile);
    console.log('Saved', outputFile);
  } catch (error) {
    console.error('Error replacing text:', error.message);
    process.exitCode = 1;
  }
}

async function addContentToPresentation(args) {
  const [subcommand, ...subcommandArgs] = args;

  if (!subcommand) {
    console.log('Error: add command requires a subcommand');
    console.log('Usage: node src/index.js add <subcommand> [options]');
    console.log('Subcommands:');
    console.log('  text <text> <x> <y> [input-file]     Add text to existing presentation');
    console.log('  slide [input-file]                   Create new slide in existing presentation');
    process.exitCode = 2;
    return;
  }

  switch (subcommand) {
    case 'text':
      await addTextToPresentation(subcommandArgs);
      break;
    case 'slide':
      await createSlideInPresentation(subcommandArgs);
      break;
    default:
      console.log(`Unknown add subcommand: ${subcommand}`);
      console.log('Available subcommands: text, slide');
      process.exitCode = 2;
      return;
  }
}

async function addTextToPresentation(args) {
  const [text, x, y, inputArg] = args;

  if (!text || x === undefined || y === undefined) {
    console.log('Error: add:text requires <text> <x> <y> [input-file]');
    console.log('Usage: node src/index.js add:text <text> <x> <y> [input-file]');
    process.exitCode = 2;
    return;
  }

  const inputFile = path.resolve(process.cwd(), inputArg || 'modified-presentation-input.pptx');
  if (!fs.existsSync(inputFile)) {
    console.error('File not found:', inputFile);
    process.exitCode = 2;
    return;
  }

  const outputFile = path.resolve(process.cwd(), `modified-presentation-${Date.now()}.pptx`);

  try {
    const presentation = await Presentation.open(inputFile);

    // Create a new slide
    const slide = presentation.addSlide();

    // Add text to the slide
    addText(slide, text, {
      x: parseFloat(x),
      y: parseFloat(y),
      w: 8,
      h: 1,
      fontFace: 'Arial',
      fontSize: 24,
      color: '#000000'
    });

    await presentation.save(outputFile);
    console.log('Added text to presentation and saved as:', outputFile);
  } catch (error) {
    console.error('Error adding text to presentation:', error.message);
    process.exitCode = 1;
  }
}

async function createSlideInPresentation(args) {
  const inputArg = args[0];

  const inputFile = path.resolve(process.cwd(), inputArg || 'modified-presentation-input.pptx');
  if (!fs.existsSync(inputFile)) {
    console.error('File not found:', inputFile);
    process.exitCode = 2;
    return;
  }

  const outputFile = path.resolve(process.cwd(), `modified-presentation-${Date.now()}.pptx`);

  try {
    const presentation = await Presentation.open(inputFile);

    // Create a new slide
    const slide = presentation.addSlide();

    // Add default text to the new slide
    addText(slide, 'New Slide', {
      x: 0.5,
      y: 0.5,
      w: 9,
      h: 1,
      fontFace: 'Arial Black',
      fontSize: 32,
      color: '#2E74B5',
      bold: true,
      align: 'center'
    });

    await presentation.save(outputFile);
    console.log('Created new slide in presentation and saved as:', outputFile);
  } catch (error) {
    console.error('Error creating slide in presentation:', error.message);
    process.exitCode = 1;
  }
}

// ---------------------------------------------------------------------------
// Utility command: node src/index.js util <namespace> <function> [args...]
//
// Namespaces: color, sizing, positioning, export
// Example: node src/index.js util color hexToRgb "#2E74B5"
// ---------------------------------------------------------------------------
async function runUtilCommand(args) {
  const [namespace, fnName, ...rawFnArgs] = args;

  if (!namespace || !fnName) {
    console.log('Error: util command requires <namespace> <function> [args...]');
    console.log('Usage: node src/index.js util <namespace> <function> [args...]');
    console.log('Available namespaces: ' + Object.keys(UTIL_NAMESPACES).join(', '));
    process.exitCode = 2;
    return;
  }

  const mod = UTIL_NAMESPACES[namespace];
  if (!mod) {
    console.log(`Unknown utility namespace: ${namespace}`);
    console.log('Available namespaces: ' + Object.keys(UTIL_NAMESPACES).join(', '));
    process.exitCode = 2;
    return;
  }

  const fn = mod[fnName];
  if (typeof fn !== 'function') {
    console.log(`Unknown function "${fnName}" on namespace "${namespace}"`);
    const available = Object.keys(mod).filter((k) => typeof mod[k] === 'function');
    console.log('Available functions: ' + (available.length ? available.join(', ') : '(none found)'));
    process.exitCode = 2;
    return;
  }

  try {
    const fnArgs = coerceArgs(rawFnArgs);
    const result = await fn.apply(mod, fnArgs);
    printResult(result);
  } catch (error) {
    console.error(`Error calling ${namespace}.${fnName}:`, error.message);
    process.exitCode = 1;
  }
}

// ---------------------------------------------------------------------------
// Generic introspection: node src/index.js list [namespace]
//
// With no args, lists every top-level export from the API plus utility
// namespaces. With a namespace argument ("util", or a util sub-namespace
// like "color"), lists the functions/members available on it.
// ---------------------------------------------------------------------------
async function listAvailable(args) {
  const [target] = args;

  if (!target) {
    console.log('Top-level exports (api):');
    Object.keys(api)
      .sort()
      .forEach((key) => {
        const type = typeof api[key];
        console.log(`  ${key} (${type})`);
      });
    console.log('');
    console.log('Utility namespaces (use: list util.<namespace>):');
    Object.keys(UTIL_NAMESPACES).forEach((ns) => console.log(`  ${ns}`));
    return;
  }

  if (target === 'util') {
    console.log('Utility namespaces: ' + Object.keys(UTIL_NAMESPACES).join(', '));
    return;
  }

  if (target.startsWith('util.')) {
    const ns = target.split('.')[1];
    const mod = UTIL_NAMESPACES[ns];
    if (!mod) {
      console.log(`Unknown utility namespace: ${ns}`);
      console.log('Available: ' + Object.keys(UTIL_NAMESPACES).join(', '));
      process.exitCode = 2;
      return;
    }
    console.log(`Functions on util.${ns}:`);
    Object.keys(mod)
      .filter((k) => typeof mod[k] === 'function')
      .sort()
      .forEach((k) => console.log(`  ${k}`));
    return;
  }

  // Otherwise treat target as a top-level api member (class or module)
  const member = api[target];
  if (member === undefined) {
    console.log(`Unknown export: ${target}`);
    process.exitCode = 2;
    return;
  }

  if (typeof member === 'function') {
    console.log(`${target} is a function/class.`);
    if (member.prototype) {
      const methods = Object.getOwnPropertyNames(member.prototype).filter((m) => m !== 'constructor');
      if (methods.length) {
        console.log('Instance methods:');
        methods.forEach((m) => console.log(`  ${m}`));
      }
    }
  } else if (typeof member === 'object' && member !== null) {
    console.log(`Members of ${target}:`);
    Object.keys(member)
      .sort()
      .forEach((k) => console.log(`  ${k} (${typeof member[k]})`));
  } else {
    console.log(`${target} = ${member}`);
  }
}

// ---------------------------------------------------------------------------
// Generic function caller: node src/index.js call <namespace.fn> [args...]
//
// Supports:
//   call util.color hexToRgb "#FFFFFF"    -> same as "util color hexToRgb ..."
//   call ColorUtils hexToRgb "#FFFFFF"    -> call a function on a top-level
//                                            api export by name
// ---------------------------------------------------------------------------
async function callFunction(args) {
  const [target, fnName, ...rawFnArgs] = args;

  if (!target || !fnName) {
    console.log('Error: call requires <namespace> <function> [args...]');
    console.log('Usage: node src/index.js call <namespace> <function> [args...]');
    process.exitCode = 2;
    return;
  }

  // util.<ns> shorthand delegates to the util command
  if (target.startsWith('util.')) {
    const ns = target.split('.')[1];
    await runUtilCommand([ns, fnName, ...rawFnArgs]);
    return;
  }

  if (UTIL_NAMESPACES[target]) {
    await runUtilCommand([target, fnName, ...rawFnArgs]);
    return;
  }

  // Otherwise resolve against the top-level api object (e.g. ColorUtils, SizingUtils)
  const mod = api[target];
  if (!mod) {
    console.log(`Unknown namespace/export: ${target}`);
    process.exitCode = 2;
    return;
  }

  const fn = mod[fnName];
  if (typeof fn !== 'function') {
    console.log(`Unknown function "${fnName}" on "${target}"`);
    if (typeof mod === 'object') {
      const available = Object.keys(mod).filter((k) => typeof mod[k] === 'function');
      console.log('Available functions: ' + (available.length ? available.join(', ') : '(none found)'));
    }
    process.exitCode = 2;
    return;
  }

  try {
    const fnArgs = coerceArgs(rawFnArgs);
    const result = await fn.apply(mod, fnArgs);
    printResult(result);
  } catch (error) {
    console.error(`Error calling ${target}.${fnName}:`, error.message);
    process.exitCode = 1;
  }
}

// Export everything as before
const api = {
  // Core classes
  Presentation,
  Slide,
  Theme,
  SlideBuilder,
  TemplateEngine,
  DataBinder,
  AutoLayoutEngine,
  ValidationEngine,

  // Elements
  addText,
  addImage,
  addTable,
  ChartElement,
  addBarChart,
  addLineChart,
  addPieChart,
  addRectangle,
  addCircle,
  addLine,
  addArrow,
  addIcon,
  IconElement,

  // Components
  renderTitle,
  renderBulletList,
  renderMetricCard,
  renderChartCard,

  // Layouts
  TitleSlideLayout,
  createTitleSlide,
  TwoColumnLayout,
  createTwoColumnSlide,
  ComparisonLayout,
  createComparisonSlide,

  // Import modifiers
  PptModifier,

  // Utilities
  ColorUtils,
  SizingUtils,
  PositioningUtils,
  ExportUtils,

  // Constants
  Fonts,
  Themes
};

// New helper function to parse command line arguments
function parseArguments(args) {
  const parsed = {
    flags: {},
    options: {},
    positional: []
  };

  let i = 0;
  while (i < args.length) {
    const arg = args[i];
    
    // Handle flags (--flag, -f)
    if (arg.startsWith('--')) {
      const flag = arg.substring(2);
      parsed.flags[flag] = true;
    } else if (arg.startsWith('-') && arg.length > 1) {
      // Single letter flags
      const flag = arg.substring(1);
      parsed.flags[flag] = true;
    } else if (arg.startsWith('--') && arg.includes('=')) {
      // Handle --option=value format
      const [key, value] = arg.substring(2).split('=', 2);
      parsed.options[key] = value;
    } else if (arg.startsWith('-') && arg.length > 2 && arg.includes('=')) {
      // Handle -o=value format
      const [key, value] = arg.substring(1).split('=', 2);
      parsed.options[key] = value;
    } else {
      // Handle options like --option value
      if (i + 1 < args.length && !args[i + 1].startsWith('-')) {
        const nextArg = args[i + 1];
        if (!nextArg.startsWith('-')) {
          parsed.options[arg.substring(1)] = nextArg;
          i++; // Skip next argument as it's consumed as value
        } else {
          parsed.positional.push(arg);
        }
      } else {
        parsed.positional.push(arg);
      }
    }
    i++;
  }

  return parsed;
}

// Helper function to get command from arguments
function getCommandFromArgs(args) {
  if (args.length === 0) return null;
  return args[0];
}

// Helper function to extract options from args
function extractOptions(args, optionsMap = {}) {
  const options = {};
  const remainingArgs = [];
  
  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    
    // Check if it's a known option
    if (optionsMap[arg] !== undefined) {
      options[arg] = optionsMap[arg];
    } else if (arg.startsWith('--') && arg.includes('=')) {
      // Handle --option=value format
      const [key, value] = arg.substring(2).split('=', 2);
      options[key] = value;
    } else if (arg.startsWith('-') && arg.length > 1 && arg.includes('=')) {
      // Handle -o=value format
      const [key, value] = arg.substring(1).split('=', 2);
      options[key] = value;
    } else if (arg.startsWith('--') || arg.startsWith('-')) {
      // Handle --option value or -o value formats
      if (i + 1 < args.length && !args[i + 1].startsWith('-')) {
        options[arg.replace(/^-+/, '')] = args[i + 1];
        i++; // Skip next argument as it's consumed as value
      } else {
        options[arg.replace(/^-+/, '')] = true;
      }
    } else {
      remainingArgs.push(arg);
    }
  }
  
  return { options, remainingArgs };
}

// New command: inspect
async function inspectPresentation(args) {
  if (args.length === 0) {
    console.log('Error: Please provide a PPTX file path');
    console.log('Usage: node src/index.js inspect <pptx-file>');
    process.exitCode = 2;
    return;
  }

  const inputFile = path.resolve(process.cwd(), args[0]);
  if (!fs.existsSync(inputFile)) {
    console.error('File not found:', inputFile);
    process.exitCode = 2;
    return;
  }

  try {
    const presentation = await Presentation.open(inputFile);
    const slideCount = presentation._importModel?.slides?.length ?? 0;
    
    console.log(`Presentation: ${inputFile}`);
    console.log(`Slide count: ${slideCount}`);
    
    if (presentation._importModel?.slides) {
      console.log('\nSlide titles:');
      presentation._importModel.slides.forEach((slide, index) => {
        const title = slide.title || `Slide ${index + 1}`;
        console.log(`  ${index + 1}: ${title}`);
      });
      
      console.log('\nDetected text elements:');
      presentation._importModel.slides.forEach((slide, slideIndex) => {
        if (slide.shapes) {
          slide.shapes.forEach((shape, shapeIndex) => {
            if (shape.text) {
              console.log(`  Slide ${slideIndex + 1}, Shape ${shapeIndex + 1}: ${shape.text}`);
            }
          });
        }
      });
    }
  } catch (error) {
    console.error('Error inspecting presentation:', error.message);
    process.exitCode = 1;
  }
}

// New command: list
async function listAvailableCommands(args) {
  const [target] = args;
  
  if (!target) {
    console.log('Available commands:');
    console.log('  open             Open a PPTX file');
    console.log('  replace          Replace text in a PPTX file');
    console.log('  generate         Generate presentation from data');
    console.log('  create           Create new presentation with template');
    console.log('  inspect          Inspect presentation details');
    console.log('  list             List available components');
    console.log('  interactive      Interactive mode');
    console.log('  help             Show help');
    console.log('  version          Show version');
    console.log('');
    console.log('Templates:');
    console.log('  title-slide');
    console.log('  two-column');
    console.log('  comparison');
    console.log('  text');
    console.log('  table');
    console.log('  chart');
    console.log('  shapes');
    console.log('  icons');
    console.log('  image');
    console.log('  all');
    console.log('');
    console.log('Utility namespaces:');
    console.log('  color');
    console.log('  sizing');
    console.log('  positioning');
    console.log('  export');
    return;
  }

  if (target === 'templates') {
    console.log('Available templates:');
    console.log('  title-slide');
    console.log('  two-column');
    console.log('  comparison');
    console.log('  text');
    console.log('  table');
    console.log('  chart');
    console.log('  shapes');
    console.log('  icons');
    console.log('  image');
    console.log('  all');
  } else if (target === 'layouts') {
    console.log('Available layouts:');
    console.log('  title-slide');
    console.log('  two-column');
    console.log('  comparison');
  } else if (target === 'utilities') {
    console.log('Available utility namespaces:');
    console.log('  color');
    console.log('  sizing');
    console.log('  positioning');
    console.log('  export');
  } else if (target === 'components') {
    console.log('Available components:');
    console.log('  title');
    console.log('  bullet-list');
    console.log('  metric-card');
    console.log('  chart-card');
  } else {
    console.log(`Unknown list target: ${target}`);
    console.log('Available targets: templates, layouts, utilities, components');
  }
}

// New command: interactive mode
async function interactiveMode() {
  const readline = require('readline');
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
  });

  const questions = [
    'Enter presentation title: ',
    'Enter number of slides: ',
    'Enter template name (title-slide, two-column, etc.): ',
    'Enter output filename: '
  ];

  const answers = [];

  const askQuestion = (index) => {
    if (index >= questions.length) {
      rl.close();
      // Process the interactive inputs
      processInteractiveInputs(answers);
      return;
    }

    rl.question(questions[index], (answer) => {
      answers.push(answer.trim());
      askQuestion(index + 1);
    });
  };

  console.log('Interactive Mode - Enter presentation details:');
  askQuestion(0);
}

// Process interactive inputs
async function processInteractiveInputs(answers) {
  const [title, slides, template, outputFilename] = answers;
  
  if (!title || !slides || !template || !outputFilename) {
    console.log('Error: All fields are required for interactive mode');
    process.exitCode = 2;
    return;
  }

  const numSlides = parseInt(slides);
  if (isNaN(numSlides) || numSlides <= 0) {
    console.log('Error: Number of slides must be a positive integer');
    process.exitCode = 2;
    return;
  }

  try {
    await createNewPresentation([template], outputFilename);
  } catch (error) {
    console.error('Error in interactive mode:', error.message);
    process.exitCode = 1;
  }
}

// Updated generate command to support data-driven input
async function generatePresentation(args) {
  // Parse arguments for data-driven mode
  const parsed = parseArguments(args);
  const options = parsed.options;
  const positional = parsed.positional;
  
  // Check if we're using JSON data mode
  if (options.data) {
    const dataFile = path.resolve(process.cwd(), options.data);
    if (!fs.existsSync(dataFile)) {
      console.error('JSON data file not found:', dataFile);
      process.exitCode = 2;
      return;
    }
    
    try {
      const jsonData = fs.readFileSync(dataFile, 'utf8');
      const data = JSON.parse(jsonData);
      
      // Create presentation from data
      const presentation = new Presentation(data.title || 'Generated Presentation');
      
      // Use DataBinder and TemplateEngine pipeline
      const dataBinder = new DataBinder();
      const templateEngine = new TemplateEngine();
      const slideBuilder = new SlideBuilder();
      
      // Bind data to template
      const boundData = dataBinder.bind(data, {});
      
      // Generate slides using template engine
      const slides = templateEngine.generateSlides(boundData, data.template || 'default');
      
      // Build slides with slide builder
      slides.forEach((slideData) => {
        const slide = presentation.addSlide();
        slideBuilder.buildSlide(slide, slideData);
      });
      
      // Save presentation
      const outputFileName = options.output || `generated-presentation-${Date.now()}.pptx`;
      await presentation.save(outputFileName);
      console.log(`Presentation generated successfully: ${outputFileName}`);
    } catch (error) {
      console.error('Error generating presentation from JSON data:', error.message);
      process.exitCode = 1;
    }
    return;
  }
  
  // Regular mode (backward compatibility)
  if (positional.length === 0) {
    console.log('Error: Please provide an input file for generation');
    console.log('Usage: node src/index.js generate <input-file>');
    process.exitCode = 2;
    return;
  }

  const inputFile = path.resolve(process.cwd(), positional[0]);

  try {
    const inputModule = require(inputFile);

    if (typeof inputModule === 'function') {
      await inputModule();
    } else if (inputModule && typeof inputModule.default === 'function') {
      await inputModule.default();
    } else {
      console.log('Input file does not contain a valid function to execute');
      process.exitCode = 1;
    }
  } catch (error) {
    console.error('Error generating presentation:', error.message);
    process.exitCode = 1;
  }
}

// Updated create command to support data-driven options
async function createNewPresentation(args, customOutputFileName = null) {
  const parsed = parseArguments(args);
  const options = parsed.options;
  const positional = parsed.positional;
  
  if (positional.length === 0) {
    console.log('Error: Please provide a template name');
    console.log('Usage: node src/index.js create <template>');
    console.log('Available templates: title-slide, two-column, comparison, text, table, chart, shapes, icons, image, all');
    process.exitCode = 2;
    return;
  }

  const template = positional[0];
  const outputFileName = customOutputFileName || options.output || path.resolve(process.cwd(), `generated-presentation-${Date.now()}.pptx`);

  try {
    const presentation = new Presentation(options.title || `Generated Presentation - ${template}`);

    switch (template) {
      case 'title-slide': {
        createTitleSlide(presentation, {
          title: options.title || 'Welcome to My Presentation',
          subtitle: options.subtitle || 'Generated with PPT Generator CLI'
        });
        break;
      }

      case 'two-column': {
        createTwoColumnSlide(presentation, {
          title: options.title || 'Two Column Layout',
          leftContent: options.leftContent || 'Left column content',
          rightContent: options.rightContent || 'Right column content'
        });
        break;
      }

      case 'comparison': {
        createComparisonSlide(presentation, {
          title: options.title || 'Comparison View',
          leftTitle: options.leftTitle || 'Left Side',
          rightTitle: options.rightTitle || 'Right Side',
          leftItems: options.leftItems ? options.leftItems.split(',') : ['Feature A', 'Feature B'],
          rightItems: options.rightItems ? options.rightItems.split(',') : ['Option 1', 'Option 2']
        });
        break;
      }

      case 'text': {
        const slide = presentation.addSlide();
        addText(slide, options.text || 'Text Slide Demo', {
          x: 0.5,
          y: 0.5,
          w: 9,
          h: 1,
          fontFace: 'Arial Black',
          fontSize: options.fontSize || 32,
          color: options.color || '#2E74B5',
          bold: options.bold === 'true',
          align: options.align || 'center'
        });
        addText(slide, options.content || 'This slide demonstrates text support in the CLI.', {
          x: 0.75,
          y: 1.75,
          w: 8.5,
          h: 4,
          fontFace: 'Arial',
          fontSize: options.contentFontSize || 18,
          color: options.contentColor || '#333333',
          align: options.contentAlign || 'left'
        });
        break;
      }

      case 'table': {
        const slide = presentation.addSlide();
        addText(slide, options.title || 'Table Slide Demo', {
          x: 0.5,
          y: 0.5,
          w: 9,
          h: 1,
          fontFace: 'Arial Black',
          fontSize: options.titleFontSize || 32,
          color: options.titleColor || '#2E74B5',
          bold: options.titleBold === 'true',
          align: options.titleAlign || 'center'
        });
        
        const rows = options.rows ? JSON.parse(options.rows) : [
          ['Name', 'Department', 'Status'],
          ['John Smith', 'Engineering', 'Active'],
          ['Jane Doe', 'Marketing', 'Pending'],
          ['Alex Lee', 'Sales', 'Complete']
        ];
        
        addTable(slide, rows, {
          x: 0.5,
          y: 1.75,
          w: 9,
          h: 3,
          borderColor: options.borderColor || '000000',
          borderWidth: options.borderWidth || 1,
          fontSize: options.fontSize || 14,
          color: options.color || '000000',
          fillColor: options.fillColor || 'F2F2F2',
          boldHeader: options.boldHeader === 'true',
          autoFit: options.autoFit === 'true'
        });
        break;
      }

      case 'chart': {
        const slide = presentation.addSlide();
        addText(slide, options.title || 'Chart Slide Demo', {
          x: 0.5,
          y: 0.5,
          w: 9,
          h: 1,
          fontFace: 'Arial Black',
          fontSize: options.titleFontSize || 32,
          color: options.titleColor || '#2E74B5',
          bold: options.titleBold === 'true',
          align: options.titleAlign || 'center'
        });
        
        const chartData = options.data ? JSON.parse(options.data) : [
          { category: 'Q1', value: 20 },
          { category: 'Q2', value: 35 },
          { category: 'Q3', value: 40 },
          { category: 'Q4', value: 50 }
        ];
        
        addBarChart(slide, chartData, {
          x: 0.5,
          y: 1.75,
          w: 9,
          h: 4.5,
          title: options.chartTitle || 'Quarterly Growth',
          showLegend: options.showLegend === 'true',
          showValue: options.showValue === 'true',
          showCategoryName: options.showCategoryName === 'true'
        });
        break;
      }

      case 'shapes': {
        const slide = presentation.addSlide();
        addText(slide, options.title || 'Shape Slide Demo', {
          x: 0.5,
          y: 0.5,
          w: 9,
          h: 1,
          fontFace: 'Arial Black',
          fontSize: options.titleFontSize || 32,
          color: options.titleColor || '#2E74B5',
          bold: options.titleBold === 'true',
          align: options.titleAlign || 'center'
        });
        
        // Add rectangles
        if (options.rectangles) {
          const rectangles = JSON.parse(options.rectangles);
          rectangles.forEach(rect => {
            addRectangle(slide, rect);
          });
        }
        
        // Add circles
        if (options.circles) {
          const circles = JSON.parse(options.circles);
          circles.forEach(circle => {
            addCircle(slide, circle);
          });
        }
        
        // Add lines
        if (options.lines) {
          const lines = JSON.parse(options.lines);
          lines.forEach(line => {
            addLine(slide, line);
          });
        }
        
        // Add arrows
        if (options.arrows) {
          const arrows = JSON.parse(options.arrows);
          arrows.forEach(arrow => {
            addArrow(slide, arrow);
          });
        }
        break;
      }

      case 'icons': {
        const slide = presentation.addSlide();
        addText(slide, options.title || 'Icon Slide Demo', {
          x: 0.5,
          y: 0.5,
          w: 9,
          h: 1,
          fontFace: 'Arial Black',
          fontSize: options.titleFontSize || 32,
          color: options.titleColor || '#2E74B5',
          bold: options.titleBold === 'true',
          align: options.titleAlign || 'center'
        });
        
        if (options.icons) {
          const icons = JSON.parse(options.icons);
          icons.forEach(icon => {
            addIcon(slide, icon.type, icon.options);
          });
        }
        break;
      }

      case 'image': {
        const slide = presentation.addSlide();
        addText(slide, options.title || 'Image Slide Demo', {
          x: 0.5,
          y: 0.5,
          w: 9,
          h: 1,
          fontFace: 'Arial Black',
          fontSize: options.titleFontSize || 32,
          color: options.titleColor || '#2E74B5',
          bold: options.titleBold === 'true',
          align: options.titleAlign || 'center'
        });
        
        const imagePath = options.image || 'assets/sample.png';
        addImage(slide, imagePath, {
          x: options.x || 1,
          y: options.y || 1.75,
          w: options.w || 6,
          h: options.h || 4,
          transparency: options.transparency || 0,
          rotation: options.rotation || 0
        });
        break;
      }

      case 'all': {
        createTitleSlide(presentation, {
          title: options.title || 'Full Elements Demo',
          subtitle: options.subtitle || 'Text, table, charts, shapes, and icons'
        });

        const textSlide = presentation.addSlide();
        addText(textSlide, options.textTitle || 'Text Demo', { x: 0.5, y: 0.5, w: 9, h: 1, fontFace: 'Arial Black', fontSize: 30, bold: true, align: 'center' });
        addText(textSlide, options.textContent || 'This demo includes all supported CLI elements in one presentation.', { x: 0.5, y: 1.75, w: 9, h: 3.5, fontFace: 'Arial', fontSize: 18, align: 'left' });

        const tableSlide = presentation.addSlide();
        addText(tableSlide, options.tableTitle || 'Table Demo', { x: 0.5, y: 0.5, w: 9, h: 1, fontFace: 'Arial Black', fontSize: 30, bold: true, align: 'center' });
        addTable(tableSlide, [['Name', 'Role'], ['Sara', 'Developer'], ['Tom', 'Designer']], { x: 1, y: 1.75, w: 8, h: 2.5, boldHeader: true, autoFit: true });

        const chartSlide = presentation.addSlide();
        addText(chartSlide, options.chartTitle || 'Chart Demo', { x: 0.5, y: 0.5, w: 9, h: 1, fontFace: 'Arial Black', fontSize: 30, bold: true, align: 'center' });
        addBarChart(chartSlide, [{ category: 'A', value: 30 }, { category: 'B', value: 55 }, { category: 'C', value: 25 }], { x: 0.5, y: 1.75, w: 9, h: 4.5, title: 'Sales', showValue: true, showCategoryName: true });

        const shapeSlide = presentation.addSlide();
        addText(shapeSlide, options.shapeTitle || 'Shape Demo', { x: 0.5, y: 0.5, w: 9, h: 1, fontFace: 'Arial Black', fontSize: 30, bold: true, align: 'center' });
        addRectangle(shapeSlide, { x: 0.5, y: 1.75, w: 2.5, h: 1.5, fillColor: '92D050', lineColor: '000000', lineWidth: 2 });
        addCircle(shapeSlide, { x: 3.5, y: 1.75, w: 1.5, h: 1.5, fillColor: 'FF9900', lineColor: '000000', lineWidth: 2 });
        addArrow(shapeSlide, { x1: 5, y1: 3.25, x2: 8.5, y2: 3.25, lineColor: 'FF0000', lineWidth: 3 });

        const iconSlide = presentation.addSlide();
        addText(iconSlide, options.iconTitle || 'Icon Demo', { x: 0.5, y: 0.5, w: 9, h: 1, fontFace: 'Arial Black', fontSize: 30, bold: true, align: 'center' });
        addIcon(iconSlide, 'info', { x: 2, y: 1.75, color: '#1E88E5' });
        addIcon(iconSlide, 'check', { x: 4, y: 1.75, color: '#0F9D58' });
        addIcon(iconSlide, 'warning', { x: 6, y: 1.75, color: '#FBBC05' });
        addIcon(iconSlide, 'heart', { x: 8, y: 1.75, color: '#D93025' });
        break;
      }

      default:
        console.log(`Unknown template: ${template}`);
        console.log('Available templates: title-slide, two-column, comparison, text, table, chart, shapes, icons, all');
        process.exitCode = 2;
        return;
    }

    await presentation.save(outputFileName);
    console.log(`Presentation generated successfully: ${outputFileName}`);
  } catch (error) {
    console.error('Error creating presentation:', error.message);
    process.exitCode = 1;
  }
}

// Updated open command to handle data-driven arguments
async function openPresentation(args) {
  if (args.length === 0) {
    console.log('Error: Please provide a PPTX file path');
    console.log('Usage: node src/index.js open <pptx-file>');
    process.exitCode = 2;
    return;
  }

  const inputFile = path.resolve(process.cwd(), args[0]);
  if (!fs.existsSync(inputFile)) {
    console.error('File not found:', inputFile);
    process.exitCode = 2;
    return;
  }

  try {
    const presentation = await Presentation.open(inputFile);
    const slideCount = presentation._importModel?.slides?.length ?? 0;
    console.log(`Opened ${inputFile}`);
    console.log(`Slide count: ${slideCount}`);
  } catch (error) {
    console.error('Error opening presentation:', error.message);
    process.exitCode = 1;
  }
}

// Updated replace command to handle data-driven arguments
async function replaceTextInPresentation(args) {
  const parsed = parseArguments(args);
  const options = parsed.options;
  const positional = parsed.positional;
  
  const [oldText, newText, inputFile] = positional;
  if (!oldText || !newText) {
    console.log('Error: replace requires <old> <new> [input-file]');
    console.log('Usage: node src/index.js replace <old> <new> [input-file]');
    process.exitCode = 2;
    return;
  }

  const inputFilePath = path.resolve(process.cwd(), inputFile || 'modified-presentation-input.pptx');
  if (!fs.existsSync(inputFilePath)) {
    console.error('File not found:', inputFilePath);
    process.exitCode = 2;
    return;
  }

  const outputFile = options.output || path.resolve(process.cwd(), `modified-presentation-${Date.now()}.pptx`);

  try {
    const presentation = await Presentation.open(inputFilePath);
    presentation.replaceText(oldText, newText);
    await presentation.save(outputFile);
    console.log('Saved', outputFile);
  } catch (error) {
    console.error('Error replacing text:', error.message);
    process.exitCode = 1;
  }
}

// Updated add command to handle data-driven arguments
async function addContentToPresentation(args) {
  const parsed = parseArguments(args);
  const options = parsed.options;
  const positional = parsed.positional;
  
  const [subcommand, ...subcommandArgs] = positional;

  if (!subcommand) {
    console.log('Error: add command requires a subcommand');
    console.log('Usage: node src/index.js add <subcommand> [options]');
    console.log('Subcommands:');
    console.log('  text <text> <x> <y> [input-file]     Add text to existing presentation');
    console.log('  slide [input-file]                   Create new slide in existing presentation');
    process.exitCode = 2;
    return;
  }

  switch (subcommand) {
    case 'text':
      await addTextToPresentation(subcommandArgs);
      break;
    case 'slide':
      await createSlideInPresentation(subcommandArgs);
      break;
    default:
      console.log(`Unknown add subcommand: ${subcommand}`);
      console.log('Available subcommands: text, slide');
      process.exitCode = 2;
      return;
  }
}

// Updated addTextToPresentation to handle data-driven arguments
async function addTextToPresentation(args) {
  const parsed = parseArguments(args);
  const options = parsed.options;
  const positional = parsed.positional;
  
  const [text, x, y, inputFile] = positional;

  if (!text || x === undefined || y === undefined) {
    console.log('Error: add:text requires <text> <x> <y> [input-file]');
    console.log('Usage: node src/index.js add:text <text> <x> <y> [input-file]');
    process.exitCode = 2;
    return;
  }

  const inputFilePath = path.resolve(process.cwd(), inputFile || 'modified-presentation-input.pptx');
  if (!fs.existsSync(inputFilePath)) {
    console.error('File not found:', inputFilePath);
    process.exitCode = 2;
    return;
  }

  const outputFile = options.output || path.resolve(process.cwd(), `modified-presentation-${Date.now()}.pptx`);

  try {
    const presentation = await Presentation.open(inputFilePath);

    // Create a new slide
    const slide = presentation.addSlide();

    // Add text to the slide
    addText(slide, text, {
      x: parseFloat(x),
      y: parseFloat(y),
      w: options.width || 8,
      h: options.height || 1,
      fontFace: options.fontFace || 'Arial',
      fontSize: options.fontSize || 24,
      color: options.color || '#000000'
    });

    await presentation.save(outputFile);
    console.log('Added text to presentation and saved as:', outputFile);
  } catch (error) {
    console.error('Error adding text to presentation:', error.message);
    process.exitCode = 1;
  }
}

// Updated createSlideInPresentation to handle data-driven arguments
async function createSlideInPresentation(args) {
  const parsed = parseArguments(args);
  const options = parsed.options;
  const positional = parsed.positional;
  
  const inputFile = positional[0];

  const inputFilePath = path.resolve(process.cwd(), inputFile || 'modified-presentation-input.pptx');
  if (!fs.existsSync(inputFilePath)) {
    console.error('File not found:', inputFilePath);
    process.exitCode = 2;
    return;
  }

  const outputFile = options.output || path.resolve(process.cwd(), `modified-presentation-${Date.now()}.pptx`);

  try {
    const presentation = await Presentation.open(inputFilePath);

    // Create a new slide
    const slide = presentation.addSlide();

    // Add default text to the new slide
    addText(slide, options.text || 'New Slide', {
      x: options.x || 0.5,
      y: options.y || 0.5,
      w: options.w || 9,
      h: options.h || 1,
      fontFace: options.fontFace || 'Arial Black',
      fontSize: options.fontSize || 32,
      color: options.color || '#2E74B5',
      bold: options.bold === 'true',
      align: options.align || 'center'
    });

    await presentation.save(outputFile);
    console.log('Created new slide in presentation and saved as:', outputFile);
  } catch (error) {
    console.error('Error creating slide in presentation:', error.message);
    process.exitCode = 1;
  }
}

// Updated CLI handler to include new commands
async function handleCLI() {
  const args = process.argv.slice(2);

  if (args.length === 0) {
    showHelp();
    return;
  }

  const command = args[0];

  switch (command) {
    case 'help':
    case '--help':
    case '-h':
      showHelp();
      break;

    case 'version':
    case '--version':
    case '-v':
      console.log(`ppt-generator v${pkg.version}`);
      break;

    case 'generate':
      await generatePresentation(args.slice(1));
      break;

    case 'create':
      await createNewPresentation(args.slice(1));
      break;

    case 'demo':
      await createDemoPresentation();
      break;

    case 'open':
      await openPresentation(args.slice(1));
      break;

    case 'replace':
      await replaceTextInPresentation(args.slice(1));
      break;

    case 'add':
      await addContentToPresentation(args.slice(1));
      break;

    case 'util':
      await runUtilCommand(args.slice(1));
      break;

    case 'list':
      await listAvailableCommands(args.slice(1));
      break;

    case 'inspect':
      await inspectPresentation(args.slice(1));
      break;

    case 'interactive':
      await interactiveMode();
      break;

    case 'call':
      await callFunction(args.slice(1));
      break;

    default:
      console.log(`Unknown command: ${command}`);
      console.log('Use "help" to see available commands');
      process.exitCode = 1;
      break;
  }
}

// Updated help to include new commands
function showHelp() {
  console.log(`
PPT Generator CLI

Usage: node src/index.js [command] [options]

Commands:
  generate [options]             Generate presentation from input or data
  create [options]               Create new presentation with template
  open <pptx-file>               Open a PPTX file and display slide count
  replace [options]              Replace text in a PPTX and save output
  add <subcommand> [options]     Add content to existing presentation
  util <namespace> <fn> [args]   Call a utility function (color, sizing, positioning, export)
  list [target]                  List available namespaces / functions / classes
  inspect <pptx-file>            Inspect presentation details
  interactive                    Interactive mode for creating presentations
  call <namespace> <fn> [args]   Generically call any exported function by namespace.fn
  help, -h, --help               Show this help message
  version, -v, --version         Show version information

Options for generate:
  --data <file.json>             Generate from JSON data file
  --output <filename.pptx>       Specify output filename

Options for create:
  --title <title>                Presentation title
  --subtitle <subtitle>          Presentation subtitle
  --output <filename.pptx>       Specify output filename
  --left-content <content>       Left column content
  --right-content <content>      Right column content
  --text <text>                  Text content
  --content <content>            Detailed content
  --fontSize <number>            Font size for text
  --color <hex>                  Text color
  --align <left|center|right>    Text alignment
  --rows <json>                  Table rows as JSON
  --chart-title <title>          Chart title
  --show-value <true|false>      Show values on chart
  --show-category-name <true|false> Show category names on chart

Options for replace:
  --output <filename.pptx>       Specify output filename

Options for add:text:
  --output <filename.pptx>       Specify output filename
  --width <number>               Width of text box
  --height <number>              Height of text box
  --font-face <string>           Font face
  --font-size <number>           Font size
  --color <hex>                  Text color

Options for add:slide:
  --output <filename.pptx>       Specify output filename
  --text <text>                  Slide text
  --x <number>                   X position
  --y <number>                   Y position
  --w <number>                   Width
  --h <number>                   Height
  --font-face <string>           Font face
  --font-size <number>           Font size
  --color <hex>                  Text color
  --bold <true|false>            Bold text
  --align <left|center|right>    Text alignment

Templates:
  title-slide
  two-column
  comparison
  text
  table
  chart
  shapes
  icons
  image
  all

Utility namespaces (for "util" / "call"):
  color        ColorUtils       e.g. util color hexToRgb "#FF0000"
  sizing       SizingUtils      e.g. util sizing inchesToEmu 1.5
  positioning  PositioningUtils e.g. util positioning center 10 7.5 2 1
  export       ExportUtils      e.g. util export validatePath ./out.pptx

Examples:
  node src/index.js generate ./examples/test-presentation.js
  node src/index.js create title-slide --title "My Presentation"
  node src/index.js create text --title "Text Slide" --text "Hello World"
  node src/index.js create table --title "Data Table" --rows "[[\"Name\", \"Age\"], [\"John\", 30]]"
  node src/index.js demo                       # saves to demo-presentation.pptx
  node src/index.js open ./examples/test-presentation.pptx
  node src/index.js replace "Hello" "World" ./examples/test-presentation.pptx
  node src/index.js util color hexToRgb "#2E74B5"
  node src/index.js util sizing inchesToEmu 1.5
  node src/index.js list
  node src/index.js list util
  node src/index.js call util.color hexToRgb "#FFFFFF"
  node src/index.js generate --data data.json
  node src/index.js inspect ./examples/test-presentation.pptx
  node src/index.js interactive
`);
}

// If running as CLI, handle commands
if (require.main === module) {
  handleCLI().catch((err) => {
    console.error('CLI error:', err.message || err);
    process.exit(1);
  });
}

module.exports = api;
