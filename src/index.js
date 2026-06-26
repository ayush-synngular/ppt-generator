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

// PptModifier — direct zip-level editing of existing PPTX files
const PptModifier = require('./import/PptModifier');
const PptReader   = require('./import/PptReader');

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
  console.log('');
  console.log('Slide Editing Commands (operate on existing PPTX files):');
  console.log('  delete-slide --input <file.pptx> --slide <n> [--output <file.pptx>]');
  console.log('  duplicate-slide --input <file.pptx> --slide <n> [--output <file.pptx>]');
  console.log('  move-slide --input <file.pptx> --from <n> --to <n> [--output <file.pptx>]');
  console.log('  slide-count --input <file.pptx>');
  console.log('  replace-global --input <file.pptx> --old "text" --new "text" [--flags gi] [--output <file.pptx>]');
  console.log('  set-background --input <file.pptx> --slide <n> --color <hex> [--output <file.pptx>]');
  console.log('  add-image --input <file.pptx> --slide <n> --path <img> [--x n] [--y n] [--w n] [--h n] [--output <file.pptx>]');
  console.log('  add-slide --input <file.pptx> --text "content" [--x n] [--y n] [--w n] [--h n] [--fontSize n] [--color hex] [--bold true] [--align left|center|right] [--fontFace name] [--output <file.pptx>]');
  console.log('');
  console.log('File System Commands:');
  console.log('  read-file --path <file> [--startLine <n>] [--endLine <n>]');
  console.log('  write-file --path <file> --content "text"');
  console.log('  edit-file --path <file> --replace "text" [--search "old"] [--startLine <n>] [--endLine <n>]');
  console.log('  grep --pattern <regex> --path <dir> [--include <glob>]');
  console.log('  glob --pattern <glob> --path <dir>');
  console.log('  list-dir --path <dir> [--recursive true]');
  console.log('  create-dir --path <dir>');
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
  console.log('  node src/index.js delete-slide --input ai-edu.pptx --slide 5 --output ai-edu-updated.pptx');
  console.log('  node src/index.js duplicate-slide --input ai-edu.pptx --slide 2 --output ai-edu-updated.pptx');
  console.log('  node src/index.js move-slide --input ai-edu.pptx --from 3 --to 1 --output ai-edu-updated.pptx');
  console.log('  node src/index.js slide-count --input ai-edu.pptx');
  console.log('  node src/index.js replace-global --input ai-edu.pptx --old "AI" --new "Artificial Intelligence" --flags gi --output ai-edu-updated.pptx');
  console.log('  node src/index.js set-background --input ai-edu.pptx --slide 1 --color FF0000 --output ai-edu-updated.pptx');
  console.log('  node src/index.js add-image --input ai-edu.pptx --slide 1 --path logo.png --x 0.5 --y 0.5 --w 3 --h 2 --output ai-edu-updated.pptx');
  console.log('  node src/index.js add-slide --input ai-edu.pptx --text "New Slide Content" --output ai-edu-updated.pptx');
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

// -------------------------------------------------------------------
// Shared helper: open a PPTX via PptReader, run a modifier callback,
// then save to outputName. Keeps all new handlers DRY.
// -------------------------------------------------------------------
async function withModifier(inputFile, outputName, fn) {
  if (!inputFile) {
    console.error('❌ Missing --input <file.pptx>');
    process.exit(1);
  }
  if (!require('fs').existsSync(inputFile)) {
    console.error(`❌ Input file not found: ${inputFile}`);
    process.exit(1);
  }
  try {
    const presentation = await new PptReader().read(inputFile);
    const modifier = new PptModifier(presentation);
    await fn(modifier);
    await modifier.save(outputName);
    console.log(`🎉 Done. Saved to: ${outputName}`);
  } catch (err) {
    console.error(`❌ Error: ${err.message}`);
  }
}

function runDeleteSlide(flags, outputName) {
  const slideNumber = parseInt(flags.slide);
  if (!slideNumber || slideNumber < 1) {
    console.error('❌ Missing or invalid --slide <n> (must be integer >= 1).');
    process.exit(1);
  }
  withModifier(flags.input, outputName, (modifier) => {
    modifier.deleteSlide(slideNumber);
  });
}

function runDuplicateSlide(flags, outputName) {
  const slideNumber = parseInt(flags.slide);
  if (!slideNumber || slideNumber < 1) {
    console.error('❌ Missing or invalid --slide <n> (must be integer >= 1).');
    process.exit(1);
  }
  withModifier(flags.input, outputName, (modifier) => {
    modifier.duplicateSlide(slideNumber);
  });
}

function runMoveSlide(flags, outputName) {
  const from = parseInt(flags.from);
  const to   = parseInt(flags.to);
  if (!from || from < 1 || !to || to < 1) {
    console.error('❌ Missing or invalid --from <n> and/or --to <n> (must be integers >= 1).');
    process.exit(1);
  }
  withModifier(flags.input, outputName, (modifier) => {
    modifier.moveSlide(from, to);
  });
}

async function runSlideCount(flags) {
  const inputFile = flags.input;
  if (!inputFile) {
    console.error('❌ Missing --input <file.pptx>');
    process.exit(1);
  }
  if (!require('fs').existsSync(inputFile)) {
    console.error(`❌ Input file not found: ${inputFile}`);
    process.exit(1);
  }
  try {
    const presentation = await new PptReader().read(inputFile);
    const modifier = new PptModifier(presentation);
    console.log(`📊 Slide count: ${modifier.getSlideCount()}`);
  } catch (err) {
    console.error(`❌ Error: ${err.message}`);
  }
}

function runReplaceGlobal(flags, outputName) {
  const oldText = flags.old;
  const newText = flags.new;
  if (!oldText || newText === undefined) {
    console.error('❌ Missing --old "text" and/or --new "text".');
    process.exit(1);
  }
  const regexFlags = flags.flags || 'g';
  withModifier(flags.input, outputName, (modifier) => {
    modifier.replaceTextGlobal(oldText, newText, regexFlags);
  });
}

function runSetBackground(flags, outputName) {
  const slideNumber = parseInt(flags.slide);
  const color = flags.color;
  if (!slideNumber || slideNumber < 1) {
    console.error('❌ Missing or invalid --slide <n> (must be integer >= 1).');
    process.exit(1);
  }
  if (!color) {
    console.error('❌ Missing --color <hex> (e.g. FF0000 or #FF0000).');
    process.exit(1);
  }
  withModifier(flags.input, outputName, (modifier) => {
    modifier.setSlideBackground(slideNumber, color);
  });
}

function runAddImage(flags, outputName) {
  const slideNumber = parseInt(flags.slide);
  const imagePath   = flags.path;
  if (!slideNumber || slideNumber < 1) {
    console.error('❌ Missing or invalid --slide <n> (must be integer >= 1).');
    process.exit(1);
  }
  if (!imagePath) {
    console.error('❌ Missing --path <image-file>.');
    process.exit(1);
  }
  const opts = {
    x: parseFloat(flags.x) || 0.5,
    y: parseFloat(flags.y) || 0.5,
    w: parseFloat(flags.w) || 3,
    h: parseFloat(flags.h) || 2,
  };
  withModifier(flags.input, outputName, (modifier) => {
    modifier.addImageToSlide(slideNumber, imagePath, opts);
  });
}

function runAddSlide(flags, outputName) {
  const text = flags.text !== undefined ? flags.text : '';
  const opts = {
    x:        parseFloat(flags.x)        || 0.5,
    y:        parseFloat(flags.y)        || 0.5,
    w:        parseFloat(flags.w)        || 9,
    h:        parseFloat(flags.h)        || 1,
    fontSize: parseFloat(flags.fontSize) || 24,
    color:    flags.color                || '000000',
    bold:     flags.bold === 'true',
    align:    flags.align                || 'left',
    fontFace: flags.fontFace             || 'Arial',
  };
  withModifier(flags.input, outputName, (modifier) => {
    modifier.addSlide(text, opts);
  });
}

// -------------------------------------------------------------------
// File-system utility handlers (from tool.json: read-file, write-file,
// edit-file, grep, glob, list-dir, create-dir)
// These operate on plain files — not on PPTX internals.
// -------------------------------------------------------------------

function runReadFile(flags) {
  const filePath = flags.path;
  if (!filePath) {
    console.error('❌ Missing --path <file>');
    process.exit(1);
  }
  if (!fs.existsSync(filePath)) {
    console.error(`❌ File not found: ${filePath}`);
    process.exit(1);
  }
  try {
    const lines = fs.readFileSync(filePath, 'utf8').split('\n');
    const startLine = flags.startLine ? parseInt(flags.startLine) - 1 : 0;
    const endLine   = flags.endLine   ? parseInt(flags.endLine)       : lines.length;
    const selected  = lines.slice(startLine, endLine);
    selected.forEach((line, i) => console.log(`${startLine + i + 1}\t${line}`));
  } catch (err) {
    console.error(`❌ Error reading file: ${err.message}`);
  }
}

function runWriteFile(flags) {
  const filePath = flags.path;
  const content  = flags.content;
  if (!filePath) {
    console.error('❌ Missing --path <file>');
    process.exit(1);
  }
  if (content === undefined) {
    console.error('❌ Missing --content "text"');
    process.exit(1);
  }
  try {
    const dir = require('path').dirname(filePath);
    if (dir && dir !== '.') fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(filePath, content, 'utf8');
    console.log(`🎉 Written to: ${filePath}`);
  } catch (err) {
    console.error(`❌ Error writing file: ${err.message}`);
  }
}

function runEditFile(flags) {
  const filePath = flags.path;
  const replace  = flags.replace;
  if (!filePath) {
    console.error('❌ Missing --path <file>');
    process.exit(1);
  }
  if (replace === undefined) {
    console.error('❌ Missing --replace "text"');
    process.exit(1);
  }
  if (!fs.existsSync(filePath)) {
    console.error(`❌ File not found: ${filePath}`);
    process.exit(1);
  }
  try {
    let content = fs.readFileSync(filePath, 'utf8');

    if (flags.search !== undefined) {
      // Search-replace mode
      if (!content.includes(flags.search)) {
        console.error(`❌ Search string not found in file: ${flags.search}`);
        process.exit(1);
      }
      content = content.split(flags.search).join(replace);
    } else if (flags.startLine !== undefined) {
      // Line-range replacement mode
      const lines     = content.split('\n');
      const startLine = parseInt(flags.startLine) - 1;
      const endLine   = flags.endLine ? parseInt(flags.endLine) : startLine + 1;
      lines.splice(startLine, endLine - startLine, replace);
      content = lines.join('\n');
    } else {
      console.error('❌ Provide either --search "text" or --startLine <n> for edit-file.');
      process.exit(1);
    }

    fs.writeFileSync(filePath, content, 'utf8');
    console.log(`🎉 File edited: ${filePath}`);
  } catch (err) {
    console.error(`❌ Error editing file: ${err.message}`);
  }
}

function runGrep(flags) {
  const pattern   = flags.pattern;
  const dirPath   = flags.path;
  const include   = flags.include || '*';
  if (!pattern || !dirPath) {
    console.error('❌ Missing --pattern <regex> and/or --path <dir>');
    process.exit(1);
  }
  if (!fs.existsSync(dirPath)) {
    console.error(`❌ Directory not found: ${dirPath}`);
    process.exit(1);
  }
  try {
    const { execSync } = require('child_process');
    const cmd = process.platform === 'win32'
      ? `findstr /S /N /R "${pattern}" "${dirPath}\\${include}"`
      : `grep -rn --include="${include}" "${pattern}" "${dirPath}"`;
    const result = execSync(cmd, { encoding: 'utf8', stdio: ['pipe', 'pipe', 'ignore'] });
    process.stdout.write(result);
  } catch (err) {
    // grep exits with code 1 when no matches — that is not an error
    if (err.stdout) process.stdout.write(err.stdout);
    else console.log('(no matches found)');
  }
}

function runGlob(flags) {
  const pattern = flags.pattern;
  const dirPath = flags.path;
  if (!pattern || !dirPath) {
    console.error('❌ Missing --pattern <glob> and/or --path <dir>');
    process.exit(1);
  }
  if (!fs.existsSync(dirPath)) {
    console.error(`❌ Directory not found: ${dirPath}`);
    process.exit(1);
  }
  try {
    const { execSync } = require('child_process');
    const cmd = process.platform === 'win32'
      ? `dir /S /B "${require('path').join(dirPath, pattern)}"`
      : `find "${dirPath}" -path "${require('path').join(dirPath, pattern)}" -type f`;
    const result = execSync(cmd, { encoding: 'utf8', stdio: ['pipe', 'pipe', 'ignore'] });
    process.stdout.write(result);
  } catch (err) {
    if (err.stdout) process.stdout.write(err.stdout);
    else console.log('(no files matched)');
  }
}

function runListDir(flags) {
  const dirPath   = flags.path;
  const recursive = flags.recursive === 'true' || flags.recursive === true;
  if (!dirPath) {
    console.error('❌ Missing --path <dir>');
    process.exit(1);
  }
  if (!fs.existsSync(dirPath)) {
    console.error(`❌ Directory not found: ${dirPath}`);
    process.exit(1);
  }
  try {
    function listDir(dir, indent) {
      const entries = fs.readdirSync(dir, { withFileTypes: true });
      entries.forEach(entry => {
        const fullPath = require('path').join(dir, entry.name);
        if (entry.isDirectory()) {
          console.log(`${indent}📁 ${entry.name}/`);
          if (recursive) listDir(fullPath, indent + '  ');
        } else {
          const size = fs.statSync(fullPath).size;
          console.log(`${indent}📄 ${entry.name} (${size} bytes)`);
        }
      });
    }
    console.log(`📂 ${dirPath}`);
    listDir(dirPath, '  ');
  } catch (err) {
    console.error(`❌ Error listing directory: ${err.message}`);
  }
}

function runCreateDir(flags) {
  const dirPath = flags.path;
  if (!dirPath) {
    console.error('❌ Missing --path <dir>');
    process.exit(1);
  }
  try {
    fs.mkdirSync(dirPath, { recursive: true });
    console.log(`🎉 Directory created: ${dirPath}`);
  } catch (err) {
    console.error(`❌ Error creating directory: ${err.message}`);
  }
}

/**
 * Main execution function
 * This function handles CLI commands and can be called programmatically
 */
function main() {
  // Check if we're running in machine interface mode
  const workspacePath = process.env.WORKSPACE_PATH;
  
  // If WORKSPACE_PATH is set, we're in machine interface mode
  if (workspacePath) {
    // Machine interface mode
    if (process.argv.length < 3) {
      console.error(JSON.stringify({
        success: false,
        error: {
          code: "INVALID_ARGS",
          message: "Operation name and JSON parameters required",
          recovery: "Usage: WORKSPACE_PATH=/path/to/workspace node dist/index.js <operation> '{\"param\":\"value\"}'"
        }
      }));
      process.exit(1);
    }
    
    const operation = process.argv[2];
    const paramsJson = process.argv[3];
    
    // Validate workspace path
    const fs = require('fs');
    if (!fs.existsSync(workspacePath)) {
      console.error(JSON.stringify({
        success: false,
        error: {
          code: "WORKSPACE_NOT_FOUND",
          message: `Workspace path does not exist: ${workspacePath}`,
          recovery: "Ensure the workspace path exists and is accessible"
        }
      }));
      process.exit(1);
    }
    
    try {
      const params = JSON.parse(paramsJson);
      
      // Simplified approach: directly call the appropriate function based on operation
      try {
        // For machine interface, we'll execute the operation and return JSON
        switch(operation) {
          case 'demo':
            runDemo(params, params.output);
            break;
          case 'inspect':
            runInspect([null, params.file]);
            break;
          case 'create':
            runCreate(params, [null, params.subCommand], params.output);
            break;
          case 'replace':
            runReplace([null, params.old, params.new, params.file], params.output);
            break;
          case 'generate':
            runGenerate(params, params.output);
            break;
          case 'template':
            runTemplate(params, [null, params.templateType], params.output);
            break;
          case 'bind':
            runBind(params);
            break;
          case 'export':
            runExport(params);
            break;
          case 'delete-slide':
            runDeleteSlide(params, params.output);
            break;
          case 'duplicate-slide':
            runDuplicateSlide(params, params.output);
            break;
          case 'move-slide':
            runMoveSlide(params, params.output);
            break;
          case 'slide-count':
            runSlideCount(params);
            // Note: slide-count prints directly to console, so we don't return anything here
            return;
          case 'replace-global':
            runReplaceGlobal(params, params.output);
            break;
          case 'set-background':
            runSetBackground(params, params.output);
            break;
          case 'add-image':
            runAddImage(params, params.output);
            break;
          case 'add-slide':
            runAddSlide(params, params.output);
            break;
          case 'read-file':
            runReadFile(params);
            return;
          case 'write-file':
            runWriteFile(params);
            break;
          case 'edit-file':
            runEditFile(params);
            break;
          case 'grep':
            runGrep(params);
            return;
          case 'glob':
            runGlob(params);
            return;
          case 'list-dir':
            runListDir(params);
            return;
          case 'create-dir':
            runCreateDir(params);
            break;
          default:
            console.error(JSON.stringify({
              success: false,
              error: {
                code: "UNKNOWN_OPERATION",
                message: `Unknown operation: ${operation}`,
                recovery: "Available operations: demo, inspect, create, replace, generate, template, bind, export, delete-slide, duplicate-slide, move-slide, slide-count, replace-global, set-background, add-image, add-slide"
              }
            }));
            process.exit(1);
        }
        
        // If we reach here, operation was successful
        console.log(JSON.stringify({
          success: true,
          data: {}
        }));
      } catch (error) {
        console.error(JSON.stringify({
          success: false,
          error: {
            code: "OPERATION_ERROR",
            message: error.message,
            recovery: "Check the operation parameters and ensure the workspace is accessible"
          }
        }));
        process.exit(1);
      }
    } catch (parseError) {
      console.error(JSON.stringify({
        success: false,
        error: {
          code: "JSON_PARSE_ERROR",
          message: "Failed to parse JSON parameters",
          recovery: "Ensure parameters are valid JSON"
        }
      }));
      process.exit(1);
    }
  } else {
    // Regular CLI mode - preserve existing behavior
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

      case 'delete-slide':
        runDeleteSlide(flags, outputName);
        break;

      case 'duplicate-slide':
        runDuplicateSlide(flags, outputName);
        break;

      case 'move-slide':
        runMoveSlide(flags, outputName);
        break;

      case 'slide-count':
        runSlideCount(flags);
        break;

      case 'replace-global':
        runReplaceGlobal(flags, outputName);
        break;

      case 'set-background':
        runSetBackground(flags, outputName);
        break;

      case 'add-image':
        runAddImage(flags, outputName);
        break;

      case 'add-slide':
        runAddSlide(flags, outputName);
        break;

      case 'read-file':
        runReadFile(flags);
        break;

      case 'write-file':
        runWriteFile(flags);
        break;

      case 'edit-file':
        runEditFile(flags);
        break;

      case 'grep':
        runGrep(flags);
        break;

      case 'glob':
        runGlob(flags);
        break;

      case 'list-dir':
        runListDir(flags);
        break;

      case 'create-dir':
        runCreateDir(flags);
        break;

      case 'help':
      default:
        printHelp();
    }
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

  // Import / modifier
  PptModifier,
  PptReader,

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

  // CLI handlers (also usable programmatically)
  runDeleteSlide,
  runDuplicateSlide,
  runMoveSlide,
  runSlideCount,
  runReplaceGlobal,
  runSetBackground,
  runAddImage,
  runAddSlide,
  runReadFile,
  runWriteFile,
  runEditFile,
  runGrep,
  runGlob,
  runListDir,
  runCreateDir,

  // Main function
  main
};