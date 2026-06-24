'use strict';

const {
  Presentation,
  Theme
} = require('../src/index');
const AutoLayoutEngine = require('../src/core/AutoLayoutEngine');
const { addText } = require('../src/elements/text');
const { addRectangle } = require('../src/elements/shape');

function renderSlideHeader(slide, title, theme) {
  addText(slide, title, {
    x: 0.5,
    y: 0.25,
    w: 12.33,
    h: 0.5,
    fontFace: theme.getFont('titleFont'),
    fontSize: theme.getFontSize('titleFontSize'),
    bold: true,
    align: 'center'
  });
}

function renderContentBounds(slide, contentArea) {
  addRectangle(slide, {
    x: contentArea.x,
    y: contentArea.y,
    w: contentArea.w,
    h: contentArea.h,
    fillColor: '#F4F4F4',
    lineColor: '#CCCCCC',
    lineWidth: 1
  });
  addText(slide, 'Content Area', {
    x: contentArea.x + 0.1,
    y: contentArea.y + 0.1,
    w: 3,
    h: 0.3,
    fontSize: 12,
    color: '#333333'
  });
}

function logLayout(name, items) {
  console.log(`=== ${name} ===`);
  items.forEach((item, index) => {
    console.log(`item ${index + 1}: ${JSON.stringify(item)}`);
  });
}

async function testAutoLayout() {
  const theme = new Theme({ slideMargin: 0.75 });
  const engine = new AutoLayoutEngine(theme);
  const ppt = new Presentation('Auto Layout Demo');

  const slideWidth = 13.33;
  const slideHeight = 7.5;
  const contentArea = engine.getContentArea(slideWidth, slideHeight);

  // Slide 1: Content Area
  const slide1 = ppt.addSlide();
  renderSlideHeader(slide1, 'Content Area', theme);
  renderContentBounds(slide1, contentArea);

  // Slide 2: 2 Column Layout
  const slide2 = ppt.addSlide();
  renderSlideHeader(slide2, '2 Column Layout', theme);
  renderContentBounds(slide2, contentArea);
  const twoColumns = engine.calculateColumnLayout(2, contentArea);
  logLayout('2 Column Layout', twoColumns);
  twoColumns.forEach((col, index) => {
    addRectangle(slide2, { x: col.x, y: col.y + 0.5, w: col.w, h: 1.1, fillColor: '#D9EDF7', lineColor: '#0B5394' });
    addText(slide2, `Col ${index + 1}`, { x: col.x + 0.05, y: col.y + 0.55, w: col.w - 0.1, h: 0.4, fontSize: 12 });
  });

  // Slide 3: 3 Column Layout
  const slide3 = ppt.addSlide();
  renderSlideHeader(slide3, '3 Column Layout', theme);
  renderContentBounds(slide3, contentArea);
  const threeColumns = engine.calculateColumnLayout(3, contentArea);
  logLayout('3 Column Layout', threeColumns);
  threeColumns.forEach((col, index) => {
    addRectangle(slide3, { x: col.x, y: col.y + 0.5, w: col.w, h: 1.1, fillColor: '#E9F7DF', lineColor: '#38761D' });
    addText(slide3, `Col ${index + 1}`, { x: col.x + 0.05, y: col.y + 0.55, w: col.w - 0.1, h: 0.4, fontSize: 12 });
  });

  // Slide 4: Grid Layout
  const slide4 = ppt.addSlide();
  renderSlideHeader(slide4, 'Grid Layout', theme);
  renderContentBounds(slide4, contentArea);
  const gridItems = Array.from({ length: 6 }, (_, index) => ({ id: index + 1 }));
  const grid = engine.calculateGridLayout(gridItems, contentArea);
  logLayout('Grid Layout', grid);
  grid.forEach((cell, index) => {
    addRectangle(slide4, { x: cell.x, y: cell.y, w: cell.w, h: cell.h, fillColor: '#FCE5CD', lineColor: '#E69138' });
    addText(slide4, `Cell ${index + 1}`, { x: cell.x + 0.05, y: cell.y + 0.05, w: cell.w - 0.1, h: 0.3, fontSize: 10 });
  });

  // Slide 5: Vertical Stack Layout
  const slide5 = ppt.addSlide();
  renderSlideHeader(slide5, 'Vertical Stack Layout', theme);
  renderContentBounds(slide5, contentArea);
  const stackElements = [
    { h: 0.7 },
    { h: 0.9 },
    { h: 0.8 },
    { h: 1.2 }
  ];
  const stack = engine.calculateVerticalStack(stackElements, contentArea);
  logLayout('Vertical Stack Layout', stack);
  stack.forEach((item, index) => {
    addRectangle(slide5, { x: item.x, y: item.y, w: item.w, h: item.h, fillColor: '#DDEBF7', lineColor: '#2E75B6' });
    addText(slide5, `Item ${index + 1}`, { x: item.x + 0.05, y: item.y + 0.05, w: item.w - 0.1, h: 0.25, fontSize: 10 });
  });

  // Slide 6: Text Estimation
  const slide6 = ppt.addSlide();
  renderSlideHeader(slide6, 'Text Height Estimation', theme);
  renderContentBounds(slide6, contentArea);
  const sampleText = 'This is a sample text estimate that should wrap across multiple lines depending on width.';
  const textWidth = contentArea.w;
  const estimatedHeight = engine.estimateTextHeight(sampleText, textWidth, 14);
  const textBounds = { x: contentArea.x, y: contentArea.y + 0.5, w: textWidth, h: Math.min(estimatedHeight, contentArea.h - 0.5) };
  logLayout('Text Estimation', [textBounds]);
  addRectangle(slide6, { x: textBounds.x, y: textBounds.y, w: textBounds.w, h: textBounds.h, fillColor: '#F4F4F4', lineColor: '#999999' });
  addText(slide6, `Estimated height: ${estimatedHeight.toFixed(2)} in`, { x: textBounds.x + 0.05, y: textBounds.y + 0.05, w: textBounds.w - 0.1, h: 0.3, fontSize: 10 });

  await ppt.save('test-auto-layout.pptx');
  console.log('Presentation saved successfully as test-auto-layout.pptx');
}

testAutoLayout().catch(error => {
  console.error('Error creating auto layout presentation:', error);
});
