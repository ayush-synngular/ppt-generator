'use strict';

// const pptx = require('pptxgenjs');
// const ShapeType = pptx.ShapeType;

function resolveCoordinates(options) {
  const { x1, y1, x2, y2, x = 1, y = 1, w = 2, h = 2 } = options;

  if (x1 !== undefined && y1 !== undefined && x2 !== undefined && y2 !== undefined) {
    const minX = Math.min(x1, x2);
    const minY = Math.min(y1, y2);
    const width = Math.max(Math.abs(x2 - x1), 0.01);
    const height = Math.max(Math.abs(y2 - y1), 0.01);
    return { x: minX, y: minY, w: width, h: height };
  }

  return { x, y, w, h };
}

function buildShapeOptions(options = {}) {
  const { fillColor, lineColor = '000000', lineWidth = 1, transparency, rotation } = options;
  const coords = resolveCoordinates(options);

  const shapeOptions = {
    ...coords,
    line: {
      type: 'solid',
      color: lineColor,
      width: lineWidth
    }
  };

  if (fillColor !== undefined) {
    shapeOptions.fill = { color: fillColor };
  }

  if (transparency !== undefined) {
    if (!shapeOptions.fill) {
      shapeOptions.fill = { transparency };
    } else {
      shapeOptions.fill.transparency = transparency;
    }
    shapeOptions.line.transparency = transparency;
  }

  if (rotation !== undefined) {
    shapeOptions.rotate = rotation;
  }

  return shapeOptions;
}

function addRectangle(slide, options = {}) {
  const shapeOptions = buildShapeOptions(options);
  return slide.addShape('rect', shapeOptions);
}

function addCircle(slide, options = {}) {
  const shapeOptions = buildShapeOptions(options);
  return slide.addShape('ellipse', shapeOptions);
}

function addLine(slide, options = {}) {
  const shapeOptions = buildShapeOptions(options);
  return slide.addShape('line', shapeOptions);
}

function addArrow(slide, options = {}) {
  const shapeOptions = buildShapeOptions(options);
  shapeOptions.line.beginArrowType = 'none';
  shapeOptions.line.endArrowType = 'triangle';
  return slide.addShape('line', shapeOptions);
}

module.exports = {
  addRectangle,
  addCircle,
  addLine,
  addArrow
};
