'use strict';

class SlideModel {
  constructor({ slideNumber = 0, text = [], shapes = [] } = {}) {
    this.slideNumber = slideNumber;
    this.text = Array.isArray(text) ? text : [];
    this.shapes = Array.isArray(shapes) ? shapes : [];
  }

  addShape(shape) {
    this.shapes.push(shape);
  }

  toJSON() {
    return {
      slideNumber: this.slideNumber,
      text: this.text,
      shapes: this.shapes.map((shape) => (typeof shape.toJSON === 'function' ? shape.toJSON() : shape))
    };
  }
}

module.exports = SlideModel;
