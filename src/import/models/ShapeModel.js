'use strict';

class ShapeModel {
  constructor({ type = 'shape', text = '', bounds = {} } = {}) {
    this.type = type;
    this.text = typeof text === 'string' ? text : '';
    this.bounds = bounds;
  }

  toJSON() {
    return {
      type: this.type,
      text: this.text,
      bounds: this.bounds
    };
  }
}

module.exports = ShapeModel;
