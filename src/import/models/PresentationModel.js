'use strict';

class PresentationModel {
  constructor({ filePath = '', metadata = {}, slides = [] } = {}) {
    this.filePath = filePath;
    this.metadata = metadata;
    this.slides = slides;
  }

  addSlide(slide) {
    this.slides.push(slide);
  }

  getSlideCount() {
    return this.slides.length;
  }

  toJSON() {
    return {
      filePath: this.filePath,
      metadata: this.metadata,
      slides: this.slides.map((slide) => (typeof slide.toJSON === 'function' ? slide.toJSON() : slide))
    };
  }
}

module.exports = PresentationModel;
