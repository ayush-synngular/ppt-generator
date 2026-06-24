'use strict';

const fs = require('fs');

class PptModifier {
  constructor(presentation) {
    if (!presentation || typeof presentation !== 'object') {
      throw new Error('PptModifier requires a valid presentation object.');
    }

    this.presentation = presentation;
    this.replacements = [];
  }

  getSlides() {
    return Array.isArray(this.presentation.slides) ? this.presentation.slides : [];
  }

  getSlide(slideNumber) {
    if (!Number.isInteger(slideNumber) || slideNumber < 1) {
      return undefined;
    }

    return this.getSlides().find((slide) => slide.slideNumber === slideNumber);
  }

  replaceText(oldText, newText) {
    if (typeof oldText !== 'string' || typeof newText !== 'string') {
      return;
    }

    this.getSlides().forEach((slide) => {
      const replacements = this._replaceTextInSlideObject(slide, oldText, newText);
      if (replacements > 0) {
        this.replacements.push({ slideNumber: slide.slideNumber, oldText, newText });
      }
    });
  }

  replaceTextInSlide(slideNumber, oldText, newText) {
    if (typeof oldText !== 'string' || typeof newText !== 'string') {
      return;
    }

    const slide = this.getSlide(slideNumber);
    if (!slide) {
      return;
    }

    const replacements = this._replaceTextInSlideObject(slide, oldText, newText);
    if (replacements > 0) {
      this.replacements.push({ slideNumber, oldText, newText });
    }
  }

  findText(text) {
    if (typeof text !== 'string') {
      return [];
    }

    const matches = [];
    this.getSlides().forEach((slide) => {
      if (Array.isArray(slide.text)) {
        slide.text.forEach((segment) => {
          if (segment === text) {
            matches.push({ slideNumber: slide.slideNumber, text: segment });
          }
        });
      }

      if (Array.isArray(slide.shapes)) {
        slide.shapes.forEach((shape) => {
          if (shape && typeof shape.text === 'string' && shape.text === text) {
            matches.push({ slideNumber: slide.slideNumber, text: shape.text });
          }
        });
      }
    });

    return matches;
  }

  toJSON() {
    if (typeof this.presentation.toJSON === 'function') {
      return this.presentation.toJSON();
    }

    return {
      ...this.presentation,
      slides: this.getSlides().map((slide) => (typeof slide.toJSON === 'function' ? slide.toJSON() : slide))
    };
  }

  saveJSON(filePath) {
    if (typeof filePath !== 'string' || filePath.trim() === '') {
      throw new Error('A valid file path is required to save JSON.');
    }

    const json = this.toJSON();
    fs.writeFileSync(filePath, JSON.stringify(json, null, 2), 'utf8');
  }

  async save(filePath) {
    if (typeof filePath !== 'string' || filePath.trim() === '') {
      throw new Error('A valid file path is required to save PPTX.');
    }

    const AdmZip = require('adm-zip');

    const originalFile = this.presentation.filePath;
    if (!originalFile || typeof originalFile !== 'string' || !fs.existsSync(originalFile)) {
      throw new Error('Original PPTX file is required to preserve layout and media.');
    }

    const zip = new AdmZip(originalFile);
    const slideEntries = zip
      .getEntries()
      .filter((entry) => /^ppt\/slides\/slide\d+\.xml$/.test(entry.entryName));

    if (this.replacements.length === 0) {
      zip.writeZip(filePath);
      return;
    }

    const escapeRegExp = (value) => value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

    for (const entry of slideEntries) {
      let xml = entry.getData().toString('utf8');
      this.replacements.forEach(({ oldText, newText }) => {
        const pattern = new RegExp(`<a:t>${escapeRegExp(oldText)}</a:t>`, 'g');
        xml = xml.replace(pattern, `<a:t>${newText}</a:t>`);
      });
      zip.updateFile(entry.entryName, Buffer.from(xml, 'utf8'));
    }

    zip.writeZip(filePath);
  }

  _replaceTextInSlideObject(slide, oldText, newText) {
    if (!slide || typeof slide !== 'object') {
      return 0;
    }

    let replacements = 0;

    if (Array.isArray(slide.text)) {
      slide.text = slide.text.map((segment) => {
        if (segment === oldText) {
          replacements += 1;
          return newText;
        }
        return segment;
      });
    }

    if (Array.isArray(slide.shapes)) {
      slide.shapes = slide.shapes.map((shape) => {
        if (shape && typeof shape.text === 'string' && shape.text === oldText) {
          replacements += 1;
          return { ...shape, text: newText };
        }
        return shape;
      });
    }

    return replacements;
  }

  _applyTextReplacements(parsedSlide) {
    const target = parsedSlide;
    const oldNewPairs = this.replacements;

    const replaceTextInNode = (node) => {
      if (!node || typeof node !== 'object') {
        return;
      }

      if (node['a:t'] && typeof node['a:t'] === 'string') {
        oldNewPairs.forEach(({ oldText, newText }) => {
          if (node['a:t'] === oldText) {
            node['a:t'] = newText;
          }
        });
      }

      Object.values(node).forEach((child) => {
        if (Array.isArray(child)) {
          child.forEach(replaceTextInNode);
        } else if (typeof child === 'object') {
          replaceTextInNode(child);
        }
      });
    };

    replaceTextInNode(target);
    return target;
  }
}

module.exports = PptModifier;
