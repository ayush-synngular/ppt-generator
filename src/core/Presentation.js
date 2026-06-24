'use strict';

const pptx = require('pptxgenjs');

/**
 * Presentation class representing a PowerPoint presentation
 */
class Presentation {
  constructor(title = 'Untitled Presentation') {
    // Create a new pptx instance
    this.pptx = new pptx();
    
    // Set presentation properties
    this.pptx.layout = "LAYOUT_WIDE";
    this.pptx.author = "PPT Generator Utility";
    this.pptx.company = "Internal";
    
    // Store the title
    this.title = title;
    this.slides = [];
    this._importModel = null;
    this._modifier = null;
  }

  static async open(filePath) {
    const PptReader = require('../import/PptReader');
    const PptModifier = require('../import/PptModifier');

    const reader = new PptReader();
    const importedPresentation = await reader.read(filePath);

    const presentation = new Presentation(importedPresentation.filePath || 'Imported Presentation');
    presentation._importModel = importedPresentation;
    presentation._modifier = new PptModifier(importedPresentation);
    return presentation;
  }

  /**
   * Add a slide to the presentation
   * @param {Slide} slide - The slide to add
   * @returns {Object} The created slide
   */
  addSlide(slide) {
    // Create a new slide in pptxgenjs
    const newSlide = this.pptx.addSlide();
    this.slides.push(newSlide);
    return newSlide;
  }

  replaceText(oldText, newText) {
    if (!this._modifier || typeof this._modifier.replaceText !== 'function') {
      throw new Error('Presentation.replaceText is only available for opened presentations.');
    }

    this._modifier.replaceText(oldText, newText);
  }

  slide(index) {
    if (!this._importModel || !Array.isArray(this._importModel.slides)) {
      return undefined;
    }

    if (!Number.isInteger(index) || index < 1) {
      return undefined;
    }

    return this._importModel.slides.find((slide) => slide.slideNumber === index);
  }

  /**
   * NOTE: PptxGenJS does not expose a public slide deletion API.
   * Slides cannot be removed once created via the public library surface.
   * To remove slides, create only the slides that should exist in the final
   * presentation.
   */

  /**
   * Set the title for the presentation
   * @param {string} title - The presentation title
   */
  setTitle(title) {
    this.title = title;
    this.pptx.title = title;
  }

  /**
   * Set the subject for the presentation
   * @param {string} subject - The presentation subject
   */
  setSubject(subject) {
    this.pptx.subject = subject;
  }

  /**
   * Save the presentation to a file
   * @param {string} fileName - The name of the file to save
   * @returns {Promise<void>} Promise that resolves when saving is complete
   */
  async save(fileName) {
    if (this._importModel && this._modifier && typeof this._modifier.save === 'function') {
      await this._modifier.save(fileName);
      return;
    }

    return new Promise((resolve, reject) => {
      this.pptx.writeFile({ fileName }, (err) => {
        if (err) {
          reject(err);
        } else {
          resolve();
        }
      });
    });
  }
}

module.exports = Presentation;
