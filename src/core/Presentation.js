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

    const resolvedTitle = importedPresentation.title
      || importedPresentation.filePath
      || 'Imported Presentation';

    const presentation = new Presentation(resolvedTitle);
    presentation._importModel = importedPresentation;
    presentation._modifier = new PptModifier(importedPresentation);

    // The constructor defaults `slides` to []; populate it with the
    // slides actually parsed out of the opened file so callers like
    // `inspect` (slide count, slide titles) reflect the real file.
    presentation.slides = Array.isArray(importedPresentation.slides)
      ? importedPresentation.slides
      : [];

    return presentation;
  }

  /**
   * Add a slide to the presentation
   * @param {Object} [slideInstance] - Optional internal Slide tracking instance
   * @returns {Object} The created slide canvas
   */
  addSlide(slideInstance) {
    const newSlide = this.pptx.addSlide();
    
    // Attach default track context; it will be overwritten dynamically on save
    newSlide.title = (slideInstance && slideInstance.title) 
      ? slideInstance.title 
      : `Slide ${this.slides.length + 1}`;
      
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
   * Deeply parse slide rendering elements to extract structural titles right before saving.
   * This guarantees that layouts built on raw canvas elements are properly read.
   */
  _finalizeSlideTitles() {
    this.slides.forEach((slide, idx) => {
      // If the slide already has a title that isn't a default placeholder, keep it
      if (slide.title && !slide.title.startsWith('Slide ')) {
        return;
      }

      // Check if any slide elements exist
      if (Array.isArray(slide.slideObjects)) {
        // Look for the first prominent text element rendered near the top of the slide
        const titleCandidate = slide.slideObjects.find(obj => {
          if (obj.type !== 'text' || !obj.text) return false;
          
          const textOptions = obj.options || {};
          const isBold = textOptions.bold || (textOptions.fontFace && textOptions.fontFace.includes('Black'));
          const isLarge = textOptions.fontSize && textOptions.fontSize >= 24;
          const isNearTop = textOptions.y === undefined || textOptions.y < 1.5;

          return (isBold || isLarge) && isNearTop;
        });

        if (titleCandidate && typeof titleCandidate.text === 'string') {
          slide.title = titleCandidate.text.trim();
        }
      }

      // Fallback fallback if no visual header components match our bounds rules
      if (!slide.title) {
        slide.title = `Slide ${idx + 1}`;
      }
    });
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

     // Run structural normalization to capture slide element names deterministically
     this._finalizeSlideTitles();

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

   /**
    * Insert a slide at the specified index
    * @param {number} index - The index where the slide should be inserted
    * @returns {Object} The created slide canvas
    */
   insertSlide(index) {
     // Validate index
     if (typeof index !== 'number' || index < 0 || index > this.slides.length) {
       throw new Error(`Invalid index: ${index}. Must be between 0 and ${this.slides.length}`);
     }
     
     // Create a new slide
     const newSlide = this.pptx.addSlide();
     
     // Set default title
     newSlide.title = `Slide ${this.slides.length + 1}`;
     
     // Insert the slide at the specified position
     this.slides.splice(index, 0, newSlide);
     
     return newSlide;
   }

   /**
    * Delete a slide by its index
    * @param {number} slideIndex - The index of the slide to delete
    * @returns {boolean} True if slide was deleted, false otherwise
    */
   deleteSlide(slideIndex) {
     // Validate slide index
     if (typeof slideIndex !== 'number' || slideIndex < 0 || slideIndex >= this.slides.length) {
       return false;
     }
     
     // Remove the slide from our internal array
     this.slides.splice(slideIndex, 1);
     
     // Note: The underlying pptxgenjs library doesn't support direct slide deletion
     // So we can only maintain our internal tracking. Actual deletion would require
     // recreating the presentation from scratch, which is beyond this scope.
     return true;
   }

   /**
    * Duplicate a slide by its index
    * @param {number} slideIndex - The index of the slide to duplicate
    * @returns {Object} The duplicated slide canvas
    */
   duplicateSlide(slideIndex) {
     // Validate slide index
     if (typeof slideIndex !== 'number' || slideIndex < 0 || slideIndex >= this.slides.length) {
       throw new Error(`Invalid slide index: ${slideIndex}`);
     }
     
     // Get the slide to duplicate
     const slideToDuplicate = this.slides[slideIndex];
     
     // Create a new slide
     const newSlide = this.pptx.addSlide();
     
     // Copy title from original slide
     newSlide.title = slideToDuplicate.title ? `${slideToDuplicate.title} (Copy)` : `Slide ${this.slides.length + 1}`;
     
     // Insert the duplicated slide right after the original
     this.slides.splice(slideIndex + 1, 0, newSlide);
     
     return newSlide;
   }

   /**
    * Move a slide from one position to another
    * @param {number} fromIndex - The current index of the slide
    * @param {number} toIndex - The target index for the slide
    * @returns {boolean} True if slide was moved, false otherwise
    */
   moveSlide(fromIndex, toIndex) {
     // Validate indices
     if (typeof fromIndex !== 'number' || fromIndex < 0 || fromIndex >= this.slides.length ||
         typeof toIndex !== 'number' || toIndex < 0 || toIndex >= this.slides.length) {
       return false;
     }
     
     // If moving to the same position, do nothing
     if (fromIndex === toIndex) {
       return true;
     }
     
     // Extract the slide
     const slide = this.slides.splice(fromIndex, 1)[0];
     
     // Insert at new position
     this.slides.splice(toIndex, 0, slide);
     
     return true;
   }

   /**
    * Hide a slide by its index
    * @param {number} slideIndex - The index of the slide to hide
    * @returns {boolean} True if slide was hidden, false otherwise
    */
   hideSlide(slideIndex) {
     // Validate slide index
     if (typeof slideIndex !== 'number' || slideIndex < 0 || slideIndex >= this.slides.length) {
       return false;
     }
     
     // For now, we'll just mark the slide as hidden in our internal tracking
     // Note: The underlying library doesn't support hiding slides directly
     // This is a placeholder that maintains the API contract
     const slide = this.slides[slideIndex];
     if (slide) {
       slide.hidden = true;
       return true;
     }
     
     return false;
   }
 }

module.exports = Presentation;