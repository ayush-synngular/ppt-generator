'use strict';

/**
 * Slide class representing a single slide in a presentation
 */
class Slide {
  constructor(title = 'Untitled Slide') {
    this.title = title;
    this.elements = [];
    this.layout = null;
    this.theme = null;
  }

  /**
   * Add an element to the slide
   * @param {Object} element - The element to add
   */
  addElement(element) {
    this.elements.push(element);
  }

  /**
   * Remove an element from the slide
   * @param {number} index - Index of the element to remove
   */
  removeElement(index) {
    if (index >= 0 && index < this.elements.length) {
      this.elements.splice(index, 1);
    }
  }

  /**
   * Set the layout for the slide
   * @param {Object} layout - The layout to apply
   */
  setLayout(layout) {
    this.layout = layout;
  }

  /**
   * Set the theme for the slide
   * @param {Theme} theme - The theme to apply
   */
  setTheme(theme) {
    this.theme = theme;
  }
}

module.exports = Slide;