'use strict';

/**
 * Theme class for controlling presentation styling and layout defaults.
 */
class Theme {
  /**
   * @param {Object} config - Theme configuration values.
   * @param {string} config.primaryColor - Primary brand color.
   * @param {string} config.secondaryColor - Secondary brand color.
   * @param {string} config.accentColor - Accent color for highlights.
   * @param {string} config.backgroundColor - Slide background color.
   * @param {string} config.textColor - Default body text color.
   * @param {string} config.titleFont - Font family for titles.
   * @param {string} config.bodyFont - Font family for body text.
   * @param {number} config.titleFontSize - Font size for titles.
   * @param {number} config.bodyFontSize - Font size for body text.
   * @param {number} config.slideMargin - Default slide margin in inches.
   */
  constructor(config = {}) {
    this.primaryColor = config.primaryColor || '#1F4E79';
    this.secondaryColor = config.secondaryColor || '#7A7A7A';
    this.accentColor = config.accentColor || '#FFC000';
    this.backgroundColor = config.backgroundColor || '#FFFFFF';
    this.textColor = config.textColor || '#333333';
    this.titleFont = config.titleFont || 'Arial Black';
    this.bodyFont = config.bodyFont || 'Arial';
    this.titleFontSize = typeof config.titleFontSize === 'number' ? config.titleFontSize : 32;
    this.bodyFontSize = typeof config.bodyFontSize === 'number' ? config.bodyFontSize : 18;
    this.slideMargin = typeof config.slideMargin === 'number' ? config.slideMargin : 0.5;
  }

  /**
   * Get a theme color by name.
   * @param {string} name - One of primaryColor, secondaryColor, accentColor, backgroundColor, textColor.
   * @returns {string|undefined} The configured color value.
   */
  getColor(name) {
    switch (name) {
      case 'primaryColor':
        return this.primaryColor;
      case 'secondaryColor':
        return this.secondaryColor;
      case 'accentColor':
        return this.accentColor;
      case 'backgroundColor':
        return this.backgroundColor;
      case 'textColor':
        return this.textColor;
      default:
        return undefined;
    }
  }

  /**
   * Get a font family by type.
   * @param {string} type - One of titleFont or bodyFont.
   * @returns {string|undefined} The configured font family.
   */
  getFont(type) {
    switch (type) {
      case 'titleFont':
        return this.titleFont;
      case 'bodyFont':
        return this.bodyFont;
      default:
        return undefined;
    }
  }

  /**
   * Get a font size by type.
   * @param {string} type - One of titleFontSize or bodyFontSize.
   * @returns {number|undefined} The configured font size.
   */
  getFontSize(type) {
    switch (type) {
      case 'titleFontSize':
        return this.titleFontSize;
      case 'bodyFontSize':
        return this.bodyFontSize;
      default:
        return undefined;
    }
  }

  /**
   * Get the default slide margin.
   * @returns {number} Slide margin in inches.
   */
  getSlideMargin() {
    return this.slideMargin;
  }

  /**
   * Convert the theme configuration to a plain object.
   * @returns {Object} Theme JSON object.
   */
  toJSON() {
    return {
      primaryColor: this.primaryColor,
      secondaryColor: this.secondaryColor,
      accentColor: this.accentColor,
      backgroundColor: this.backgroundColor,
      textColor: this.textColor,
      titleFont: this.titleFont,
      bodyFont: this.bodyFont,
      titleFontSize: this.titleFontSize,
      bodyFontSize: this.bodyFontSize,
      slideMargin: this.slideMargin
    };
  }
}

module.exports = Theme;
