'use strict';

/**
 * Color utility functions
 */
class ColorUtils {
  /**
   * Convert hex color to RGB
   * @param {string} hex - Hex color value
   * @returns {Object} RGB object with r, g, b properties
   */
  static hexToRgb(hex) {
    const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
    return result ? {
      r: parseInt(result[1], 16),
      g: parseInt(result[2], 16),
      b: parseInt(result[3], 16)
    } : null;
  }

  /**
   * Generate a random color
   * @returns {string} Random hex color
   */
  static randomColor() {
    return '#' + Math.floor(Math.random()*16777215).toString(16);
  }

  /**
   * Check if a color is dark
   * @param {string} hex - Hex color value
   * @returns {boolean} True if color is dark
   */
  static isDarkColor(hex) {
    const rgb = this.hexToRgb(hex);
    if (!rgb) return false;
    
    // Calculate brightness (using luminance formula)
    const brightness = (rgb.r * 299 + rgb.g * 587 + rgb.b * 114) / 1000;
    return brightness < 128;
  }
}

module.exports = ColorUtils;