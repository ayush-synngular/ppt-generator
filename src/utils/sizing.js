'use strict';

/**
 * Sizing utility functions
 */
class SizingUtils {
  /**
   * Scale a dimension proportionally
   * @param {number} originalSize - Original size
   * @param {number} scaleFactor - Scale factor
   * @returns {number} Scaled size
   */
  static scale(originalSize, scaleFactor) {
    return originalSize * scaleFactor;
  }

  /**
   * Fit content within bounds
   * @param {Object} contentSize - Content size (width, height)
   * @param {Object} bounds - Bounds size (width, height)
   * @returns {Object} Scaled size that fits within bounds
   */
  static fitWithin(contentSize, bounds) {
    const widthRatio = bounds.width / contentSize.width;
    const heightRatio = bounds.height / contentSize.height;
    const scale = Math.min(widthRatio, heightRatio);
    
    return {
      width: contentSize.width * scale,
      height: contentSize.height * scale
    };
  }

  /**
   * Calculate aspect ratio
   * @param {number} width - Width
   * @param {number} height - Height
   * @returns {number} Aspect ratio (width/height)
   */
  static aspectRatio(width, height) {
    return width / height;
  }
}

module.exports = SizingUtils;