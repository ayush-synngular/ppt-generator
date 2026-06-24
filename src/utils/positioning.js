'use strict';

/**
 * Positioning utility functions
 */
class PositioningUtils {
  /**
   * Center an element within container
   * @param {Object} elementSize - Element size (width, height)
   * @param {Object} containerSize - Container size (width, height)
   * @returns {Object} Centered position (x, y)
   */
  static center(elementSize, containerSize) {
    return {
      x: (containerSize.width - elementSize.width) / 2,
      y: (containerSize.height - elementSize.height) / 2
    };
  }

  /**
   * Position element at top left
   * @param {Object} elementSize - Element size (width, height)
   * @param {Object} containerSize - Container size (width, height)
   * @returns {Object} Top-left position (x, y)
   */
  static topLeft(elementSize, containerSize) {
    return { x: 0, y: 0 };
  }

  /**
   * Position element at top right
   * @param {Object} elementSize - Element size (width, height)
   * @param {Object} containerSize - Container size (width, height)
   * @returns {Object} Top-right position (x, y)
   */
  static topRight(elementSize, containerSize) {
    return {
      x: containerSize.width - elementSize.width,
      y: 0
    };
  }

  /**
   * Position element at bottom left
   * @param {Object} elementSize - Element size (width, height)
   * @param {Object} containerSize - Container size (width, height)
   * @returns {Object} Bottom-left position (x, y)
   */
  static bottomLeft(elementSize, containerSize) {
    return {
      x: 0,
      y: containerSize.height - elementSize.height
    };
  }

  /**
   * Position element at bottom right
   * @param {Object} elementSize - Element size (width, height)
   * @param {Object} containerSize - Container size (width, height)
   * @returns {Object} Bottom-right position (x, y)
   */
  static bottomRight(elementSize, containerSize) {
    return {
      x: containerSize.width - elementSize.width,
      y: containerSize.height - elementSize.height
    };
  }
}

module.exports = PositioningUtils;