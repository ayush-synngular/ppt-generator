'use strict';

/**
 * Auto layout engine for calculating slide content positions.
 *
 * Assumptions:
 * - Slide margins are symmetric and applied to all sides.
 * - Column and grid layouts assume a small fixed gutter spacing between items.
 * - Estimated text height is based on character count, width, and font size.
 * - Overflow detection compares element bounds to a container rectangle.
 */
class AutoLayoutEngine {
  /**
   * @param {Object} theme - Theme instance with optional slide margin.
   */
  constructor(theme = {}) {
    this.theme = theme;
    this.defaultMargin = 0.5;
    this.gutter = 0.2;
    this.lineHeightFactor = 1.2;
  }

  /**
   * Get the inner content area for a slide after applying margins.
   * @param {number} slideWidth - Total slide width in inches.
   * @param {number} slideHeight - Total slide height in inches.
   * @returns {Object} Content rectangle with x, y, w, h.
   */
  getContentArea(slideWidth, slideHeight) {
    const margin = (this.theme && typeof this.theme.getSlideMargin === 'function')
      ? this.theme.getSlideMargin()
      : this.defaultMargin;

    const x = margin;
    const y = margin;
    const w = Math.max(0, slideWidth - margin * 2);
    const h = Math.max(0, slideHeight - margin * 2);

    return { x, y, w, h };
  }

  /**
   * Calculate equal-width columns inside a content area.
   * @param {number} columns - Number of columns to place.
   * @param {Object} contentArea - Content area rectangle.
   * @returns {Object[]} Column bounds.
   */
  calculateColumnLayout(columns, contentArea) {
    const count = Math.max(1, Math.floor(columns));
    const totalGutter = this.gutter * Math.max(0, count - 1);
    const columnWidth = (contentArea.w - totalGutter) / count;
    const result = [];

    for (let index = 0; index < count; index += 1) {
      const x = contentArea.x + index * (columnWidth + this.gutter);
      result.push({ x, y: contentArea.y, w: columnWidth, h: contentArea.h });
    }

    return result;
  }

  /**
   * Calculate a simple grid layout for a list of items.
   * All generated cells are guaranteed to fit within the content area when
   * the calculated column and row gutters are respected.
   * @param {Array} items - Item list to position.
   * @param {Object} contentArea - Content area rectangle.
   * @returns {Object[]} Item bounds.
   */
  calculateGridLayout(items, contentArea) {
    const totalItems = Array.isArray(items) ? items.length : 0;
    if (totalItems === 0) {
      return [];
    }

    const columns = Math.min(3, Math.max(1, Math.ceil(Math.sqrt(totalItems))));
    const rows = Math.ceil(totalItems / columns);
    const totalHorizontalGutter = this.gutter * Math.max(0, columns - 1);
    const totalVerticalGutter = this.gutter * Math.max(0, rows - 1);
    const cellWidth = (contentArea.w - totalHorizontalGutter) / columns;
    const cellHeight = (contentArea.h - totalVerticalGutter) / rows;
    const result = [];

    for (let index = 0; index < totalItems; index += 1) {
      const row = Math.floor(index / columns);
      const column = index % columns;
      const x = contentArea.x + column * (cellWidth + this.gutter);
      const y = contentArea.y + row * (cellHeight + this.gutter);
      result.push({ x, y, w: cellWidth, h: cellHeight });
    }

    return result;
  }

  /**
   * Calculate stacked vertical positions for a list of elements.
   * The total stack height is constrained to the available content area.
   * @param {Array} elements - Elements with optional height values.
   * @param {Object} contentArea - Content area rectangle.
   * @returns {Object[]} Stacked bounds.
   */
  calculateVerticalStack(elements, contentArea) {
    const result = [];
    let currentY = contentArea.y;
    let remainingHeight = contentArea.h;

    for (const item of elements) {
      const height = (item && typeof item.h === 'number') ? item.h : 1;
      if (height > remainingHeight) {
        break;
      }

      result.push({ x: contentArea.x, y: currentY, w: contentArea.w, h: height });
      currentY += height;
      remainingHeight -= height;

      if (remainingHeight <= 0) {
        break;
      }

      currentY += this.gutter;
      remainingHeight -= this.gutter;
    }

    return result;
  }

  /**
   * Estimate text height based on width and font size.
   * The calculation is deterministic and approximates the number of lines.
   * @param {string} text - Text content.
   * @param {number} width - Available width in inches.
   * @param {number} fontSize - Font size in points.
   * @returns {number} Estimated height in inches.
   */
  estimateTextHeight(text, width, fontSize) {
    if (typeof text !== 'string' || width <= 0 || fontSize <= 0) {
      return 0;
    }

    const normalizedText = text.replace(/\s+/g, ' ').trim();
    if (!normalizedText) {
      return 0;
    }

    const averageCharWidthInches = fontSize / 72 * 0.45;
    const charsPerLine = Math.max(1, Math.floor(width / averageCharWidthInches));
    const lineCount = Math.ceil(normalizedText.length / charsPerLine);
    const lineHeightInches = fontSize / 72 * this.lineHeightFactor;

    return Math.max(lineHeightInches, lineCount * lineHeightInches);
  }

  /**
   * Detect whether a rectangular element overflows its bounds.
   * @param {Object} element - Element bounds { x, y, w, h }.
   * @param {Object} bounds - Container bounds { x, y, w, h }.
   * @returns {boolean} True if the element extends beyond the bounds.
   */
  detectOverflow(element, bounds) {
    if (!element || !bounds) {
      return false;
    }

    const elementRight = element.x + element.w;
    const elementBottom = element.y + element.h;
    const boundsRight = bounds.x + bounds.w;
    const boundsBottom = bounds.y + bounds.h;

    return (
      element.x < bounds.x ||
      element.y < bounds.y ||
      elementRight > boundsRight ||
      elementBottom > boundsBottom
    );
  }

  /**
   * Validate that bounds are fully inside a content rectangle.
   * @param {Object} item - Item bounds { x, y, w, h }.
   * @param {Object} contentArea - Container bounds { x, y, w, h }.
   * @returns {boolean} True if the item is fully contained.
   */
  validateBounds(item, contentArea) {
    if (!item || !contentArea) {
      return false;
    }

    return (
      item.x >= contentArea.x &&
      item.y >= contentArea.y &&
      item.x + item.w <= contentArea.x + contentArea.w &&
      item.y + item.h <= contentArea.y + contentArea.h
    );
  }
}

module.exports = AutoLayoutEngine;
