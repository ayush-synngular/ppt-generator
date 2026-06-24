'use strict';

/**
 * JSON-driven Slide Builder.
 *
 * Supported slide types:
 * - title-slide
 * - two-column
 * - comparison
 *
 * Config schema:
 * {
 *   slides: [
 *     { type: 'title-slide', title: 'Welcome', subtitle: 'Demo' },
 *     {
 *       type: 'two-column',
 *       title: 'Overview',
 *       leftContent: ['Point 1', 'Point 2'],
 *       rightContent: ['Point A', 'Point B']
 *     }
 *   ]
 * }
 *
 * Error behavior:
 * - Throws when the top-level config is invalid.
 * - Throws when a slide object is missing a valid type.
 * - Throws for unknown slide types.
 */
const { createTitleSlide } = require('../layouts/title-slide');
const { createTwoColumnSlide } = require('../layouts/two-column');
const { createComparisonSlide } = require('../layouts/comparison');

class SlideBuilder {
  /**
   * @param {Object} presentation - A Presentation instance.
   * @param {Object} theme - A Theme instance to apply to slides.
   */
  constructor(presentation, theme) {
    if (!presentation || typeof presentation.addSlide !== 'function') {
      throw new Error('SlideBuilder requires a valid presentation instance.');
    }

    this.presentation = presentation;
    this.theme = theme;
  }

  /**
   * Build slides from a JSON configuration.
   * @param {Object} config - Slide builder configuration.
   * @returns {Object[]} Created slides.
   */
  build(config) {
    if (!config || typeof config !== 'object') {
      throw new Error('SlideBuilder.build requires a configuration object.');
    }

    if (!Array.isArray(config.slides)) {
      throw new Error('SlideBuilder configuration must include a slides array.');
    }

    return config.slides.map((slideConfig, index) => {
      if (!slideConfig || typeof slideConfig !== 'object') {
        throw new Error(`SlideBuilder slide at index ${index} must be an object.`);
      }

      return this.buildSlide(slideConfig);
    });
  }

  /**
   * Build a single slide from a slide configuration object.
   * @param {Object} slideConfig - Slide configuration.
   * @returns {Object} Created slide.
   */
  buildSlide(slideConfig) {
    const { type } = slideConfig;

    if (!type || typeof type !== 'string') {
      throw new Error('SlideBuilder slide configuration must include a valid type string.');
    }

    const options = { ...slideConfig, theme: slideConfig.theme || this.theme };

    switch (type) {
      case 'title-slide':
        return createTitleSlide(this.presentation, options);
      case 'two-column':
        return createTwoColumnSlide(this.presentation, options);
      case 'comparison':
        return createComparisonSlide(this.presentation, options);
      default:
        throw new Error(`SlideBuilder does not support slide type '${type}'. Supported types are: title-slide, two-column, comparison.`);
    }
  }
}

module.exports = SlideBuilder;
