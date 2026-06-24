'use strict';

/**
 * Validation engine for presentation configuration objects.
 *
 * This engine performs structural validation of presentation, slide, component,
 * and theme configurations to help detect invalid or incomplete data before a
 * presentation is generated.
 */
class ValidationEngine {
  constructor() {
    this.supportedSlideTypes = [
      'title',
      'two-column',
      'comparison',
      'executive-summary',
      'quarterly-business-review',
      'project-status',
      'chart',
      'table',
      'image',
      'text'
    ];

    this.supportedChartTypes = ['bar', 'line', 'pie', 'area'];
    this.supportedTextComponents = ['title', 'text', 'bullet-list', 'metric-card', 'chart-card'];
  }

  /**
   * Validate a presentation configuration object.
   * @param {Object} config - Presentation configuration.
   * @returns {Object} Validation result.
   */
  validatePresentation(config) {
    const errors = [];
    const warnings = [];

    if (!config || typeof config !== 'object') {
      errors.push('Presentation config must be an object.');
      return { valid: false, errors, warnings };
    }

    if (!Array.isArray(config.slides)) {
      errors.push('Presentation must include a slides array.');
      return { valid: false, errors, warnings };
    }

    config.slides.forEach((slideConfig, index) => {
      const result = this.validateSlide(slideConfig);
      result.errors.forEach(error => errors.push(`Slide ${index + 1}: ${error}`));
      result.warnings.forEach(warning => warnings.push(`Slide ${index + 1}: ${warning}`));
    });

    return { valid: errors.length === 0, errors, warnings };
  }

  /**
   * Validate a slide configuration object.
   * @param {Object} slideConfig - Slide configuration.
   * @returns {Object} Validation result.
   */
  validateSlide(slideConfig) {
    const errors = [];
    const warnings = [];

    if (!slideConfig || typeof slideConfig !== 'object') {
      errors.push('Slide config must be an object.');
      return { valid: false, errors, warnings };
    }

    if (!slideConfig.type || typeof slideConfig.type !== 'string') {
      errors.push('Slide must have a type string.');
    } else if (!this.supportedSlideTypes.includes(slideConfig.type)) {
      errors.push(`Slide type '${slideConfig.type}' is not supported.`);
    }

    if (Array.isArray(slideConfig.components)) {
      slideConfig.components.forEach((componentConfig, index) => {
        const result = this.validateComponent(componentConfig);
        result.errors.forEach(error => errors.push(`Component ${index + 1}: ${error}`));
        result.warnings.forEach(warning => warnings.push(`Component ${index + 1}: ${warning}`));
      });
    }

    if (slideConfig.bounds) {
      const boundsResult = this.validateBounds(slideConfig.bounds, slideConfig.slideWidth || 0, slideConfig.slideHeight || 0);
      boundsResult.errors.forEach(error => errors.push(`Bounds: ${error}`));
      boundsResult.warnings.forEach(warning => warnings.push(`Bounds: ${warning}`));
    }

    return { valid: errors.length === 0, errors, warnings };
  }

  /**
   * Validate a component configuration object.
   * @param {Object} componentConfig - Component configuration.
   * @returns {Object} Validation result.
   */
  validateComponent(componentConfig) {
    const errors = [];
    const warnings = [];

    if (!componentConfig || typeof componentConfig !== 'object') {
      errors.push('Component config must be an object.');
      return { valid: false, errors, warnings };
    }

    if (!componentConfig.type || typeof componentConfig.type !== 'string') {
      errors.push('Component must include a type string.');
      return { valid: false, errors, warnings };
    }

    const type = componentConfig.type;

    if (this.supportedTextComponents.includes(type)) {
      if (!componentConfig.text || typeof componentConfig.text !== 'string') {
        errors.push(`Component type '${type}' requires a non-empty text property.`);
      }
    }

    switch (type) {
      case 'chart': {
        const chartResult = this.validateChart(componentConfig.chart || componentConfig);
        chartResult.errors.forEach(error => errors.push(error));
        chartResult.warnings.forEach(warning => warnings.push(warning));
        break;
      }
      case 'table': {
        const tableResult = this.validateTable(componentConfig.table || componentConfig);
        tableResult.errors.forEach(error => errors.push(error));
        tableResult.warnings.forEach(warning => warnings.push(warning));
        break;
      }
      case 'image': {
        const imageResult = this.validateImage(componentConfig.image || componentConfig);
        imageResult.errors.forEach(error => errors.push(error));
        imageResult.warnings.forEach(warning => warnings.push(warning));
        break;
      }
      default:
        if (!this.supportedTextComponents.includes(type)) {
          warnings.push(`Component type '${type}' is not explicitly validated.`);
        }
        break;
    }

    if (componentConfig.bounds) {
      const boundsResult = this.validateBounds(componentConfig.bounds, componentConfig.slideWidth || 0, componentConfig.slideHeight || 0);
      boundsResult.errors.forEach(error => errors.push(`Bounds: ${error}`));
      boundsResult.warnings.forEach(warning => warnings.push(`Bounds: ${warning}`));
    }

    return { valid: errors.length === 0, errors, warnings };
  }

  /**
   * Validate bounds values for a component or slide.
   * @param {Object} component - Bounds object with x, y, w, h.
   * @param {number} slideWidth - Slide width.
   * @param {number} slideHeight - Slide height.
   * @returns {Object} Validation result.
   */
  validateBounds(component, slideWidth, slideHeight) {
    const errors = [];
    const warnings = [];

    if (!component || typeof component !== 'object') {
      errors.push('Bounds must be an object with x, y, w, h.');
      return { valid: false, errors, warnings };
    }

    ['x', 'y', 'w', 'h'].forEach(key => {
      if (component[key] === undefined || typeof component[key] !== 'number' || Number.isNaN(component[key])) {
        errors.push(`Bounds.${key} must be a numeric value.`);
      }
    });

    if (typeof component.x === 'number' && component.x < 0) {
      errors.push('Bounds.x must be greater than or equal to 0.');
    }
    if (typeof component.y === 'number' && component.y < 0) {
      errors.push('Bounds.y must be greater than or equal to 0.');
    }
    if (typeof component.w === 'number' && component.w <= 0) {
      errors.push('Bounds.w must be greater than 0.');
    }
    if (typeof component.h === 'number' && component.h <= 0) {
      errors.push('Bounds.h must be greater than 0.');
    }

    if (typeof slideWidth === 'number' && typeof slideHeight === 'number') {
      if (component.x + component.w > slideWidth) {
        errors.push('Bounds x+w must not exceed slide width.');
      }
      if (component.y + component.h > slideHeight) {
        errors.push('Bounds y+h must not exceed slide height.');
      }
    }

    return { valid: errors.length === 0, errors, warnings };
  }

  /**
   * Validate chart configuration.
   * @param {Object} chartConfig - Chart configuration.
   * @returns {Object} Validation result.
   */
  validateChart(chartConfig) {
    const errors = [];
    const warnings = [];

    if (!chartConfig || typeof chartConfig !== 'object') {
      errors.push('Chart config must be an object.');
      return { valid: false, errors, warnings };
    }

    if (!chartConfig.type || typeof chartConfig.type !== 'string') {
      errors.push('Chart config must include a type string.');
    } else if (!this.supportedChartTypes.includes(chartConfig.type)) {
      errors.push(`Chart type '${chartConfig.type}' is not supported.`);
    }

    if (!Array.isArray(chartConfig.data)) {
      errors.push('Chart config must include a data array.');
    }

    return { valid: errors.length === 0, errors, warnings };
  }

  /**
   * Validate table configuration.
   * @param {Object} tableConfig - Table configuration.
   * @returns {Object} Validation result.
   */
  validateTable(tableConfig) {
    const errors = [];
    const warnings = [];

    if (!tableConfig || typeof tableConfig !== 'object') {
      errors.push('Table config must be an object.');
      return { valid: false, errors, warnings };
    }

    if (!Array.isArray(tableConfig.rows)) {
      errors.push('Table config must include a rows array.');
    }

    return { valid: errors.length === 0, errors, warnings };
  }

  /**
   * Validate image configuration.
   * @param {Object} imageConfig - Image configuration.
   * @returns {Object} Validation result.
   */
  validateImage(imageConfig) {
    const errors = [];
    const warnings = [];

    if (!imageConfig || typeof imageConfig !== 'object') {
      errors.push('Image config must be an object.');
      return { valid: false, errors, warnings };
    }

    if (!imageConfig.path || typeof imageConfig.path !== 'string' || imageConfig.path.trim() === '') {
      errors.push('Image config must include a valid path string.');
    }

    return { valid: errors.length === 0, errors, warnings };
  }

  /**
   * Validate theme configuration.
   * @param {Object} themeConfig - Theme configuration.
   * @returns {Object} Validation result.
   */
  validateTheme(themeConfig) {
    const errors = [];
    const warnings = [];

    if (!themeConfig || typeof themeConfig !== 'object') {
      errors.push('Theme config must be an object.');
      return { valid: false, errors, warnings };
    }

    const requiredColors = [
      'primaryColor',
      'secondaryColor',
      'accentColor',
      'backgroundColor',
      'textColor'
    ];
    requiredColors.forEach(key => {
      if (!themeConfig[key] || typeof themeConfig[key] !== 'string') {
        errors.push(`Theme must include a valid ${key} string.`);
      }
    });

    const requiredFonts = ['titleFont', 'bodyFont'];
    requiredFonts.forEach(key => {
      if (!themeConfig[key] || typeof themeConfig[key] !== 'string') {
        errors.push(`Theme must include a valid ${key} string.`);
      }
    });

    return { valid: errors.length === 0, errors, warnings };
  }
}

module.exports = ValidationEngine;
