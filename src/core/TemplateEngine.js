'use strict';

/**
 * Template engine for registering and building presentation templates.
 *
 * Template contract:
 * - A template must expose a `build(data)` method.
 * - `build(data)` should create slides on the provided presentation and return an array of slides.
 *
 * Expected usage:
 * const engine = new TemplateEngine(presentation, theme);
 * engine.registerTemplate('executive-summary', new ExecutiveSummaryTemplate(presentation, theme));
 * engine.build('executive-summary', data);
 */
const DataBinder = require('./DataBinder');

class TemplateEngine {
  /**
   * @param {Object} presentation - A Presentation instance.
   * @param {Object} theme - A Theme instance.
   */
  constructor(presentation, theme) {
    if (!presentation || typeof presentation.addSlide !== 'function') {
      throw new Error('TemplateEngine requires a valid presentation instance.');
    }

    this.presentation = presentation;
    this.theme = theme;
    this.templates = {};
    this.dataBinder = new DataBinder();
  }

  /**
   * Register a template by name.
   * @param {string} name - Template identifier.
   * @param {Object} template - Template instance with a build(data) method.
   */
  registerTemplate(name, template) {
    if (!name || typeof name !== 'string') {
      throw new Error('TemplateEngine.registerTemplate requires a valid name string.');
    }

    if (!template || typeof template.build !== 'function') {
      throw new Error('TemplateEngine.registerTemplate requires a template with a build(data) method.');
    }

    this.templates[name] = template;
  }

  /**
   * Retrieve a registered template by name.
   * @param {string} name - Template identifier.
   * @returns {Object} Template instance.
   */
  getTemplate(name) {
    const template = this.templates[name];

    if (!template) {
      throw new Error(`TemplateEngine could not find template '${name}'.`);
    }

    return template;
  }

  /**
   * Build slides using a registered template.
   * @param {string} name - Template name.
   * @param {Object} data - Template input data.
   * @returns {Object[]} Created slides.
   */
  build(name, data) {
    const template = this.getTemplate(name);
    const boundData = this.dataBinder.bind(data, data);
    const slides = template.build(boundData);

    if (!Array.isArray(slides)) {
      throw new Error(`Template '${name}' must return an array of slides from build().`);
    }

    return slides;
  }
}

module.exports = TemplateEngine;
