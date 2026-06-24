'use strict';

/**
 * DataBinder supports token replacement in template configuration objects.
 *
 * Supported token syntax:
 * - {{company}}
 * - {{revenue}}
 * - {{metrics.totalRevenue}}
 */
class DataBinder {
  /**
   * Get a nested value from data by path.
   * @param {string} path - Dot-separated path.
   * @param {Object} data - Data object.
   * @returns {*} Resolved value or undefined.
   */
  get(path, data) {
    if (!path || typeof path !== 'string' || !data || typeof data !== 'object') {
      return undefined;
    }

    const parts = path.match(/[^.[\]]+/g);
    if (!parts) {
      return undefined;
    }

    let value = data;

    for (const part of parts) {
      if (value == null) {
        return undefined;
      }

      if (Array.isArray(value) && /^\\d+$/.test(part)) {
        value = value[Number(part)];
      } else {
        value = value[part];
      }
    }

    return value;
  }

  /**
   * Replace tokens inside a string value using data.
   * @param {string} value - Input string.
   * @param {Object} data - Data object.
   * @returns {string} Bound string.
   */
  replaceTokens(value, data) {
    if (typeof value !== 'string') {
      return value;
    }

    return value.replace(/\{\{\s*([^}]+?)\s*\}\}/g, (match, path) => {
      const resolved = this.get(path, data);

      if (resolved === undefined) {
        return match;
      }

      if (resolved === null) {
        return '';
      }

      return typeof resolved === 'object' ? JSON.stringify(resolved) : String(resolved);
    });
  }

  /**
   * Bind tokens recursively in a template configuration object.
   * @param {*} templateConfig - Config object, array, or string containing tokens.
   * @param {Object} data - Data object for token resolution.
   * @returns {*} Bound configuration structure.
   */
  bind(templateConfig, data) {
    if (Array.isArray(templateConfig)) {
      return templateConfig.map(item => this.bind(item, data));
    }

    if (templateConfig && typeof templateConfig === 'object') {
      const result = {};
      for (const key of Object.keys(templateConfig)) {
        result[key] = this.bind(templateConfig[key], data);
      }
      return result;
    }

    if (typeof templateConfig === 'string') {
      return this.replaceTokens(templateConfig, data);
    }

    return templateConfig;
  }
}

module.exports = DataBinder;
