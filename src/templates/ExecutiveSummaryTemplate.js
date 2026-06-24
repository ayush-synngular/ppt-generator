'use strict';

const { createTitleSlide } = require('../layouts/title-slide');
const { renderTitle } = require('../components/title');
const { renderBulletList } = require('../components/bullet-list');
const { renderMetricCard } = require('../components/metric-card');

/**
 * Executive Summary template.
 * Expected data:
 * {
 *   title,
 *   subtitle,
 *   summary,
 *   metrics: [{ title, value }],
 *   recommendations: []
 * }
 */
class ExecutiveSummaryTemplate {
  constructor(presentation, theme) {
    this.presentation = presentation;
    this.theme = theme;
  }

  build(data = {}) {
    const {
      title = 'Executive Summary',
      subtitle = '',
      summary = '',
      metrics = [],
      recommendations = []
    } = data;

    const slides = [];

    slides.push(createTitleSlide(this.presentation, { title, subtitle, theme: this.theme }));

    const summarySlide = this.presentation.addSlide();
    renderTitle(summarySlide, { text: 'Summary', x: 0.5, y: 0.5, w: 9, h: 1 }, this.theme);
    renderBulletList(summarySlide, { items: Array.isArray(summary) ? summary : [summary], x: 0.5, y: 1.5, w: 8, h: 4 }, this.theme);
    slides.push(summarySlide);

    const metricsSlide = this.presentation.addSlide();
    renderTitle(metricsSlide, { text: 'Key Metrics', x: 0.5, y: 0.5, w: 9, h: 1 }, this.theme);

    metrics.slice(0, 4).forEach((metric, index) => {
      const x = 0.5 + (index % 2) * 4.75;
      const y = 1.5 + Math.floor(index / 2) * 2.75;
      renderMetricCard(metricsSlide, { title: metric.title || '', value: metric.value || '', x, y, w: 4, h: 2.25 }, this.theme);
    });

    slides.push(metricsSlide);

    const recommendationsSlide = this.presentation.addSlide();
    renderTitle(recommendationsSlide, { text: 'Recommendations', x: 0.5, y: 0.5, w: 9, h: 1 }, this.theme);
    renderBulletList(recommendationsSlide, { items: recommendations, x: 0.5, y: 1.5, w: 8, h: 4 }, this.theme);
    slides.push(recommendationsSlide);

    return slides;
  }
}

module.exports = ExecutiveSummaryTemplate;
