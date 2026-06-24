'use strict';

const { createTitleSlide } = require('../layouts/title-slide');
const { renderTitle } = require('../components/title');
const { renderBulletList } = require('../components/bullet-list');
const { renderChartCard } = require('../components/chart-card');

/**
 * Quarterly Business Review template.
 * Expected data:
 * {
 *   title,
 *   subtitle,
 *   kpis: [],
 *   revenueData: [],
 *   risks: [],
 *   nextQuarterPlan: []
 * }
 */
class QuarterlyBusinessReviewTemplate {
  constructor(presentation, theme) {
    this.presentation = presentation;
    this.theme = theme;
  }

  build(data = {}) {
    const {
      title = 'Quarterly Business Review',
      subtitle = '',
      kpis = [],
      revenueData = [],
      risks = [],
      nextQuarterPlan = []
    } = data;

    const slides = [];

    slides.push(createTitleSlide(this.presentation, { title, subtitle, theme: this.theme }));

    const kpiSlide = this.presentation.addSlide();
    renderTitle(kpiSlide, { text: 'KPI Overview', x: 0.5, y: 0.5, w: 9, h: 1 }, this.theme);
    renderBulletList(kpiSlide, { items: kpis, x: 0.5, y: 1.5, w: 8, h: 4 }, this.theme);
    slides.push(kpiSlide);

    const revenueSlide = this.presentation.addSlide();
    renderTitle(revenueSlide, { text: 'Revenue Trends', x: 0.5, y: 0.5, w: 9, h: 1 }, this.theme);
    renderChartCard(revenueSlide, {
      title: 'Revenue Trends',
      chartType: 'line',
      data: revenueData,
      x: 0.5,
      y: 1.25,
      w: 8,
      h: 4.5
    }, this.theme);
    slides.push(revenueSlide);

    const riskSlide = this.presentation.addSlide();
    renderTitle(riskSlide, { text: 'Risks', x: 0.5, y: 0.5, w: 9, h: 1 }, this.theme);
    renderBulletList(riskSlide, { items: risks, x: 0.5, y: 1.5, w: 8, h: 4 }, this.theme);
    slides.push(riskSlide);

    const planSlide = this.presentation.addSlide();
    renderTitle(planSlide, { text: 'Next Quarter Plan', x: 0.5, y: 0.5, w: 9, h: 1 }, this.theme);
    renderBulletList(planSlide, { items: nextQuarterPlan, x: 0.5, y: 1.5, w: 8, h: 4 }, this.theme);
    slides.push(planSlide);

    return slides;
  }
}

module.exports = QuarterlyBusinessReviewTemplate;
