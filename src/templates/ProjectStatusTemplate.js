'use strict';

const { createTitleSlide } = require('../layouts/title-slide');
const { renderTitle } = require('../components/title');
const { renderBulletList } = require('../components/bullet-list');
const { renderMetricCard } = require('../components/metric-card');

/**
 * Project Status template.
 * Expected data:
 * {
 *   title,
 *   subtitle,
 *   statusOverview: [],
 *   milestones: [],
 *   risks: [],
 *   actionItems: []
 * }
 */
class ProjectStatusTemplate {
  constructor(presentation, theme) {
    this.presentation = presentation;
    this.theme = theme;
  }

  build(data = {}) {
    const {
      title = 'Project Status',
      subtitle = '',
      statusOverview = [],
      milestones = [],
      risks = [],
      actionItems = []
    } = data;

    const slides = [];

    slides.push(createTitleSlide(this.presentation, { title, subtitle, theme: this.theme }));

    const overviewSlide = this.presentation.addSlide();
    renderTitle(overviewSlide, { text: 'Status Overview', x: 0.5, y: 0.5, w: 9, h: 1 }, this.theme);
    renderBulletList(overviewSlide, { items: statusOverview, x: 0.5, y: 1.5, w: 8, h: 4 }, this.theme);
    slides.push(overviewSlide);

    const milestoneSlide = this.presentation.addSlide();
    renderTitle(milestoneSlide, { text: 'Milestones', x: 0.5, y: 0.5, w: 9, h: 1 }, this.theme);
    renderBulletList(milestoneSlide, { items: milestones, x: 0.5, y: 1.5, w: 8, h: 4 }, this.theme);
    slides.push(milestoneSlide);

    const riskSlide = this.presentation.addSlide();
    renderTitle(riskSlide, { text: 'Risks', x: 0.5, y: 0.5, w: 9, h: 1 }, this.theme);
    renderBulletList(riskSlide, { items: risks, x: 0.5, y: 1.5, w: 8, h: 4 }, this.theme);
    slides.push(riskSlide);

    const actionSlide = this.presentation.addSlide();
    renderTitle(actionSlide, { text: 'Action Items', x: 0.5, y: 0.5, w: 9, h: 1 }, this.theme);
    renderBulletList(actionSlide, { items: actionItems, x: 0.5, y: 1.5, w: 8, h: 4 }, this.theme);
    slides.push(actionSlide);

    return slides;
  }
}

module.exports = ProjectStatusTemplate;
