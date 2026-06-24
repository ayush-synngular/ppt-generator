'use strict';

const {
  Presentation,
  Theme,
  TemplateEngine
} = require('../src/index');
const ExecutiveSummaryTemplate = require('../src/templates/ExecutiveSummaryTemplate');
const QuarterlyBusinessReviewTemplate = require('../src/templates/QuarterlyBusinessReviewTemplate');
const ProjectStatusTemplate = require('../src/templates/ProjectStatusTemplate');

async function testTemplates() {
  try {
    const theme = new Theme({
      primaryColor: '#0B5394',
      secondaryColor: '#D9E1F2',
      accentColor: '#F1C232',
      backgroundColor: '#FFFFFF',
      textColor: '#1F2937',
      titleFont: 'Calibri Light',
      bodyFont: 'Calibri',
      titleFontSize: 32,
      bodyFontSize: 18,
      slideMargin: 0.75
    });

    const ppt = new Presentation('Template Engine Demo');
    const engine = new TemplateEngine(ppt, theme);

    engine.registerTemplate('executive-summary', new ExecutiveSummaryTemplate(ppt, theme));
    engine.registerTemplate('quarterly-business-review', new QuarterlyBusinessReviewTemplate(ppt, theme));
    engine.registerTemplate('project-status', new ProjectStatusTemplate(ppt, theme));

    engine.build('executive-summary', {
      title: 'Executive Summary',
      subtitle: 'Q2 Review',
      summary: ['The business is healthy.', 'Sales increased by 15%'],
      metrics: [
        { title: 'Revenue', value: '$2.4M' },
        { title: 'Profit', value: '$540K' },
        { title: 'Growth', value: '15%' },
        { title: 'Retention', value: '92%' }
      ],
      recommendations: ['Invest in marketing', 'Improve customer support']
    });

    engine.build('quarterly-business-review', {
      title: 'Q2 Business Review',
      subtitle: 'Performance and Outlook',
      kpis: ['Sales up 15%', 'Net margin 22%', 'Customer satisfaction 88%'],
      revenueData: [
        { category: 'Jan', value: 100 },
        { category: 'Feb', value: 120 },
        { category: 'Mar', value: 140 },
        { category: 'Apr', value: 155 }
      ],
      risks: ['Market volatility', 'Supply chain delays'],
      nextQuarterPlan: ['Launch new product', 'Expand partnerships']
    });

    engine.build('project-status', {
      title: 'Project Status',
      subtitle: 'Apollo Release',
      statusOverview: ['On track', 'No critical blockers'],
      milestones: ['Design complete', 'Development 80% complete', 'User testing starts'],
      risks: ['Resource shortages', 'Schedule slippage'],
      actionItems: ['Finalize acceptance criteria', 'Prepare release notes']
    });

    await ppt.save('test-templates.pptx');
    console.log('Presentation saved successfully as test-templates.pptx');
  } catch (error) {
    console.error('Error creating templates presentation:', error);
  }
}

testTemplates();
