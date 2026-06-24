'use strict';

const DataBinder = require('../src/core/DataBinder');

const binder = new DataBinder();

const tests = [
  {
    name: 'Simple token',
    template: { title: '{{company}}' },
    data: { company: 'ABC Corp' },
    expected: { title: 'ABC Corp' }
  },
  {
    name: 'Nested token',
    template: { title: '{{metrics.revenue}}' },
    data: { metrics: { revenue: 1500000 } },
    expected: { title: '1500000' }
  },
  {
    name: 'Multiple tokens in same string',
    template: { title: '{{company}} Revenue {{metrics.revenue}}' },
    data: { company: 'ABC Corp', metrics: { revenue: 1500000 } },
    expected: { title: 'ABC Corp Revenue 1500000' }
  },
  {
    name: 'Array access',
    template: { title: '{{risks[0]}}' },
    data: { risks: ['Operational risk', 'Market risk'] },
    expected: { title: 'Operational risk' }
  },
  {
    name: 'Missing token',
    template: { title: '{{unknownField}}' },
    data: { company: 'ABC Corp' },
    expected: { title: '{{unknownField}}' }
  },
  {
    name: 'Null value',
    template: { title: '{{company}}' },
    data: { company: null },
    expected: { title: '' }
  }
];

function isEqual(actual, expected) {
  return JSON.stringify(actual) === JSON.stringify(expected);
}

function runTest(test) {
  const result = binder.bind(test.template, test.data);
  const pass = isEqual(result, test.expected);
  console.log(`${test.name}: ${pass ? 'PASS' : 'FAIL'}`);
  if (!pass) {
    console.log('  expected:', JSON.stringify(test.expected));
    console.log('  actual:  ', JSON.stringify(result));
  }
}

for (const test of tests) {
  runTest(test);
}
