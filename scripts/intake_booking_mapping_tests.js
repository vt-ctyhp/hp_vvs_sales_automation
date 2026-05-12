const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const repo = path.resolve(__dirname, '..');

const context = {
  console,
  Logger: { log() {} },
  PropertiesService: {
    getScriptProperties() {
      return { getProperty() { return ''; }, setProperty() {}, deleteProperty() {} };
    },
  },
  Utilities: {
    formatDate(date, _tz, fmt) {
      const d = new Date(date);
      const pad = n => String(n).padStart(2, '0');
      if (fmt === 'HH:mm') return `${pad(d.getHours())}:${pad(d.getMinutes())}`;
      if (fmt === 'yyyyMMddHHmmss') {
        return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
      }
      if (fmt === 'M/d/yyyy') return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`;
      return d.toISOString();
    },
  },
  ScriptApp: { getProjectTriggers() { return []; }, deleteTrigger() {} },
  LockService: { getScriptLock() { return { tryLock() { return true; }, releaseLock() {} }; } },
};

vm.createContext(context);
vm.runInContext(fs.readFileSync(path.join(repo, 'Resolver.js'), 'utf8'), context, { filename: 'Resolver.js' });
vm.runInContext(fs.readFileSync(path.join(repo, 'Acuity_HPUSA.js'), 'utf8'), context, { filename: 'Acuity_HPUSA.js' });

function test(name, fn) {
  try {
    fn();
    console.log(`PASS ${name}`);
  } catch (err) {
    console.error(`FAIL ${name}`);
    throw err;
  }
}

function assertPlainDeepEqual(actual, expected) {
  assert.deepStrictEqual(JSON.parse(JSON.stringify(actual)), expected);
}

test('Resolver accepts exact Google Form diamond and budget labels', () => {
  const nv = {
    'Diamond Type': ['Lab Diamond'],
    'Budget Range': ['$1,000 - $5,000'],
  };
  assert.strictEqual(context.intakeDiamondAnswer_(nv), 'Lab');
  assert.strictEqual(context.intakeBudgetAnswer_(nv), '$1,000 - $5,000');
  assertPlainDeepEqual(context.parseBudget_('$1,000 - $5,000'), { min: 1000, max: 5000 });
});

test('Resolver accepts customer-question wording for lab/natural and budget', () => {
  const nv = {
    'Are you looking for lab-grown or natural diamond?': ['Natural diamond'],
    'What budget are you looking for?': ['$5k - $10k'],
  };
  assert.strictEqual(context.intakeDiamondAnswer_(nv), 'Natural');
  assert.strictEqual(context.intakeBudgetAnswer_(nv), '$5,001 - $10,000');
  assertPlainDeepEqual(context.parseBudget_('$5,001 - $10,000'), { min: 5001, max: 10000 });
});

test('Resolver handles both diamond types and open-ended budgets', () => {
  const nv = {
    'Lab or Natural': ['Lab and Natural'],
    'Budget': ['$20k+'],
  };
  assert.strictEqual(context.intakeDiamondAnswer_(nv), 'Both');
  assert.strictEqual(context.intakeBudgetAnswer_(nv), '$20,001+');
  assertPlainDeepEqual(context.parseBudget_('$20,001+'), { min: 20001, max: '' });
});

test('Resolver preserves multiple selected canonical budget ranges', () => {
  const nv = {
    'Budget Range': ['$1,000 - $5,000, $5,001 - $10,000'],
  };
  assert.strictEqual(context.intakeBudgetAnswer_(nv), '$1,000 - $5,000, $5,001 - $10,000');
  assertPlainDeepEqual(context.parseBudget_(context.intakeBudgetAnswer_(nv)), { min: 1000, max: 10000 });
});

test('Acuity fuzzy answer extraction finds changed question labels', () => {
  const formData = {
    'Are you looking for lab-grown or natural diamond?': 'Lab-grown',
    'What budget are you looking for?': '$10k - $15k',
  };
  assert.strictEqual(
    context.acuityFindAnswer_(formData, ['preferred diamond type'], [['lab', 'natural']]),
    'Lab-grown'
  );
  assert.strictEqual(
    context.acuityFindAnswer_(formData, ['preferred price range'], [['budget']]),
    '$10k - $15k'
  );
});

test('Acuity normalizes diamond and budget for form submission and Master edits', () => {
  assertPlainDeepEqual(context.acuityNormalizeDiamond_('Lab or Natural'), ['Lab Diamond', 'Natural Diamond']);
  assert.strictEqual(context.acuityDiamondMasterValue_(['Lab Diamond', 'Natural Diamond']), 'Both');
  assert.strictEqual(context.acuityNormalizeBudget_('$10k - $15k'), '$10,001 - $15,000');
  assert.strictEqual(context.acuityNormalizeBudget_('Over $20k'), '$20,001+');
});

test('Acuity field map carries two checkbox answers when customer says both', () => {
  const fieldMap = context.acuityToFormFieldMap_({
    id: 'TEST-ACUITY-MAP',
    firstName: 'Test',
    lastName: 'Customer',
    email: 'test@example.com',
    phone: '4085551212',
    type: 'In-Person Custom Design Consultation',
    datetime: '2026-05-12T10:30:00-07:00',
    forms: [{
      values: [
        { name: 'Are you looking for lab-grown or natural diamond?', value: 'Both' },
        { name: 'What budget are you looking for?', value: '$15k - $20k' },
      ],
    }],
  });
  assertPlainDeepEqual(fieldMap['Diamond Type'], ['Lab Diamond', 'Natural Diamond']);
  assertPlainDeepEqual(fieldMap['Budget Range'], ['$15,001 - $20,000']);
});

console.log('All intake booking mapping tests passed.');
