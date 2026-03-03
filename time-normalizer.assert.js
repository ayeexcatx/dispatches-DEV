const fs = require('fs');
const assert = require('assert');

function extract(name, src) {
  const idx = src.indexOf(`function ${name}`);
  if (idx < 0) throw new Error(`Function not found: ${name}`);
  let start = src.indexOf('{', idx);
  let depth = 0;
  for (let i = start; i < src.length; i += 1) {
    if (src[i] === '{') depth += 1;
    if (src[i] === '}') {
      depth -= 1;
      if (depth === 0) return src.slice(idx, i + 1);
    }
  }
  throw new Error(`Function parse failed: ${name}`);
}

const webApp = fs.readFileSync('./webApp.js', 'utf8');
const dispatch = fs.readFileSync('./dispatch.js', 'utf8');

const clientFn = extract('normalizeTimeInput', webApp);
const serverFn = extract('normalizeTimeInput_', dispatch);

eval(`${clientFn}\n${serverFn}`);

assert.strictEqual(normalizeTimeInput('17:45'), '05:45 PM');
assert.strictEqual(normalizeTimeInput_('17:45'), '05:45 PM');

assert.strictEqual(normalizeTimeInput('17:45:00'), '05:45 PM');
assert.strictEqual(normalizeTimeInput_('17:45:00'), '05:45 PM');

assert.throws(() => normalizeTimeInput('17:45:30'));
assert.throws(() => normalizeTimeInput_('17:45:30'));

console.log('time-normalizer assertions passed');
