/**
 * Standalone Node harness (same pattern as test_unit_conversion.js) that loads
 * the REAL utils.js and exercises sanitizeString()'s formula-injection guard.
 *
 * Regression target: sanitizeString() used to prepend an apostrophe to ANY
 * string starting with =, @, +, or -, which setValue() stores as a literal
 * character -- corrupting phone numbers like "+91 98765 43210". The fix
 * exempts '+' when what follows is phone/number-like (digits, spaces,
 * dashes, parens, dots). '=', '@', '-' remain unconditionally escaped
 * (negative numbers being quoted is intentional, unchanged behavior).
 *
 * Run: node .pw-test/test_sanitize_string.js
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');

const sandbox = { console };
sandbox.global = sandbox;
const ctx = vm.createContext(sandbox);

vm.runInContext(fs.readFileSync(path.join(ROOT, 'utils.js'), 'utf8'), ctx, { filename: 'utils.js' });

const { sanitizeString } = ctx;

let failures = 0;
function assert(cond, msg) {
  if (!cond) { failures++; console.error('FAIL:', msg); }
  else console.log('PASS:', msg);
}

// Phone numbers must survive untouched (the corruption this fix targets).
assert(sanitizeString('+91 98765 43210') === '+91 98765 43210',
  `phone with spaces untouched (got ${JSON.stringify(sanitizeString('+91 98765 43210'))})`);
assert(sanitizeString('+91 98765') === '+91 98765',
  `phone untouched (got ${JSON.stringify(sanitizeString('+91 98765'))})`);
assert(sanitizeString('+1 (555) 123-4567') === '+1 (555) 123-4567',
  `phone with parens/dashes untouched (got ${JSON.stringify(sanitizeString('+1 (555) 123-4567'))})`);

// Negative numbers: escaping is intentional and must NOT change.
assert(sanitizeString('-42') === "'-42",
  `negative number still escaped (got ${JSON.stringify(sanitizeString('-42'))})`);

// Real formula-injection payloads must still be escaped.
assert(sanitizeString('=SUM(A1)') === "'=SUM(A1)",
  `'=' formula escaped (got ${JSON.stringify(sanitizeString('=SUM(A1)'))})`);
assert(sanitizeString('@cmd') === "'@cmd",
  `'@' formula escaped (got ${JSON.stringify(sanitizeString('@cmd'))})`);
assert(sanitizeString('+A1+B2') === "'+A1+B2",
  `'+' followed by non-phone content still escaped (got ${JSON.stringify(sanitizeString('+A1+B2'))})`);

// Already-safe / plain values pass through unchanged.
assert(sanitizeString("'already-quoted") === "'already-quoted",
  `pre-quoted value left alone (got ${JSON.stringify(sanitizeString("'already-quoted"))})`);
assert(sanitizeString('normal text') === 'normal text',
  `plain text untouched (got ${JSON.stringify(sanitizeString('normal text'))})`);
assert(sanitizeString(null) === '', 'null -> empty string');
assert(sanitizeString(undefined) === '', 'undefined -> empty string');
assert(sanitizeString('  padded  ') === 'padded', 'whitespace trimmed');

// Leading tab/CR/LF payloads are implicitly covered: trim() strips them first,
// exposing the real leading char to the check.
assert(sanitizeString('\t=SUM(A1)') === "'=SUM(A1)",
  `tab-prefixed formula still escaped after trim (got ${JSON.stringify(sanitizeString('\t=SUM(A1)'))})`);

console.log(failures === 0 ? '\nALL TESTS PASSED' : `\n${failures} TEST(S) FAILED`);
process.exitCode = failures === 0 ? 0 : 1;
