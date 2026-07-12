/**
 * Standalone Node harness verifying the new Log.info/warn/error helper
 * (utils.js) that replaces console.* for server diagnostics. Confirms it
 * writes through Logger.log (the only reliably-visible sink for anonymous
 * web-app executions), applies a severity prefix, and -- unlike a naive
 * pass-through to Logger.log's printf-style formatting, which silently
 * drops a second raw argument when the first string has no "%s" -- joins
 * every argument the way console.* does, stringifying Error objects to
 * their stack and plain objects to JSON so nothing caught gets lost.
 *
 * Run: node .pw-test/test_log_helper.js
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');

const loggerCalls = [];
const sandbox = {
  console,
  Logger: { log: (msg) => loggerCalls.push(msg) }
};
sandbox.global = sandbox;
const ctx = vm.createContext(sandbox);

vm.runInContext(fs.readFileSync(path.join(ROOT, 'utils.js'), 'utf8'), ctx, { filename: 'utils.js' });

// `const Log` doesn't become a vm context property automatically (unlike a
// `function` declaration), so re-expose it explicitly.
vm.runInContext('global.Log = Log;', ctx, { filename: 'expose.js' });
const { Log } = ctx;

let failures = 0;
function assert(cond, msg) {
  if (!cond) { failures++; console.error('FAIL:', msg); }
  else console.log('PASS:', msg);
}

loggerCalls.length = 0;
Log.info('[myFunc] simple message');
assert(loggerCalls.length === 1, 'info() writes exactly one Logger.log line');
assert(loggerCalls[0] === '[INFO] [myFunc] simple message', `info() applies severity prefix, preserves existing [tag] (got ${JSON.stringify(loggerCalls[0])})`);

loggerCalls.length = 0;
Log.warn('[myFunc] something looked off:', 42);
assert(loggerCalls[0] === '[WARN] [myFunc] something looked off: 42', `warn() joins a non-string second argument (got ${JSON.stringify(loggerCalls[0])})`);

loggerCalls.length = 0;
const err = new Error('boom');
Log.error('[myFunc] caught error:', err);
assert(loggerCalls[0].startsWith('[ERROR] [myFunc] caught error: Error: boom'),
  `error() with an Error object includes its stack, not just "[object Object]" (got ${JSON.stringify(loggerCalls[0])})`);

loggerCalls.length = 0;
Log.error('[myFunc] bad payload:', { code: 7, reason: 'x' });
assert(loggerCalls[0] === '[ERROR] [myFunc] bad payload: {"code":7,"reason":"x"}',
  `error() JSON-stringifies a plain object argument instead of dropping it (got ${JSON.stringify(loggerCalls[0])})`);

// The exact failure mode Log.* exists to avoid: Logger.log itself would
// silently drop this second argument since the first string has no "%s".
loggerCalls.length = 0;
sandbox.Logger.log('[myFunc] a plain message with no format specifier', err);
assert(loggerCalls[0] === '[myFunc] a plain message with no format specifier',
  `sanity check: raw Logger.log really does drop a 2nd arg with no %s (got ${JSON.stringify(loggerCalls[0])})`);

console.log(failures === 0 ? '\nALL TESTS PASSED' : `\n${failures} TEST(S) FAILED`);
process.exitCode = failures === 0 ? 0 : 1;
