"""
audit_app.py — Static audit suite for the Maharaja Bikes ERP Apps Script project.

Runs a battery of structural/consistency checks across the .js and .html source
files and prints a categorized report (ERROR / WARNING / INFO). Findings are also
written to audit_report.txt. Exits with status 1 if any ERROR-level issue exists.

Checks performed:
  1. JSON config validity         (appsscript.json, .clasp.json)
  2. JS syntax                    (node --check on every top-level .js file)
  3. Inline <script> syntax       (node --check on inline scripts in .html files)
  4. Include graph                (include('X') -> X.html exists; orphaned files)
  5. Global namespace collisions  (duplicate top-level function/const names)
  6. Client/server API contract   (Api.call('fn') -> fn exists as a server function)
  7. LockService usage            (tryLock without matching finally/releaseLock)
  8. DOM id cross-reference        (getElementById/querySelector vs id="..." in DOM)
  9. Config key usage              (APP_CONFIG.*/<SHEET>_COL.* references are valid)
 10. TODO/FIXME/placeholder markers and leftover console/debugger statements
 11. appsscript.json security posture notes
"""

import os
import re
import json
import subprocess
import sys
import tempfile
from collections import Counter

ROOT = os.path.dirname(os.path.abspath(__file__))


# ─────────────────────────────────────────────────────────────────────────
# Reporting
# ─────────────────────────────────────────────────────────────────────────

class Report:
    def __init__(self):
        self.items = []  # (severity, category, message)

    def add(self, severity, category, message):
        self.items.append((severity, category, message))

    def error(self, category, message):
        self.add('ERROR', category, message)

    def warn(self, category, message):
        self.add('WARNING', category, message)

    def info(self, category, message):
        self.add('INFO', category, message)

    def by_severity(self, severity):
        return [i for i in self.items if i[0] == severity]

    def print_and_save(self):
        order = ['ERROR', 'WARNING', 'INFO']
        lines = []
        lines.append('=' * 78)
        lines.append('AUDIT REPORT — Maharaja Bikes ERP (Apps Script project)')
        lines.append('=' * 78)
        for severity in order:
            items = self.by_severity(severity)
            lines.append('')
            lines.append(f'-- {severity} ({len(items)}) ' + '-' * (60 - len(severity)))
            if not items:
                lines.append('  (none)')
            for _, category, message in items:
                lines.append(f'  [{category}] {message}')

        lines.append('')
        lines.append('=' * 78)
        lines.append(
            f"SUMMARY: {len(self.by_severity('ERROR'))} error(s), "
            f"{len(self.by_severity('WARNING'))} warning(s), "
            f"{len(self.by_severity('INFO'))} info note(s)"
        )
        lines.append('=' * 78)

        text = '\n'.join(lines)
        print(text)

        with open(os.path.join(ROOT, 'audit_report.txt'), 'w', encoding='utf-8') as f:
            f.write(text + '\n')


# ─────────────────────────────────────────────────────────────────────────
# File discovery helpers
# ─────────────────────────────────────────────────────────────────────────

def list_root_files(ext):
    return sorted(
        f for f in os.listdir(ROOT)
        if f.lower().endswith(ext) and os.path.isfile(os.path.join(ROOT, f))
    )


def read_file(name):
    with open(os.path.join(ROOT, name), encoding='utf-8') as f:
        return f.read()


# ─────────────────────────────────────────────────────────────────────────
# 1. JSON config validity
# ─────────────────────────────────────────────────────────────────────────

def check_json_configs(report):
    for name in ('appsscript.json', '.clasp.json'):
        path = os.path.join(ROOT, name)
        if not os.path.exists(path):
            report.warn('json-config', f'{name} not found')
            continue
        try:
            with open(path, encoding='utf-8') as f:
                json.load(f)
        except json.JSONDecodeError as e:
            report.error('json-config', f'{name} is not valid JSON: {e}')


# ─────────────────────────────────────────────────────────────────────────
# 2. JS syntax check
# ─────────────────────────────────────────────────────────────────────────

def check_js_syntax(report, js_files):
    for name in js_files:
        path = os.path.join(ROOT, name)
        result = subprocess.run(['node', '--check', path], capture_output=True, text=True)
        if result.returncode != 0:
            report.error('js-syntax', f'{name}: {result.stderr.strip()}')


# ─────────────────────────────────────────────────────────────────────────
# 3. Inline <script> syntax check (HTML files)
# ─────────────────────────────────────────────────────────────────────────

SCRIPT_TAG_RE = re.compile(r'<script\b([^>]*)>(.*?)</script>', re.IGNORECASE | re.DOTALL)


def check_inline_html_scripts(report, html_files):
    for name in html_files:
        content = read_file(name)
        if '<?' in content:
            # Contains Apps Script scriptlets; only safe to check if no <script>
            # block itself contains a scriptlet (verified separately).
            pass

        for idx, m in enumerate(SCRIPT_TAG_RE.finditer(content), start=1):
            attrs, body = m.group(1), m.group(2)
            if 'src=' in attrs.lower():
                continue  # external script, nothing to check
            if '<?' in body:
                report.warn(
                    'html-script-syntax',
                    f'{name}: <script> block {idx} contains Apps Script scriptlets '
                    f'(<? ... ?>) — skipped (cannot syntax-check as plain JS)'
                )
                continue
            if not body.strip():
                continue

            with tempfile.NamedTemporaryFile(
                mode='w', suffix='.js', delete=False, encoding='utf-8'
            ) as tmp:
                tmp.write(body)
                tmp_path = tmp.name
            try:
                result = subprocess.run(
                    ['node', '--check', tmp_path], capture_output=True, text=True
                )
                if result.returncode != 0:
                    report.error(
                        'html-script-syntax',
                        f'{name}: <script> block {idx} failed syntax check: '
                        f'{result.stderr.strip()}'
                    )
            finally:
                os.remove(tmp_path)


# ─────────────────────────────────────────────────────────────────────────
# 4. Include graph (include('X') -> X.html exists; orphaned files)
# ─────────────────────────────────────────────────────────────────────────

INCLUDE_RE = re.compile(r"include\(\s*['\"]([^'\"]+)['\"]\s*\)")


def check_includes(report, html_files):
    all_html = set(html_files)

    # Resolve the include graph starting at Index.html
    reachable = set()

    def visit(name):
        filename = f'{name}.html'
        if filename in reachable:
            return
        if filename not in all_html:
            report.error('include-graph', f"'{name}' is included but {filename} does not exist")
            return
        reachable.add(filename)
        content = read_file(filename)
        for ref in INCLUDE_RE.findall(content):
            visit(ref)

    if 'Index.html' not in all_html:
        report.error('include-graph', 'Index.html not found — cannot resolve include graph')
        return reachable

    reachable.add('Index.html')
    content = read_file('Index.html')
    for ref in INCLUDE_RE.findall(content):
        visit(ref)

    orphans = sorted(all_html - reachable)
    for name in orphans:
        report.warn(
            'include-graph',
            f'{name} is not reachable from Index.html via include() — '
            f'dead/unused file or not wired into the build'
        )

    return reachable


# ─────────────────────────────────────────────────────────────────────────
# 5. Global namespace collisions (Apps Script merges all .js files into one scope)
# ─────────────────────────────────────────────────────────────────────────

TOP_FUNC_RE = re.compile(r'^function\s+([A-Za-z_$][\w$]*)\s*\(', re.MULTILINE)
TOP_DECL_RE = re.compile(r'^(?:const|let|var)\s+([A-Za-z_$][\w$]*)\b', re.MULTILINE)


def check_global_namespace(report, js_files):
    functions = {}
    declarations = {}

    for name in js_files:
        content = read_file(name)
        for m in TOP_FUNC_RE.finditer(content):
            functions.setdefault(m.group(1), []).append(name)
        for m in TOP_DECL_RE.finditer(content):
            declarations.setdefault(m.group(1), []).append(name)

    for fn, files in functions.items():
        if len(files) > 1:
            report.error(
                'global-namespace',
                f"function '{fn}' is declared in multiple files {files} — "
                f"Apps Script shares one global scope, so one definition silently "
                f"shadows the other"
            )

    for decl, files in declarations.items():
        if len(files) > 1:
            report.error(
                'global-namespace',
                f"top-level '{decl}' is declared in multiple files {files} — "
                f"duplicate const/let/var at global scope causes a redeclaration error"
            )

    # A function and a const/let/var sharing the same name is also a collision
    shared = set(functions) & set(declarations)
    for name in shared:
        report.error(
            'global-namespace',
            f"'{name}' is declared as both a function ({functions[name]}) and a "
            f"const/let/var ({declarations[name]})"
        )

    return functions


# ─────────────────────────────────────────────────────────────────────────
# 6. Client/server API contract (Api.call('fn', ...) -> server function exists)
# ─────────────────────────────────────────────────────────────────────────

API_CALL_RE = re.compile(r"Api\.call\(\s*['\"]([A-Za-z_$][\w$]*)['\"]")


def check_api_endpoints(report, html_files, server_functions):
    used = {}
    for name in html_files:
        content = read_file(name)
        for m in API_CALL_RE.finditer(content):
            used.setdefault(m.group(1), set()).add(name)

    for fn, files in used.items():
        if fn not in server_functions:
            report.error(
                'api-contract',
                f"Api.call('{fn}') referenced in {sorted(files)} but no top-level "
                f"function '{fn}' exists in any .js file"
            )
        elif fn.startswith('_'):
            report.warn(
                'api-contract',
                f"Api.call('{fn}') in {sorted(files)} invokes '{fn}', which follows "
                f"the project's '_private helper' naming convention"
            )


# ─────────────────────────────────────────────────────────────────────────
# 7. LockService usage (tryLock without matching finally + releaseLock)
# ─────────────────────────────────────────────────────────────────────────

FUNC_START_RE = re.compile(r'^function\s+([A-Za-z_$][\w$]*)', re.MULTILINE)


def check_lock_usage(report, js_files):
    for name in js_files:
        content = read_file(name)
        starts = [(m.start(), m.group(1)) for m in FUNC_START_RE.finditer(content)]
        for i, (pos, fn_name) in enumerate(starts):
            end = starts[i + 1][0] if i + 1 < len(starts) else len(content)
            chunk = content[pos:end]
            if '.tryLock(' not in chunk:
                continue
            if '.releaseLock()' not in chunk:
                report.error(
                    'lock-usage',
                    f"{name}: function '{fn_name}' calls tryLock() but never "
                    f"releaseLock() — lock will remain held until timeout"
                )
            elif 'finally' not in chunk:
                report.warn(
                    'lock-usage',
                    f"{name}: function '{fn_name}' calls tryLock()/releaseLock() "
                    f"but releaseLock() is not inside a finally block — an early "
                    f"exception path could leak the lock"
                )


# ─────────────────────────────────────────────────────────────────────────
# 8. DOM id cross-reference
# ─────────────────────────────────────────────────────────────────────────

ID_ATTR_RE = re.compile(r'\bid\s*=\s*["\']([A-Za-z][\w-]*)["\']')
GET_ELEMENT_BY_ID_RE = re.compile(r"getElementById\(\s*['\"]([A-Za-z][\w-]*)['\"]\s*\)")
QUERY_SELECTOR_RE = re.compile(r"querySelector(?:All)?\(\s*['\"]#([A-Za-z][\w-]*)['\"]")
JQUERY_ID_RE = re.compile(r"\$\(\s*['\"]#([A-Za-z][\w-]*)['\"]\s*\)")


def check_dom_ids(report, active_html_files):
    occurrences = []  # list of (id, file)
    for name in active_html_files:
        content = read_file(name)
        for m in ID_ATTR_RE.finditer(content):
            occurrences.append((m.group(1), name))

    counts = Counter(id_ for id_, _ in occurrences)
    files_by_id = {}
    for id_, name in occurrences:
        files_by_id.setdefault(id_, []).append(name)

    # Duplicate ids across the merged DOM (Index.html stitches all of these together)
    for id_, count in counts.items():
        if count > 1:
            report.warn(
                'dom-ids',
                f"id=\"{id_}\" appears {count} times across {files_by_id[id_]} — "
                f"duplicate DOM ids make getElementById() ambiguous"
            )

    ids_defined = set(counts)
    referenced = {}  # id -> set(files)
    for name in active_html_files:
        content = read_file(name)
        for rgx in (GET_ELEMENT_BY_ID_RE, QUERY_SELECTOR_RE, JQUERY_ID_RE):
            for m in rgx.finditer(content):
                referenced.setdefault(m.group(1), set()).add(name)

    for id_, files in referenced.items():
        if id_ not in ids_defined:
            report.warn(
                'dom-ids',
                f"'#{id_}' is referenced in {sorted(files)} but no element with "
                f"id=\"{id_}\" is defined in the active HTML files"
            )


# ─────────────────────────────────────────────────────────────────────────
# 9. Config key usage (APP_CONFIG.*, <SHEET>_COL.*, etc.)
# ─────────────────────────────────────────────────────────────────────────

FROZEN_OBJECT_RE = re.compile(r'^const\s+([A-Za-z_$][\w$]*)\s*=\s*Object\.freeze\(\{', re.MULTILINE)
KEY_RE = re.compile(r"^\s*'?([A-Za-z_$][\w$]*)'?\s*:")
NESTED_FREEZE_RE = re.compile(r'^\s*Object\.freeze\(\{')


def _extract_braced_body(content, start):
    """start points just after an opening '{'. Returns (body, index_after_closing_brace)."""
    depth = 1
    i = start
    n = len(content)
    in_str = None
    while i < n:
        c = content[i]
        if in_str:
            if c == '\\':
                i += 2
                continue
            if c == in_str:
                in_str = None
            i += 1
            continue
        if c in ('"', "'", '`'):
            in_str = c
            i += 1
            continue
        if c == '{':
            depth += 1
        elif c == '}':
            depth -= 1
            if depth == 0:
                return content[start:i], i + 1
        i += 1
    return content[start:], n


def _parse_object_keys(body, prefix, objects):
    keys = set()
    i = 0
    n = len(body)
    depth = 0
    in_str = None
    while i < n:
        c = body[i]
        if in_str:
            if c == '\\':
                i += 2
                continue
            if c == in_str:
                in_str = None
            i += 1
            continue
        if c in ('"', "'", '`'):
            in_str = c
            i += 1
            continue
        if c == '{':
            depth += 1
            i += 1
            continue
        if c == '}':
            depth -= 1
            i += 1
            continue
        if depth == 0:
            m = KEY_RE.match(body[i:])
            if m:
                key = m.group(1)
                keys.add(key)
                rest = body[i + m.end():]
                m2 = NESTED_FREEZE_RE.match(rest)
                if m2:
                    sub_start = i + m.end() + m2.end()
                    sub_body, _ = _extract_braced_body(body, sub_start)
                    objects[f'{prefix}.{key}'] = _parse_object_keys(sub_body, f'{prefix}.{key}', objects)
                i += m.end()
                continue
        i += 1
    return keys


def parse_frozen_objects(content):
    # Strip single-line and multi-line comments first to prevent comments
    # (specifically apostrophes/braces in comments) from corrupting the parser state.
    content = re.sub(r'/\*.*?\*/', '', content, flags=re.DOTALL)
    content = re.sub(r'//.*$', '', content, flags=re.MULTILINE)
    
    objects = {}
    for m in FROZEN_OBJECT_RE.finditer(content):
        name = m.group(1)
        body, _ = _extract_braced_body(content, m.end())
        objects[name] = _parse_object_keys(body, name, objects)
    return objects


KEY_ACCESS_RE = re.compile(r'\b([A-Z][A-Z0-9_]*(?:\.[A-Z][A-Z0-9_]*)*)\.([A-Za-z_$][\w$]*)')


def check_config_keys(report, js_files):
    if 'config.js' not in js_files:
        report.warn('config-keys', 'config.js not found — skipping config key audit')
        return

    objects = parse_frozen_objects(read_file('config.js'))
    if not objects:
        report.warn('config-keys', 'No Object.freeze() configuration blocks found in config.js')
        return

    known_prefixes = set(objects)

    for name in js_files:
        content = read_file(name)
        for m in KEY_ACCESS_RE.finditer(content):
            prefix, key = m.group(1), m.group(2)
            if prefix not in known_prefixes:
                continue
            valid_keys = objects[prefix]
            if key in ('hasOwnProperty', 'freeze', 'keys', 'values', 'entries'):
                continue
            if key not in valid_keys:
                report.error(
                    'config-keys',
                    f"{name}: '{prefix}.{key}' is not a defined key in config.js "
                    f"(valid keys: {sorted(valid_keys)})"
                )


# ─────────────────────────────────────────────────────────────────────────
# 10. TODO/FIXME/placeholder markers and leftover console/debugger statements
# ─────────────────────────────────────────────────────────────────────────

MARKER_RE = re.compile(r'\b(TODO|FIXME|HACK|XXX|BUG)\b\s*:', re.IGNORECASE)
PLACEHOLDER_RE = re.compile(r'INSERT[ _][A-Z ]*HERE|YOUR[_ ][A-Z_ ]*HERE|REPLACE[_ ]ME', re.IGNORECASE)
CONSOLE_RE = re.compile(r'\bconsole\.(log|debug|warn|error)\(')
DEBUGGER_RE = re.compile(r'^\s*debugger\s*;', re.MULTILINE)


def check_markers(report, all_files):
    for name in all_files:
        content = read_file(name)
        for i, line in enumerate(content.splitlines(), start=1):
            if MARKER_RE.search(line):
                report.info('markers', f'{name}:{i}: {line.strip()}')
            if PLACEHOLDER_RE.search(line):
                report.warn('markers', f'{name}:{i}: unresolved placeholder — {line.strip()}')
        if DEBUGGER_RE.search(content):
            report.warn('markers', f'{name}: contains a debugger; statement')


def check_console_usage(report, js_files, html_files):
    for name in js_files:
        content = read_file(name)
        count = len(CONSOLE_RE.findall(content))
        if count:
            report.info('console', f'{name}: {count} console.* call(s)')


# ─────────────────────────────────────────────────────────────────────────
# 11. appsscript.json security posture notes
# ─────────────────────────────────────────────────────────────────────────

def check_security_posture(report):
    path = os.path.join(ROOT, 'appsscript.json')
    if not os.path.exists(path):
        return
    with open(path, encoding='utf-8') as f:
        config = json.load(f)

    webapp = config.get('webapp', {})
    access = webapp.get('access')
    execute_as = webapp.get('executeAs')
    if access == 'ANYONE' and execute_as == 'USER_DEPLOYING':
        report.info(
            'security',
            "appsscript.json webapp.access is 'ANYONE' with executeAs "
            "'USER_DEPLOYING' — anyone with the URL can read/write the "
            "spreadsheet using the deploying user's permissions. Confirm this "
            "is intentional for this ERP system."
        )


# ─────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────

def main():
    # Source files may contain non-ASCII characters (₹, ─, etc.) that the
    # default Windows console encoding can't render — fall back gracefully.
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

    report = Report()

    js_files = list_root_files('.js')
    html_files = list_root_files('.html')
    all_files = js_files + html_files

    print(f'Discovered {len(js_files)} .js file(s) and {len(html_files)} .html file(s)\n')

    print('[1/8] Validating JSON configs...')
    check_json_configs(report)

    print('[2/8] Checking JS syntax (node --check)...')
    check_js_syntax(report, js_files)

    print('[3/8] Checking inline <script> syntax in HTML...')
    check_inline_html_scripts(report, html_files)

    print('[4/8] Resolving include() graph...')
    active_html = check_includes(report, html_files)

    print('[5/8] Checking global namespace for collisions...')
    server_functions = check_global_namespace(report, js_files)

    print('[6/8] Cross-checking Api.call() endpoints against server functions...')
    check_api_endpoints(report, html_files, server_functions)

    print('[7/8] Checking LockService acquire/release pairing...')
    check_lock_usage(report, js_files)

    print('[8/8] Cross-checking DOM ids and config keys...')
    check_dom_ids(report, sorted(active_html))
    check_config_keys(report, js_files)

    check_markers(report, all_files)
    check_console_usage(report, js_files, html_files)
    check_security_posture(report)

    print()
    report.print_and_save()

    if report.by_severity('ERROR'):
        sys.exit(1)


if __name__ == '__main__':
    main()
