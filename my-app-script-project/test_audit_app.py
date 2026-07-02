"""
Unit tests for audit_app.py — the static audit suite for the ERP Apps Script project.

Each check function operates on a list of filenames plus a Report; tests redirect
audit_app.ROOT to a temp directory so checks run against fixture files instead of
the real project, then assert on the resulting Report items.
"""

import json
import os

import pytest

import audit_app
from audit_app import Report


@pytest.fixture
def project(tmp_path, monkeypatch):
    """Point audit_app.ROOT at an empty temp dir; return a writer helper."""
    monkeypatch.setattr(audit_app, 'ROOT', str(tmp_path))

    def write(name, content):
        path = tmp_path / name
        path.write_text(content, encoding='utf-8')
        return path

    write.root = tmp_path
    return write


@pytest.fixture
def report():
    return Report()


# ─────────────────────────────────────────────────────────────────────────
# Report
# ─────────────────────────────────────────────────────────────────────────

class TestReport:
    def test_add_methods_tag_correct_severity(self):
        r = Report()
        r.error('cat', 'bad')
        r.warn('cat', 'meh')
        r.info('cat', 'fyi')
        assert r.by_severity('ERROR') == [('ERROR', 'cat', 'bad')]
        assert r.by_severity('WARNING') == [('WARNING', 'cat', 'meh')]
        assert r.by_severity('INFO') == [('INFO', 'cat', 'fyi')]

    def test_by_severity_empty_when_none_added(self):
        assert Report().by_severity('ERROR') == []

    def test_print_and_save_writes_report_file(self, project, capsys):
        r = Report()
        r.error('x', 'something broke')
        r.print_and_save()

        out = capsys.readouterr().out
        assert '1 error(s)' in out
        assert '[x] something broke' in out

        saved = (project.root / 'audit_report.txt').read_text(encoding='utf-8')
        assert '[x] something broke' in saved
        assert 'SUMMARY: 1 error(s), 0 warning(s), 0 info note(s)' in saved

    def test_print_and_save_none_placeholder_for_empty_category(self, project, capsys):
        Report().print_and_save()
        out = capsys.readouterr().out
        assert out.count('(none)') == 3


# ─────────────────────────────────────────────────────────────────────────
# File discovery helpers
# ─────────────────────────────────────────────────────────────────────────

class TestFileDiscovery:
    def test_list_root_files_filters_by_extension_and_case(self, project):
        project('a.js', '')
        project('B.JS', '')
        project('c.html', '')
        result = audit_app.list_root_files('.js')
        assert result == ['B.JS', 'a.js']

    def test_list_root_files_ignores_directories(self, project):
        (project.root / 'subdir.js').mkdir()
        project('real.js', '')
        assert audit_app.list_root_files('.js') == ['real.js']

    def test_read_file_returns_contents(self, project):
        project('foo.txt', 'hello world')
        assert audit_app.read_file('foo.txt') == 'hello world'


# ─────────────────────────────────────────────────────────────────────────
# 1. JSON config validity
# ─────────────────────────────────────────────────────────────────────────

class TestJsonConfigs:
    def test_missing_files_produce_warnings(self, project, report):
        audit_app.check_json_configs(report)
        assert len(report.by_severity('WARNING')) == 2

    def test_valid_json_produces_no_findings(self, project, report):
        project('appsscript.json', json.dumps({'timeZone': 'Etc/UTC'}))
        project('.clasp.json', json.dumps({'scriptId': 'abc'}))
        audit_app.check_json_configs(report)
        assert report.items == []

    def test_invalid_json_is_an_error(self, project, report):
        project('appsscript.json', '{not valid json')
        project('.clasp.json', '{}')
        audit_app.check_json_configs(report)
        errors = report.by_severity('ERROR')
        assert len(errors) == 1
        assert 'appsscript.json is not valid JSON' in errors[0][2]


# ─────────────────────────────────────────────────────────────────────────
# 2. JS syntax check
# ─────────────────────────────────────────────────────────────────────────

class TestJsSyntax:
    def test_valid_js_has_no_errors(self, project, report):
        project('ok.js', 'function f() { return 1; }')
        audit_app.check_js_syntax(report, ['ok.js'])
        assert report.items == []

    def test_invalid_js_is_reported_as_error(self, project, report):
        project('bad.js', 'function f( { return 1; }')
        audit_app.check_js_syntax(report, ['bad.js'])
        errors = report.by_severity('ERROR')
        assert len(errors) == 1
        assert 'bad.js' in errors[0][2]


# ─────────────────────────────────────────────────────────────────────────
# 3. Inline <script> syntax check
# ─────────────────────────────────────────────────────────────────────────

class TestInlineHtmlScripts:
    def test_valid_inline_script_has_no_findings(self, project, report):
        project('page.html', '<html><script>var x = 1;</script></html>')
        audit_app.check_inline_html_scripts(report, ['page.html'])
        assert report.items == []

    def test_invalid_inline_script_is_an_error(self, project, report):
        project('page.html', '<html><script>function( {</script></html>')
        audit_app.check_inline_html_scripts(report, ['page.html'])
        errors = report.by_severity('ERROR')
        assert len(errors) == 1
        assert 'page.html' in errors[0][2]

    def test_script_with_src_attribute_is_skipped(self, project, report):
        project('page.html', '<script src="lib.js">garbage((</script>')
        audit_app.check_inline_html_scripts(report, ['page.html'])
        assert report.items == []

    def test_scriptlet_blocks_are_skipped_with_warning(self, project, report):
        project('page.html', '<script><?= foo() ?> var x = 1;</script>')
        audit_app.check_inline_html_scripts(report, ['page.html'])
        warnings = report.by_severity('WARNING')
        assert len(warnings) == 1
        assert 'scriptlets' in warnings[0][2]

    def test_empty_script_block_is_skipped(self, project, report):
        project('page.html', '<script>   </script>')
        audit_app.check_inline_html_scripts(report, ['page.html'])
        assert report.items == []


# ─────────────────────────────────────────────────────────────────────────
# 4. Include graph
# ─────────────────────────────────────────────────────────────────────────

class TestIncludeGraph:
    def test_missing_index_html_is_an_error(self, project, report):
        audit_app.check_includes(report, [])
        errors = report.by_severity('ERROR')
        assert len(errors) == 1
        assert 'Index.html not found' in errors[0][2]

    def test_simple_chain_is_fully_reachable(self, project, report):
        project('Index.html', "<?!= include('Body') ?>")
        project('Body.html', '<div>hi</div>')
        reachable = audit_app.check_includes(report, ['Index.html', 'Body.html'])
        assert reachable == {'Index.html', 'Body.html'}
        assert report.items == []

    def test_include_of_missing_file_is_an_error(self, project, report):
        project('Index.html', "<?!= include('Missing') ?>")
        audit_app.check_includes(report, ['Index.html'])
        errors = report.by_severity('ERROR')
        assert len(errors) == 1
        assert "Missing.html does not exist" in errors[0][2]

    def test_orphan_file_produces_warning(self, project, report):
        project('Index.html', '<div>no includes</div>')
        project('Orphan.html', '<div>unused</div>')
        audit_app.check_includes(report, ['Index.html', 'Orphan.html'])
        warnings = report.by_severity('WARNING')
        assert len(warnings) == 1
        assert 'Orphan.html' in warnings[0][2]

    def test_circular_includes_do_not_infinite_loop(self, project, report):
        project('Index.html', "<?!= include('A') ?>")
        project('A.html', "<?!= include('Index') ?>")
        reachable = audit_app.check_includes(report, ['Index.html', 'A.html'])
        assert reachable == {'Index.html', 'A.html'}


# ─────────────────────────────────────────────────────────────────────────
# 5. Global namespace collisions
# ─────────────────────────────────────────────────────────────────────────

class TestGlobalNamespace:
    def test_unique_functions_produce_no_findings(self, project, report):
        project('a.js', 'function foo() {}')
        project('b.js', 'function bar() {}')
        functions = audit_app.check_global_namespace(report, ['a.js', 'b.js'])
        assert report.items == []
        assert set(functions) == {'foo', 'bar'}

    def test_duplicate_function_across_files_is_an_error(self, project, report):
        project('a.js', 'function foo() {}')
        project('b.js', 'function foo() {}')
        audit_app.check_global_namespace(report, ['a.js', 'b.js'])
        errors = report.by_severity('ERROR')
        assert len(errors) == 1
        assert "function 'foo' is declared in multiple files" in errors[0][2]

    def test_duplicate_top_level_const_is_an_error(self, project, report):
        project('a.js', 'const X = 1;')
        project('b.js', 'const X = 2;')
        audit_app.check_global_namespace(report, ['a.js', 'b.js'])
        errors = report.by_severity('ERROR')
        assert len(errors) == 1
        assert "top-level 'X' is declared in multiple files" in errors[0][2]

    def test_function_and_const_name_collision_is_an_error(self, project, report):
        project('a.js', 'function dup() {}')
        project('b.js', 'const dup = 1;')
        audit_app.check_global_namespace(report, ['a.js', 'b.js'])
        errors = report.by_severity('ERROR')
        assert any("declared as both a function" in e[2] for e in errors)


# ─────────────────────────────────────────────────────────────────────────
# 6. Client/server API contract
# ─────────────────────────────────────────────────────────────────────────

class TestApiEndpoints:
    def test_known_function_produces_no_error(self, project, report):
        project('page.html', "Api.call('getClients')")
        audit_app.check_api_endpoints(report, ['page.html'], {'getClients': ['server.js']})
        assert report.items == []

    def test_unknown_function_is_an_error(self, project, report):
        project('page.html', "Api.call('doesNotExist')")
        audit_app.check_api_endpoints(report, ['page.html'], {})
        errors = report.by_severity('ERROR')
        assert len(errors) == 1
        assert "doesNotExist" in errors[0][2]

    def test_private_convention_function_is_a_warning(self, project, report):
        project('page.html', "Api.call('_privateHelper')")
        audit_app.check_api_endpoints(report, ['page.html'], {'_privateHelper': ['server.js']})
        warnings = report.by_severity('WARNING')
        assert len(warnings) == 1
        assert '_privateHelper' in warnings[0][2]


# ─────────────────────────────────────────────────────────────────────────
# 7. LockService usage
# ─────────────────────────────────────────────────────────────────────────

class TestLockUsage:
    def test_no_lock_calls_no_findings(self, project, report):
        project('a.js', 'function f() { return 1; }')
        audit_app.check_lock_usage(report, ['a.js'])
        assert report.items == []

    def test_trylock_without_releaselock_is_an_error(self, project, report):
        project('a.js', 'function f() { lock.tryLock(1000); }')
        audit_app.check_lock_usage(report, ['a.js'])
        errors = report.by_severity('ERROR')
        assert len(errors) == 1
        assert "'f'" in errors[0][2]

    def test_release_outside_finally_is_a_warning(self, project, report):
        project('a.js', 'function f() { lock.tryLock(1000); lock.releaseLock(); }')
        audit_app.check_lock_usage(report, ['a.js'])
        warnings = report.by_severity('WARNING')
        assert len(warnings) == 1
        assert 'not inside a finally block' in warnings[0][2]

    def test_release_inside_finally_has_no_findings(self, project, report):
        project('a.js', (
            'function f() {\n'
            '  lock.tryLock(1000);\n'
            '  try {} finally { lock.releaseLock(); }\n'
            '}\n'
        ))
        audit_app.check_lock_usage(report, ['a.js'])
        assert report.items == []


# ─────────────────────────────────────────────────────────────────────────
# 8. DOM id cross-reference
# ─────────────────────────────────────────────────────────────────────────

class TestDomIds:
    def test_referenced_and_defined_id_has_no_findings(self, project, report):
        project('page.html', '<div id="box"></div><script>getElementById("box")</script>')
        audit_app.check_dom_ids(report, ['page.html'])
        assert report.items == []

    def test_duplicate_id_is_a_warning(self, project, report):
        project('page.html', '<div id="box"></div><div id="box"></div>')
        audit_app.check_dom_ids(report, ['page.html'])
        warnings = report.by_severity('WARNING')
        assert any('appears 2 times' in w[2] for w in warnings)

    def test_reference_to_undefined_id_is_a_warning(self, project, report):
        project('page.html', '<script>document.getElementById("ghost")</script>')
        audit_app.check_dom_ids(report, ['page.html'])
        warnings = report.by_severity('WARNING')
        assert any("'#ghost'" in w[2] for w in warnings)

    def test_jquery_and_query_selector_styles_are_recognized(self, project, report):
        project('page.html', (
            '<div id="a"></div><div id="b"></div>'
            '<script>$("#a"); document.querySelector("#b");</script>'
        ))
        audit_app.check_dom_ids(report, ['page.html'])
        assert report.items == []


# ─────────────────────────────────────────────────────────────────────────
# 9. Config key usage
# ─────────────────────────────────────────────────────────────────────────

class TestConfigKeys:
    def test_missing_config_js_warns(self, project, report):
        audit_app.check_config_keys(report, [])
        warnings = report.by_severity('WARNING')
        assert len(warnings) == 1
        assert 'config.js not found' in warnings[0][2]

    def test_no_frozen_objects_warns(self, project, report):
        project('config.js', 'const X = {};')
        audit_app.check_config_keys(report, ['config.js'])
        warnings = report.by_severity('WARNING')
        assert len(warnings) == 1
        assert 'No Object.freeze()' in warnings[0][2]

    def test_valid_key_access_has_no_findings(self, project, report):
        project('config.js', "const APP_CONFIG = Object.freeze({\n  FOO: 'bar'\n});")
        project('other.js', 'console.log(APP_CONFIG.FOO);')
        audit_app.check_config_keys(report, ['config.js', 'other.js'])
        assert report.items == []

    def test_invalid_key_access_is_an_error(self, project, report):
        project('config.js', "const APP_CONFIG = Object.freeze({\n  FOO: 'bar'\n});")
        project('other.js', 'console.log(APP_CONFIG.MISSING);')
        audit_app.check_config_keys(report, ['config.js', 'other.js'])
        errors = report.by_severity('ERROR')
        assert len(errors) == 1
        assert 'APP_CONFIG.MISSING' in errors[0][2]

    def test_nested_frozen_object_keys_are_tracked(self, project, report):
        project('config.js', (
            "const APP_CONFIG = Object.freeze({\n"
            "  BILL_COL: Object.freeze({\n"
            "    ID: 0\n"
            "  })\n"
            "});"
        ))
        project('other.js', 'console.log(APP_CONFIG.BILL_COL.ID);')
        audit_app.check_config_keys(report, ['config.js', 'other.js'])
        assert report.items == []

    def test_comments_do_not_corrupt_key_parsing(self, project, report):
        project('config.js', (
            "// comment with a brace { and quote '\n"
            "const APP_CONFIG = Object.freeze({\n"
            "  FOO: 'bar' /* inline { comment */\n"
            "});"
        ))
        project('other.js', 'console.log(APP_CONFIG.FOO);')
        audit_app.check_config_keys(report, ['config.js', 'other.js'])
        assert report.items == []

    def test_whitelisted_object_methods_are_not_flagged(self, project, report):
        project('config.js', "const APP_CONFIG = Object.freeze({\n  FOO: 'bar'\n});")
        project('other.js', 'Object.keys(APP_CONFIG); APP_CONFIG.hasOwnProperty("FOO");')
        audit_app.check_config_keys(report, ['config.js', 'other.js'])
        assert report.items == []


# ─────────────────────────────────────────────────────────────────────────
# 10. Markers and console usage
# ─────────────────────────────────────────────────────────────────────────

class TestMarkers:
    def test_todo_marker_is_info(self, project, report):
        project('a.js', '// TODO: fix this later')
        audit_app.check_markers(report, ['a.js'])
        infos = report.by_severity('INFO')
        assert len(infos) == 1
        assert 'TODO' in infos[0][2]

    def test_placeholder_text_is_a_warning(self, project, report):
        project('a.js', "const apiKey = 'YOUR_API_KEY_HERE';")
        audit_app.check_markers(report, ['a.js'])
        warnings = report.by_severity('WARNING')
        assert len(warnings) == 1
        assert 'placeholder' in warnings[0][2]

    def test_debugger_statement_is_a_warning(self, project, report):
        project('a.js', 'function f() {\n  debugger;\n}')
        audit_app.check_markers(report, ['a.js'])
        warnings = report.by_severity('WARNING')
        assert len(warnings) == 1
        assert 'debugger;' in warnings[0][2]

    def test_clean_file_has_no_findings(self, project, report):
        project('a.js', 'function f() { return 1; }')
        audit_app.check_markers(report, ['a.js'])
        assert report.items == []


class TestConsoleUsage:
    def test_console_calls_are_counted_as_info(self, project, report):
        project('a.js', 'console.log("x"); console.warn("y"); console.error("z");')
        audit_app.check_console_usage(report, ['a.js'], [])
        infos = report.by_severity('INFO')
        assert len(infos) == 1
        assert '3 console.* call(s)' in infos[0][2]

    def test_no_console_calls_no_findings(self, project, report):
        project('a.js', 'function f() { return 1; }')
        audit_app.check_console_usage(report, ['a.js'], [])
        assert report.items == []


# ─────────────────────────────────────────────────────────────────────────
# 11. Security posture
# ─────────────────────────────────────────────────────────────────────────

class TestSecurityPosture:
    def test_missing_appsscript_json_no_findings(self, project, report):
        audit_app.check_security_posture(report)
        assert report.items == []

    def test_anyone_with_user_deploying_is_an_info_note(self, project, report):
        project('appsscript.json', json.dumps({
            'webapp': {'access': 'ANYONE', 'executeAs': 'USER_DEPLOYING'}
        }))
        audit_app.check_security_posture(report)
        infos = report.by_severity('INFO')
        assert len(infos) == 1
        assert 'ANYONE' in infos[0][2]

    def test_restricted_access_has_no_findings(self, project, report):
        project('appsscript.json', json.dumps({
            'webapp': {'access': 'DOMAIN', 'executeAs': 'USER_DEPLOYING'}
        }))
        audit_app.check_security_posture(report)
        assert report.items == []
