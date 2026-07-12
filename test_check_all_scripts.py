"""
Unit tests for check_all_scripts.py — extracts inline <script> blocks from
dist/index.html and syntax-checks each one with `node --check`.

check_all_scripts() reads from a hardcoded "dist/index.html" path relative to the
current working directory and writes/removes temp_check_N.js files in the cwd, so
tests run inside a temp directory (via monkeypatch.chdir) to avoid touching the
real project tree.
"""

import os

import pytest

from check_all_scripts import check_all_scripts


@pytest.fixture
def sandbox(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    return tmp_path


def write_dist(sandbox, html):
    dist = sandbox / 'dist'
    dist.mkdir()
    (dist / 'index.html').write_text(html, encoding='utf-8')


class TestCheckAllScripts:
    def test_missing_dist_index_prints_message_and_returns(self, sandbox, capsys):
        check_all_scripts()
        assert 'dist/index.html not found.' in capsys.readouterr().out

    def test_valid_script_reports_success(self, sandbox, capsys):
        write_dist(sandbox, '<html><script>var x = 1;</script></html>')
        check_all_scripts()
        out = capsys.readouterr().out
        assert 'Script tag 1: SUCCESS' in out

    def test_invalid_script_reports_failure_with_stderr(self, sandbox, capsys):
        write_dist(sandbox, '<html><script>function( {</script></html>')
        check_all_scripts()
        out = capsys.readouterr().out
        assert 'Script tag 1: FAILED' in out

    def test_script_with_src_is_skipped(self, sandbox, capsys):
        write_dist(sandbox, '<html><script src="lib.js">garbage((</script></html>')
        check_all_scripts()
        out = capsys.readouterr().out
        assert 'Script tag' not in out

    def test_multiple_scripts_are_indexed_in_order(self, sandbox, capsys):
        write_dist(sandbox, (
            '<script>var a = 1;</script>'
            '<script src="lib.js">ignored</script>'
            '<script>var b = 2;</script>'
        ))
        check_all_scripts()
        out = capsys.readouterr().out
        assert 'Script tag 1: SUCCESS' in out
        assert 'Script tag 2: SUCCESS' in out
        assert 'Script tag 3' not in out

    def test_temp_files_are_cleaned_up_after_success(self, sandbox):
        write_dist(sandbox, '<html><script>var x = 1;</script></html>')
        check_all_scripts()
        leftovers = [f for f in os.listdir(sandbox) if f.startswith('temp_check_')]
        assert leftovers == []

    def test_temp_files_are_cleaned_up_even_after_failure(self, sandbox):
        # The final cleanup pass removes any leftover temp_check_*.js
        # unconditionally, so a failed syntax check doesn't leave debris behind.
        write_dist(sandbox, '<html><script>function( {</script></html>')
        check_all_scripts()
        leftovers = [f for f in os.listdir(sandbox) if f.startswith('temp_check_')]
        assert leftovers == []
