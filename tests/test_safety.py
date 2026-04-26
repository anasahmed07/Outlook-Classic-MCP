"""Tests for utils.safety — safe_dasl escaping."""

from outlook_mcp.utils.safety import safe_dasl


def test_safe_dasl_passthrough():
    assert safe_dasl("hello world") == "hello world"


def test_safe_dasl_escapes_sql_wildcards():
    assert safe_dasl("50% off") == "50[%] off"
    assert safe_dasl("first_name") == "first[_]name"


def test_safe_dasl_escapes_quotes():
    assert safe_dasl("o'reilly") == "o''reilly"
    assert safe_dasl('say "hi"') == 'say ""hi""'


def test_safe_dasl_escapes_combined():
    assert safe_dasl("a%b_c'd\"e") == "a[%]b[_]c''d\"\"e"


def test_safe_dasl_handles_none():
    assert safe_dasl(None) == ""
