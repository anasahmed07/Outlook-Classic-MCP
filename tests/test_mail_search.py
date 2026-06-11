"""Tests for the pure-Python parts of mail search."""

import sys

import pytest

pytestmark = pytest.mark.skipif(
    sys.platform != "win32", reason="client modules import pywin32"
)


def test_split_single_word():
    from outlook_mcp.client.mail import split_search_words

    anchor, remaining = split_search_words("Budget")
    assert anchor == "budget"
    assert remaining == []


def test_split_multi_word_picks_longest_anchor():
    from outlook_mcp.client.mail import split_search_words

    anchor, remaining = split_search_words("teams not working")
    assert anchor == "working"
    assert sorted(remaining) == ["not", "teams"]


def test_split_handles_duplicate_longest():
    from outlook_mcp.client.mail import split_search_words

    anchor, remaining = split_search_words("alpha alpha beta")
    assert anchor == "alpha"
    # only one instance of the anchor is removed
    assert sorted(remaining) == ["alpha", "beta"]


def test_split_whitespace_only_falls_back_to_raw():
    from outlook_mcp.client.mail import split_search_words

    anchor, remaining = split_search_words("  ")
    assert anchor == "  "
    assert remaining == []
