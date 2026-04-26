"""Tests for utils.paths — attachment + output_dir validation."""

import os
from pathlib import Path

import pytest

from outlook_mcp.errors import OutlookError
from outlook_mcp.utils.paths import validate_attachment_path, validate_output_dir


def test_attachment_rejects_empty():
    with pytest.raises(OutlookError):
        validate_attachment_path("")


def test_attachment_rejects_relative():
    with pytest.raises(OutlookError, match="absolute"):
        validate_attachment_path("just/a/relative.txt")


def test_attachment_rejects_missing(tmp_path):
    missing = tmp_path / "nope.txt"
    with pytest.raises(OutlookError, match="not found"):
        validate_attachment_path(str(missing))


def test_attachment_accepts_real_file_under_profile(monkeypatch, tmp_path):
    monkeypatch.setenv("USERPROFILE", str(tmp_path))
    f = tmp_path / "ok.txt"
    f.write_text("hi")
    out = validate_attachment_path(str(f))
    assert os.path.normcase(out) == os.path.normcase(str(f.resolve()))


def test_attachment_rejects_outside_profile(monkeypatch, tmp_path):
    profile = tmp_path / "profile"
    profile.mkdir()
    elsewhere = tmp_path / "elsewhere.txt"
    elsewhere.write_text("nope")
    monkeypatch.setenv("USERPROFILE", str(profile))
    monkeypatch.delenv("OUTLOOK_MCP_ALLOW_ANY_PATH", raising=False)
    with pytest.raises(OutlookError, match="outside the user profile"):
        validate_attachment_path(str(elsewhere))


def test_attachment_allows_anywhere_with_env(monkeypatch, tmp_path):
    profile = tmp_path / "profile"
    profile.mkdir()
    elsewhere = tmp_path / "elsewhere.txt"
    elsewhere.write_text("ok")
    monkeypatch.setenv("USERPROFILE", str(profile))
    monkeypatch.setenv("OUTLOOK_MCP_ALLOW_ANY_PATH", "1")
    assert validate_attachment_path(str(elsewhere))


def test_output_dir_creates(monkeypatch, tmp_path):
    monkeypatch.setenv("USERPROFILE", str(tmp_path))
    target = tmp_path / "fresh" / "dir"
    out = validate_output_dir(str(target))
    assert Path(out).is_dir()


def test_output_dir_rejects_relative():
    with pytest.raises(OutlookError, match="absolute"):
        validate_output_dir("rel/dir")
