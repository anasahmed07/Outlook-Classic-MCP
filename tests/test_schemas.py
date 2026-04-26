"""Tests for schemas — Recurrence model."""

import pytest
from pydantic import ValidationError

from outlook_mcp.schemas import Recurrence, ResponseFormat


def test_recurrence_minimal():
    r = Recurrence(type="weekly")
    assert r.type == "weekly"
    assert r.interval == 1
    assert r.occurrences is None
    assert r.end_date is None


def test_recurrence_full():
    r = Recurrence(type="monthly", interval=2, occurrences=12)
    assert r.interval == 2
    assert r.occurrences == 12


def test_recurrence_rejects_bad_type():
    with pytest.raises(ValidationError):
        Recurrence(type="hourly")


def test_recurrence_rejects_zero_interval():
    with pytest.raises(ValidationError):
        Recurrence(type="daily", interval=0)


def test_recurrence_rejects_zero_occurrences():
    with pytest.raises(ValidationError):
        Recurrence(type="daily", occurrences=0)


def test_response_format_values():
    assert ResponseFormat("json") == ResponseFormat.JSON
    assert ResponseFormat("markdown") == ResponseFormat.MARKDOWN
    with pytest.raises(ValueError):
        ResponseFormat("xml")
