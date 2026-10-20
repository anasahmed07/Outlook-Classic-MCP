"""Tests for _mail_full()'s attachment handling."""

import sys

import pytest

pytestmark = pytest.mark.skipif(
    sys.platform != "win32", reason="client modules import pywin32"
)


class _NormalAttachment:
    FileName = "report.pdf"
    Size = 12345


class _BrokenFileNameAttachment:
    """Simulates an embedded .msg / OLE attachment: FileName and Size raise.

    Verified live against real Outlook data: certain attachment types
    raise pywintypes.com_error on these properties.
    """

    DisplayName = "embedded.msg"

    @property
    def FileName(self):
        raise Exception("Kan de actie niet uitvoeren op dit type bijlage.")

    @property
    def Size(self):
        raise Exception("Kan de actie niet uitvoeren op dit type bijlage.")


class _FakeAttachments:
    def __init__(self, items):
        self._items = items
        self.Count = len(items)

    def __iter__(self):
        return iter(self._items)


class _ExplodingAttachments:
    """Touching this at all (Count or iteration) fails the test."""

    @property
    def Count(self):
        raise AssertionError("Attachments.Count must not be accessed when include_attachments=False")

    def __iter__(self):
        raise AssertionError("Attachments must not be enumerated when include_attachments=False")


def _fake_item(attachments):
    from types import SimpleNamespace

    return SimpleNamespace(
        EntryID="id1",
        ConversationID="conv1",
        Subject="Test",
        SenderName="Someone",
        SenderEmailAddress="someone@example.com",
        To="",
        CC="",
        BCC="",
        ReceivedTime=None,
        SentOn=None,
        UnRead=False,
        FlagStatus=0,
        Importance=1,
        Categories="",
        Body="hello",
        Attachments=attachments,
    )


def test_mail_full_falls_back_when_attachment_property_raises():
    from outlook_mcp.client.mail import _mail_full

    item = _fake_item(_FakeAttachments([_NormalAttachment(), _BrokenFileNameAttachment()]))
    result = _mail_full(item)

    assert result["attachment_count"] == 2
    assert result["attachments"][0]["filename"] == "report.pdf"
    assert result["attachments"][0]["size_bytes"] == 12345
    # Falls back to DisplayName when FileName raises, instead of failing
    # the whole outlook_get_mail call.
    assert result["attachments"][1]["filename"] == "embedded.msg"
    assert result["attachments"][1]["size_bytes"] is None


def test_mail_full_names_unavailable_attachment_when_no_property_works():
    from outlook_mcp.client.mail import _mail_full

    class _TotallyBroken(_BrokenFileNameAttachment):
        DisplayName = ""  # falsy too — nothing to fall back to

    item = _fake_item(_FakeAttachments([_TotallyBroken()]))
    result = _mail_full(item)

    assert result["attachments"][0]["filename"] == "(attachment 1, name unavailable)"


def test_mail_full_skips_attachments_entirely_when_disabled():
    from outlook_mcp.client.mail import _mail_full

    item = _fake_item(_ExplodingAttachments())
    result = _mail_full(item, include_attachments=False)

    assert result["attachments"] is None
    assert "attachment_count" not in result
