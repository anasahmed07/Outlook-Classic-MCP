"""save_only for reply/forward: Save() to Drafts, never Send()."""

from __future__ import annotations

from outlook_mcp.client import mail as mc


class _Item:
    def __init__(self):
        self.Subject = "RE: hello"
        self.BodyFormat = 1  # plain text
        self.Body = ""
        self.HTMLBody = ""
        self.EntryID = "draft-1"
        self.To = ""
        self.CC = ""
        self.Attachments = self
        self.saved = False
        self.sent = False

    def Save(self):
        self.saved = True

    def Send(self):
        self.sent = True


class _Original:
    def __init__(self, item):
        self._item = item

    def Reply(self):
        return self._item

    ReplyAll = Forward = Reply


def _patch(monkeypatch, item):
    monkeypatch.setattr(mc, "get_item_by_id", lambda ns, eid, *a, **k: _Original(item))


def test_reply_save_only(monkeypatch):
    item = _Item()
    _patch(monkeypatch, item)
    out = mc.reply_mail(None, None, entry_id="x", body="hi", save_only=True)
    assert out["status"] == "saved_to_drafts" and out["entry_id"] == "draft-1"
    assert item.saved and not item.sent


def test_reply_default_still_sends(monkeypatch):
    item = _Item()
    _patch(monkeypatch, item)
    assert mc.reply_mail(None, None, entry_id="x", body="hi")["status"] == "sent"
    assert item.sent and not item.saved


def test_forward_save_only(monkeypatch):
    item = _Item()
    _patch(monkeypatch, item)
    out = mc.forward_mail(None, None, entry_id="x", to=["a@b.c"], body="fyi", save_only=True)
    assert out["status"] == "saved_to_drafts" and item.saved and not item.sent


def test_forward_default_still_sends(monkeypatch):
    item = _Item()
    _patch(monkeypatch, item)
    assert mc.forward_mail(None, None, entry_id="x", to=["a@b.c"])["status"] == "sent"
    assert item.sent and not item.saved


# ---- OUTLOOK_MCP_DRAFTS_ONLY: nothing is ever sent -------------------------


class _Outlook:
    def __init__(self, item):
        self._item = item

    def CreateItem(self, _kind):
        return self._item


def test_drafts_only_forces_reply_to_drafts(monkeypatch):
    monkeypatch.setenv("OUTLOOK_MCP_DRAFTS_ONLY", "1")
    item = _Item()
    _patch(monkeypatch, item)
    out = mc.reply_mail(None, None, entry_id="x", body="hi")  # save_only not passed
    assert out["status"] == "saved_to_drafts" and "NOT sent" in out["note"]
    assert item.saved and not item.sent


def test_drafts_only_forces_forward_to_drafts(monkeypatch):
    monkeypatch.setenv("OUTLOOK_MCP_DRAFTS_ONLY", "true")
    item = _Item()
    _patch(monkeypatch, item)
    out = mc.forward_mail(None, None, entry_id="x", to=["a@b.c"], save_only=False)
    assert out["status"] == "saved_to_drafts" and item.saved and not item.sent


def test_drafts_only_forces_send_mail_to_drafts(monkeypatch):
    monkeypatch.setenv("OUTLOOK_MCP_DRAFTS_ONLY", "yes")
    item = _Item()
    out = mc.send_mail(
        _Outlook(item), None, to=["a@b.c"], subject="s", body="b", save_only=False
    )
    assert out["status"] == "saved_to_drafts" and "NOT sent" in out["note"]
    assert item.saved and not item.sent


def test_drafts_only_off_values_still_send(monkeypatch):
    for value in ("", "0", "off", "no"):
        monkeypatch.setenv("OUTLOOK_MCP_DRAFTS_ONLY", value)
        item = _Item()
        out = mc.send_mail(
            _Outlook(item), None, to=["a@b.c"], subject="s", body="b"
        )
        assert out["status"] == "sent" and item.sent and not item.saved
