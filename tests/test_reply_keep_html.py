"""Replies/forwards must not flatten an HTML message to plain text."""

from __future__ import annotations

from outlook_mcp.client import mail as mc

SIG = "<html><body><p>sig <a href='https://x'>link</a></p></body></html>"


class _Item:
    def __init__(self, fmt, html=SIG, body="plain original"):
        self.BodyFormat = fmt
        self.HTMLBody = html
        self._body = body
        self.body_assigned = False
        self.Subject = "RE: x"
        self.To = ""
        self.CC = ""
        self.Attachments = self
        self.sent = False

    @property
    def Body(self):
        return self._body

    @Body.setter
    def Body(self, value):
        self._body = value
        self.body_assigned = True

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


def test_reply_html_item_keeps_html(monkeypatch):
    item = _Item(fmt=2)
    _patch(monkeypatch, item)
    mc.reply_mail(None, None, entry_id="x", body="hello <b>& bye\nline2")
    assert not item.body_assigned
    assert item.HTMLBody.startswith("<html><body>hello &lt;b&gt;&amp; bye<br>line2<br><br><p>sig")
    assert "href='https://x'" in item.HTMLBody


def test_reply_html_without_body_tag(monkeypatch):
    item = _Item(fmt=2, html="<p>quoted</p>")
    _patch(monkeypatch, item)
    mc.reply_mail(None, None, entry_id="x", body="hi")
    assert item.HTMLBody == "hi<br><br><p>quoted</p>"


def test_reply_plain_item_uses_body(monkeypatch):
    item = _Item(fmt=1)
    _patch(monkeypatch, item)
    mc.reply_mail(None, None, entry_id="x", body="hi")
    assert item.Body == "hi\n\nplain original"


def test_forward_html_item_keeps_html(monkeypatch):
    item = _Item(fmt=2)
    _patch(monkeypatch, item)
    mc.forward_mail(None, None, entry_id="x", to=["a@b.c"], body="fyi")
    assert not item.body_assigned and "fyi<br><br><p>sig" in item.HTMLBody


def test_forward_without_note_untouched(monkeypatch):
    item = _Item(fmt=2)
    _patch(monkeypatch, item)
    mc.forward_mail(None, None, entry_id="x", to=["a@b.c"])
    assert item.HTMLBody == SIG and not item.body_assigned
