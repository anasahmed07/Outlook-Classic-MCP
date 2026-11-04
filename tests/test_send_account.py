"""Tests for the optional sender `account` parameter of send_mail."""

import sys
from unittest.mock import MagicMock

import pytest

pytestmark = pytest.mark.skipif(
    sys.platform != "win32", reason="client modules import pywin32"
)


def _fake_outlook(addresses):
    accounts = []
    for addr in addresses:
        acc = MagicMock()
        acc.SmtpAddress = addr
        acc.DisplayName = addr
        accounts.append(acc)
    coll = MagicMock()
    coll.Count = len(accounts)
    coll.Item.side_effect = lambda i: accounts[i - 1]
    outlook = MagicMock()
    outlook.Session.Accounts = coll
    return outlook, accounts


def test_find_account_is_case_insensitive():
    from outlook_mcp.client.mail import find_account

    outlook, accounts = _fake_outlook(["a@x.com", "Info@Hotel.com"])
    assert find_account(outlook, "info@hotel.com") is accounts[1]


def test_find_account_unknown_lists_available():
    from outlook_mcp.client.mail import find_account
    from outlook_mcp.errors import OutlookError

    outlook, _ = _fake_outlook(["a@x.com"])
    with pytest.raises(OutlookError, match="a@x.com"):
        find_account(outlook, "nope@x.com")


def test_send_mail_with_account_saves_into_account_drafts():
    from outlook_mcp.client.mail import send_mail

    outlook, accounts = _fake_outlook(["a@x.com", "info@hotel.com"])
    mail = MagicMock()
    accounts[1].DeliveryStore.GetDefaultFolder.return_value.Items.Add.return_value = mail

    send_mail(
        outlook, MagicMock(), to=["g@y.com"], subject="s", body="b",
        save_only=True, account="info@hotel.com",
    )

    assert mail.SendUsingAccount is accounts[1]
    mail.Save.assert_called_once()
    outlook.CreateItem.assert_not_called()


def test_send_mail_without_account_uses_default_profile():
    from outlook_mcp.client.mail import send_mail

    outlook, _ = _fake_outlook(["a@x.com"])
    send_mail(outlook, MagicMock(), to=["g@y.com"], subject="s", body="b", save_only=True)
    outlook.CreateItem.assert_called_once()


def test_send_with_signature_closes_inspector_and_sends_reloaded_item(monkeypatch):
    from outlook_mcp.client import mail as mail_mod

    outlook, accounts = _fake_outlook(["info@hotel.com"])
    draft = MagicMock()
    draft.HTMLBody = "<html><body><p>SIG</p></body></html>"
    draft.EntryID, draft.Parent.StoreID = "E1", "S1"
    accounts[0].DeliveryStore.GetDefaultFolder.return_value.Items.Add.return_value = draft
    reloaded = MagicMock()
    monkeypatch.setattr(mail_mod, "get_item_by_id", lambda ns, eid, sid=None: reloaded)

    mail_mod.send_mail(outlook, MagicMock(), to=["g@y.com"], subject="s", body="Hallo",
                       account="info@hotel.com", signature=True)

    draft.Save.assert_called_once()
    draft.GetInspector.Close.assert_called_once_with(0)
    draft.Send.assert_not_called()
    reloaded.Send.assert_called_once()
    assert "Hallo" in draft.HTMLBody and "SIG" in draft.HTMLBody
