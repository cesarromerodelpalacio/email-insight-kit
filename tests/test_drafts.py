from datetime import datetime

from inboxkit.drafts import DEFAULT_TEMPLATES, priority_records, suggest_reply
from inboxkit.models import EmailRecord
from inboxkit.outlook_source import synthetic_inbox


def _rec(email, subj, unread=True):
    return EmailRecord("Ana Ruiz", email, subj, datetime.now(), unread, entry_id="X1")


def test_suggest_reply_personal():
    r = _rec("jefe@empresa.com", "Pregunta rapida")
    body = suggest_reply(r)
    assert body.startswith("Hola Ana") and "reviso" in body.lower()


def test_suggest_reply_newsletter_empty():
    r = _rec("promo@mailchimp.com", "Newsletter julio")
    assert suggest_reply(r) == ""   # no se responde a newsletters


def test_suggest_reply_reunion():
    r = _rec("invitation@teams.com", "Invitacion: sync")
    assert "reunion" in suggest_reply(r).lower()


def test_priority_records_filters():
    recs = synthetic_inbox(200, seed=3)
    prio = priority_records(recs, limit=10)
    assert len(prio) <= 10
    # todos son no leidos y personal/reunion
    from inboxkit.models import categorize
    assert all(r.unread and categorize(r) in ("personal", "reunion") for r in prio)
