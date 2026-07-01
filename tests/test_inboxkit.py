from datetime import datetime, timedelta

from inboxkit.analyze import analyze
from inboxkit.models import EmailRecord, categorize
from inboxkit.report import build_html, write_excel
from inboxkit.outlook_source import synthetic_inbox


def test_domain_and_categorize():
    r = EmailRecord("N", "promo@mailchimp.com", "Newsletter julio", datetime.now(), True)
    assert r.domain == "mailchimp.com"
    assert categorize(r) == "newsletter"
    r2 = EmailRecord("J", "jefe@empresa.com", "Pregunta rapida", datetime.now())
    assert categorize(r2) == "personal"


def test_synthetic_and_analyze():
    recs = synthetic_inbox(n=200, seed=7)
    assert len(recs) == 200
    rep = analyze(recs, dias=30)
    assert rep.total == 200
    assert 0 <= rep.unread <= 200
    assert rep.top_senders and rep.top_senders[0][1] > 0
    assert len(rep.by_hour) == 24
    assert rep.categorias  # hay categorias
    assert rep.unread_rate == round(100 * rep.unread / rep.total)


def test_reports(tmp_path):
    rep = analyze(synthetic_inbox(120), dias=30)
    html = build_html(rep)
    assert "Panel de bandeja" in html and "Prioridad" in html
    x = write_excel(rep, tmp_path / "b.xlsx")
    assert x.exists() and x.stat().st_size > 0
