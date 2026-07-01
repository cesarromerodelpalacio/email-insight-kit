"""Analitica de bandeja: volumen, remitentes, dominios, franjas, categorias."""
from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime

from .models import EmailRecord, categorize


@dataclass
class InboxReport:
    total: int
    unread: int
    generated_at: str
    dias_analizados: int
    top_senders: list = field(default_factory=list)       # [(email, count)]
    top_domains: list = field(default_factory=list)
    by_day: list = field(default_factory=list)             # [(fecha, count)]
    by_hour: list = field(default_factory=list)            # [(hora, count)] 0-23
    categorias: list = field(default_factory=list)         # [(cat, count)]
    prioridad: list = field(default_factory=list)          # no leidos importantes

    @property
    def unread_rate(self) -> int:
        return round(100 * self.unread / self.total) if self.total else 0


def analyze(records: list[EmailRecord], dias: int = 30) -> InboxReport:
    total = len(records)
    unread = sum(1 for r in records if r.unread)
    senders = Counter(r.sender_email.lower() for r in records if r.sender_email)
    domains = Counter(r.domain for r in records)
    days = Counter(r.received.strftime("%Y-%m-%d") for r in records)
    hours = Counter(r.received.hour for r in records)
    cats = Counter(categorize(r) for r in records)

    # prioridad: no leidos que NO son newsletter/notificacion (lo que de verdad requiere accion)
    prioridad = [
        (r.sender_email, r.subject)
        for r in records
        if r.unread and categorize(r) in ("personal", "reunion")
    ][:15]

    return InboxReport(
        total=total,
        unread=unread,
        generated_at=datetime.now().strftime("%Y-%m-%d %H:%M"),
        dias_analizados=dias,
        top_senders=senders.most_common(10),
        top_domains=domains.most_common(8),
        by_day=sorted(days.items()),
        by_hour=[(h, hours.get(h, 0)) for h in range(24)],
        categorias=cats.most_common(),
        prioridad=prioridad,
    )
