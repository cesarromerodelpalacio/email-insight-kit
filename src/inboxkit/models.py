"""Modelo generico de correo (desacoplado de Outlook para poder testear)."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime


@dataclass
class EmailRecord:
    sender_name: str
    sender_email: str
    subject: str
    received: datetime
    unread: bool = False

    @property
    def domain(self) -> str:
        e = (self.sender_email or "").strip().lower()
        return e.split("@")[-1] if "@" in e else "(desconocido)"


# Reglas simples de categorizacion (keyword/dominio). Ampliables.
_RULES = {
    "newsletter": ["newsletter", "unsubscribe", "noreply", "no-reply", "mailchimp", "sendgrid", "marketing", "boletin"],
    "reunion": ["invitation", "meeting", "calendar", "teams", "zoom", "reunion", "invitacion", "cita"],
    "notificacion": ["notification", "alert", "automated", "do-not-reply", "aviso", "notificacion", "report"],
}


def categorize(rec: EmailRecord) -> str:
    hay = f"{rec.sender_email} {rec.subject}".lower()
    for cat, kws in _RULES.items():
        if any(k in hay for k in kws):
            return cat
    return "personal"
