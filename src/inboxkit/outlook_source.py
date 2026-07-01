"""Fuentes de datos: Outlook (COM, local) y generador sintetico para demo/test."""
from __future__ import annotations

import random
from datetime import datetime, timedelta

from .models import EmailRecord


def read_outlook(days: int = 30, limit: int = 3000) -> list[EmailRecord]:
    """Lee la bandeja de Outlook Desktop via COM. Requiere Outlook abierto + pywin32."""
    import win32com.client as w

    ns = w.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = ns.GetDefaultFolder(6)
    items = inbox.Items
    items.Sort("[ReceivedTime]", True)
    corte = datetime.now() - timedelta(days=days)
    out: list[EmailRecord] = []
    for m in items:
        if len(out) >= limit:
            break
        try:
            rt = m.ReceivedTime
            recv = datetime(rt.year, rt.month, rt.day, rt.hour, rt.minute)
            if recv < corte:
                break
            out.append(EmailRecord(
                sender_name=getattr(m, "SenderName", "") or "",
                sender_email=getattr(m, "SenderEmailAddress", "") or "",
                subject=getattr(m, "Subject", "") or "",
                received=recv,
                unread=bool(getattr(m, "UnRead", False)),
                entry_id=getattr(m, "EntryID", "") or "",
            ))
        except Exception:
            continue
    return out


def synthetic_inbox(n: int = 200, seed: int = 7) -> list[EmailRecord]:
    """Bandeja sintetica realista para demo/test (sin datos reales)."""
    rnd = random.Random(seed)
    senders = [
        ("Boletin Tech", "newsletter@techmail.com"), ("RRHH", "rrhh@empresa.com"),
        ("Jefe Equipo", "jefe@empresa.com"), ("Cliente A", "contacto@clientea.com"),
        ("No Reply", "no-reply@notificaciones.com"), ("Calendario", "invitation@teams.com"),
        ("Proveedor", "ventas@proveedor.com"), ("Marketing", "promo@mailchimp.com"),
        ("Compañero", "colega@empresa.com"), ("Soporte", "support@saas.io"),
    ]
    subjects = ["Reunion semanal", "Tu factura", "Newsletter de julio", "Accion requerida: revisar",
                "Invitacion: sync", "Alerta del sistema", "Propuesta comercial", "Recordatorio", "Pregunta rapida", "Informe"]
    out = []
    for _ in range(n):
        nm, em = rnd.choice(senders)
        recv = datetime.now() - timedelta(days=rnd.randint(0, 29), hours=rnd.randint(0, 23), minutes=rnd.randint(0, 59))
        out.append(EmailRecord(nm, em, rnd.choice(subjects), recv, unread=rnd.random() < 0.35))
    return out

