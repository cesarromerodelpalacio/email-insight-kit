"""Creacion de BORRADORES de respuesta (humano en el bucle: nunca envia).

suggest_reply es puro (testeable). create_draft_replies opera sobre Outlook via COM
y SOLO guarda borradores; el usuario revisa y envia."""
from __future__ import annotations

from .models import EmailRecord, categorize

# Plantillas por categoria (el primer nombre del remitente se inserta en {name}).
DEFAULT_TEMPLATES = {
    "personal": "Hola {name},\n\nGracias por tu correo. Lo reviso y te respondo con detalle en breve.\n\nUn saludo,",
    "reunion": "Hola {name},\n\nConfirmo la reunion. Si necesitas que prepare algo, dime.\n\nUn saludo,",
    "notificacion": "",  # normalmente no requieren respuesta
    "newsletter": "",
}


def _first_name(rec: EmailRecord) -> str:
    n = (rec.sender_name or "").strip()
    return n.split()[0] if n else "hola"


def suggest_reply(rec: EmailRecord, templates: dict | None = None) -> str:
    """Devuelve el texto de respuesta sugerido (o '' si la categoria no lo amerita)."""
    t = templates or DEFAULT_TEMPLATES
    tpl = t.get(categorize(rec), t.get("personal", ""))
    return tpl.format(name=_first_name(rec)) if tpl else ""


def priority_records(records: list[EmailRecord], limit: int = 10) -> list[EmailRecord]:
    """No leidos que requieren accion (personal/reunion), con id de Outlook."""
    return [r for r in records if r.unread and categorize(r) in ("personal", "reunion")][:limit]


def create_draft_replies(records: list[EmailRecord], templates: dict | None = None, limit: int = 10) -> dict:
    """Crea BORRADORES de respuesta en Outlook (COM). NUNCA envia. Devuelve reporte."""
    import win32com.client as w

    ns = w.Dispatch("Outlook.Application").GetNamespace("MAPI")
    created, skipped = 0, 0
    for rec in priority_records(records, limit):
        body = suggest_reply(rec, templates)
        if not body or not rec.entry_id:
            skipped += 1
            continue
        try:
            original = ns.GetItemFromID(rec.entry_id)
            reply = original.Reply()
            reply.Body = body + "\n\n" + (reply.Body or "")
            reply.Save()  # BORRADOR; el usuario revisa y envia
            created += 1
        except Exception:
            skipped += 1
    return {"borradores_creados": created, "omitidos": skipped}
