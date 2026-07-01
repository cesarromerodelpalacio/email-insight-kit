"""CLI de Email Insight Kit."""
from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

from . import __version__
from .analyze import analyze
from .report import write_excel, write_html


def build_parser():
    p = argparse.ArgumentParser(prog="inboxkit",
        description="Analitica local de tu bandeja Outlook + borradores de respuesta (para supervisar y enviar).")
    p.add_argument("--days", type=int, default=30)
    p.add_argument("--limit", type=int, default=3000)
    p.add_argument("--demo", action="store_true", help="Bandeja sintetica (sin tocar Outlook).")
    p.add_argument("--draft-replies", action="store_true",
                   help="Crea BORRADORES de respuesta para los correos prioritarios (nunca envia; tu supervisas).")
    p.add_argument("--reply-limit", type=int, default=10, help="Maximo de borradores a crear.")
    p.add_argument("-o", "--output", default="outputs")
    p.add_argument("--no-excel", action="store_true")
    p.add_argument("--version", action="version", version=f"inboxkit {__version__}")
    return p


def run(args):
    if args.demo:
        from .outlook_source import synthetic_inbox
        records = synthetic_inbox()
        print(f"[inboxkit] Modo demo: {len(records)} correos sinteticos.")
    else:
        try:
            from .outlook_source import read_outlook
            records = read_outlook(days=args.days, limit=args.limit)
        except ImportError:
            print("[inboxkit] Falta pywin32 (o no es Windows). Usa --demo.", file=sys.stderr); return 2
        except Exception as e:
            print(f"[inboxkit] No se pudo leer Outlook (abre Outlook Desktop): {e}", file=sys.stderr); return 2
        print(f"[inboxkit] Leidos {len(records)} correos de los ultimos {args.days} dias.")

    if not records:
        print("[inboxkit] Sin correos en el rango.", file=sys.stderr); return 2

    rep = analyze(records, dias=args.days)
    out = Path(args.output); stamp = datetime.now().strftime("%Y%m%d_%H%M")
    print(f"[inboxkit] HTML  -> {write_html(rep, out / f'bandeja_{stamp}.html')}")
    if not args.no_excel:
        print(f"[inboxkit] Excel -> {write_excel(rep, out / f'bandeja_{stamp}.xlsx')}")
    print(f"[inboxkit] {rep.total} correos, {rep.unread} sin leer ({rep.unread_rate}%). Prioridad: {len(rep.prioridad)} accionables.")

    if args.draft_replies:
        from .drafts import priority_records, suggest_reply
        prio = priority_records(records, args.reply_limit)
        if args.demo:
            print(f"[inboxkit] (demo) borradores que se crearian: {sum(1 for r in prio if suggest_reply(r))}")
            for r in prio[:3]:
                if suggest_reply(r):
                    print(f"   -> a {r.sender_email}: '{r.subject[:40]}'")
        else:
            from .drafts import create_draft_replies
            res = create_draft_replies(records, limit=args.reply_limit)
            print(f"[inboxkit] Borradores de respuesta creados: {res['borradores_creados']} (omitidos {res['omitidos']}).")
            print("[inboxkit] Revisalos en Outlook > Borradores y envia los que quieras. NADA se envio solo.")
    return 0


def main(argv=None):
    args = build_parser().parse_args(argv)
    try:
        return run(args)
    except (FileNotFoundError, ValueError) as e:
        print(f"[inboxkit] Error: {e}", file=sys.stderr); return 1


if __name__ == "__main__":
    raise SystemExit(main())
