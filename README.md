# Email Insight Kit

Analitica **local** de tu bandeja de Outlook. Te dice **quien te inunda**, cuantos no leidos
tienes, en que franjas horarias llega el correo, que es newsletter vs importante, y una lista de
**no leidos que requieren accion**. Panel HTML + Excel.

Todo se procesa **en tu maquina** (Outlook Desktop via COM). Nada sube a la nube. Por
**Cesar Romero del Palacio** · Licencia MIT.

## Uso
```bash
pip install -e .
pip install pywin32           # para leer Outlook (Windows)
inboxkit --days 30            # lee tu bandeja y genera el panel
inboxkit --demo               # prueba con bandeja sintetica (sin tocar Outlook)
```

Requiere **Outlook Desktop abierto** y tu cuenta logueada (no usa Graph ni OAuth).

## Desarrollo
```bash
pip install -e ".[dev]"
pytest -v
```
