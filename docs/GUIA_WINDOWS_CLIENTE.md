# Guia rapida Windows - MICE Travel Bot

## Instalacion desde cero

En cada PC Windows, abrir PowerShell y ejecutar:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
iwr -UseB https://raw.githubusercontent.com/raulmm78/mice-travel-bot/main/instalar_desde_github_windows.ps1 -OutFile "$env:TEMP\instalar_mice_travel_bot.ps1"
powershell -ExecutionPolicy Bypass -File "$env:TEMP\instalar_mice_travel_bot.ps1"
```

Esto instala o comprueba:

- Python 3.
- Git.
- Dependencias Python.
- Proyecto desde GitHub.
- Accesos directos en el escritorio.

El cliente no necesita ChatGPT ni Codex.

## Configuracion

Abrir el acceso directo `Configurar MICE Travel Bot` y rellenar `config/.env`.

Minimo:

```text
OPENAI_API_KEY=
IMAP_HOST=
IMAP_PORT=993
IMAP_USER=
IMAP_PASSWORD=
IMAP_FOLDER=INBOX
SMTP_HOST=
SMTP_PORT=587
SMTP_USER=
SMTP_PASSWORD=
MAIL_FROM=
MAIL_TO=
ALERT_EMAIL_TO=
OUTPUT_DIR=
EVENT_OUTPUT_DIR=
PROCESSED_IDS_PATH=
NN_TEMPLATE_PATH=
WATCH_INTERVAL_SECONDS=300
```

## Arranque

Usar el acceso directo:

```text
MICE Travel Bot
```

Al arrancar:

1. Busca actualizaciones en GitHub.
2. Si hay version nueva, la descarga.
3. Abre el panel local.
4. No reinicia el PC.

Antes de la primera prueba, usar el acceso directo:

```text
Diagnostico MICE Travel Bot
```

Debe mostrar OK en:

- Python.
- OpenAI API key.
- Configuracion IMAP.
- Conexion IMAP.
- Plantilla Excel.
- Carpeta global escribible.
- Carpeta eventos escribible.
- Registro anti-duplicados escribible.

## Recomendacion para varios ordenadores

Primera version:

- Solo un PC con el bot en `ON`.
- Todos pueden abrir/ver los Excel desde OneDrive.
- El registro `_bot_processed_message_ids.json` debe estar en la carpeta compartida de OneDrive.

Si varios PCs tienen el bot en `ON`, hay menos riesgo de duplicados por el `Message-ID`, pero sigue pudiendo haber conflicto si OneDrive tarda en sincronizar.
