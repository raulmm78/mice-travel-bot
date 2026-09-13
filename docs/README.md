# MICE Travel Bot

Bot local para leer correos de solicitudes de viaje, extraer datos, generar Excel global y crear Excel por evento.

## Estructura

```text
MICE Travel Bot/
  abrir_panel_mac.command
  abrir_panel_windows.bat
  diagnostico_mac.command
  diagnostico_windows.bat
  instalar_mac.command
  instalar_windows.bat
  cerrar_panel_mac.command
  cerrar_panel_windows.bat
  instalar_desde_github_windows.ps1

  app/
    process_emails.py
    update_before_start.py
    requirements.txt
    run_panel_mac.command
    run_panel_windows.bat
    assets/

  config/
    .env.example
    .env

  docs/
    README.md
    CHECKLIST_ENTREGA.md
    DATOS_A_PEDIR_CLIENTE.md
    GUIA_WINDOWS_CLIENTE.md
    GUIA_MAC_PRUEBAS.md

  examples/
    demo_email_ok.txt
    demo_email_pendiente.txt

  data/
    emails/
    emails_archive/
    outputs/
```

En la raiz solo quedan los archivos que normalmente se ejecutan a doble clic.

## Archivos que se ejecutan

Mac:

```text
instalar_mac.command
diagnostico_mac.command
abrir_panel_mac.command
cerrar_panel_mac.command
```

Windows:

```text
instalar_windows.bat
diagnostico_windows.bat
abrir_panel_windows.bat
cerrar_panel_windows.bat
```

Instalacion desde GitHub en Windows:

```text
instalar_desde_github_windows.ps1
```

## Configuracion

Las claves van en:

```text
config/.env
```

La plantilla sin claves esta en:

```text
config/.env.example
```

El archivo `config/.env` no se sube a GitHub.

Variables principales:

```text
OPENAI_API_KEY=
OPENAI_MODEL=gpt-4.1-mini
EMAIL_AGENT_PORT=8765
AUTO_UPDATE_ENABLED=1
AUTO_UPDATE_BRANCH=main

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
XLSX_PATH=
EVENT_OUTPUT_DIR=
EVENT_ROUTES_PATH=
PROCESSED_IDS_PATH=
NN_TEMPLATE_PATH=
WATCH_INTERVAL_SECONDS=300
```

## OneDrive

Primera version recomendada: escribir en una carpeta local sincronizada de OneDrive.

Ejemplo Windows:

```text
OUTPUT_DIR=C:\Users\Usuario\OneDrive - MICE TRAVEL\EMPRESAS\Novo Nordisk\_global
XLSX_PATH=C:\Users\Usuario\OneDrive - MICE TRAVEL\EMPRESAS\Novo Nordisk\_global\viajes_global.xlsx
EVENT_OUTPUT_DIR=C:\Users\Usuario\OneDrive - MICE TRAVEL\EMPRESAS\Novo Nordisk
EVENT_ROUTES_PATH=C:\Users\Usuario\OneDrive - MICE TRAVEL\EMPRESAS\Novo Nordisk\event_routes.json
PROCESSED_IDS_PATH=C:\Users\Usuario\OneDrive - MICE TRAVEL\EMPRESAS\Novo Nordisk\_bot_processed_message_ids.json
NN_TEMPLATE_PATH=C:\Users\Usuario\OneDrive - MICE TRAVEL\Plantillas\LISTADO PARA VOLCAR LOS DATOS NN.xlsx
```

## Excel por evento

Cuando llega un evento nuevo, el bot agrupa por `NOMBRE EVENTO/CONGRESO/CURSO` y crea un Excel sugerido automaticamente. Ademas lo muestra en el panel como `Eventos nuevos`, para confirmar con el boton `Elegir Excel` si ese evento debe ir a ese archivo u otro.

Formato del nombre:

```text
NO ENVIAR -----LISTADO NOMBRECONGRESO FECHAINICIOCONGRESO.xlsx
```

Ejemplo:

```text
NO ENVIAR -----LISTADO ESC MUNICH 28 AGO.xlsx
```

## Comportamiento

- Lee correos no leidos con asunto `NUEVA SOLICITUD`.
- Extrae datos del email.
- Marca el correo como leido despues de importarlo.
- Genera Excel global, CSV y JSON.
- Genera Excel por evento.
- Permite elegir desde el panel el Excel global, la plantilla y el Excel asignado a cada evento nuevo.
- Si falta informacion, marca la solicitud como `pendiente_revision`.
- Si hay fallo tecnico, envia aviso por email si SMTP esta configurado.
- Guarda `Message-ID` en `_bot_processed_message_ids.json` para reducir duplicados.
- Antes de arrancar, busca actualizaciones en GitHub.

## Prueba antes del cliente

1. Ejecutar instalador.
2. Rellenar `config/.env`.
3. Ejecutar diagnostico.
4. Si todo sale OK, abrir panel.
5. Enviar correos de prueba.
6. Pulsar ON.
7. Verificar Excel global y Excel por evento.
8. Probar OneDrive con Excel abierto en dos equipos.

## GitHub

Repositorio:

```text
https://github.com/raulmm78/mice-travel-bot
```

No subir nunca:

- `config/.env`
- Excels generados
- CSV/JSON/logs
- correos reales
