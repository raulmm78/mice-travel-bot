# Checklist de entrega MICE Travel Bot

## Antes de ir al cliente

- Confirmar que el repositorio de GitHub esta accesible.
- Confirmar que `.env` no esta subido a GitHub.
- Llevar anotadas las claves de correo/OpenAI por un canal seguro.
- Llevar localizada la plantilla Excel oficial.
- Confirmar la carpeta local de OneDrive donde deben crearse los listados.

## En cada ordenador Windows

1. Instalar Python 3.
2. Instalar Git si se quiere autoactualizacion desde GitHub.
3. Iniciar sesion en OneDrive y esperar a que sincronice.
4. Descargar o clonar el proyecto.
5. Ejecutar `instalar_windows.bat`.
6. Configurar `.env`.
7. Ejecutar `abrir_panel_windows.bat`.
8. Comprobar luces verdes de IMAP y Excel.
9. Ejecutar `diagnostico_windows.bat` antes de hacer la prueba real.

## Variables minimas de `.env`

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

## Prueba con OneDrive

1. Abrir el Excel desde dos ordenadores.
2. Dejar el bot en `ON` solo en un ordenador.
3. Enviar 3 correos de prueba al buzon.
4. Verificar que el bot detecta los correos no leidos.
5. Verificar que los marca como leidos.
6. Verificar que crea o actualiza el Excel global.
7. Verificar que crea el Excel por evento con formato `NO ENVIAR ---- NOMBRECONGRESO FECHAINICIOCONGRESO.xlsx`.
8. Verificar que el segundo ordenador ve los cambios por OneDrive.
9. Editar manualmente alguna fila y comprobar si OneDrive genera conflicto.

## Prueba de Excel abierto

1. PC 1: dejar el bot en `OFF`.
2. PC 1 y PC 2: abrir el mismo Excel de evento desde OneDrive.
3. PC 2: escribir algo manual en una celda libre y guardar.
4. Confirmar que PC 1 ve el cambio sincronizado.
5. Enviar un correo nuevo de ese mismo evento.
6. PC 1: poner el bot en `ON`.
7. Comprobar una de estas salidas:
   - si el Excel se actualiza en ambos PCs, OneDrive permite el flujo local;
   - si Excel bloquea el archivo, debe llegar un email de aviso de error;
   - si OneDrive crea una copia en conflicto, hay que pasar a Microsoft 365/SharePoint API.

## Criterio para produccion

- Si OneDrive sincroniza bien con los Excel abiertos, se puede entregar como version local con OneDrive sincronizado.
- Si aparecen bloqueos o copias en conflicto, pasar a fase Microsoft 365/SharePoint API.
- Si varios ordenadores tienen que ejecutar el bot, mantener el registro `_bot_processed_message_ids.json` en una carpeta compartida de OneDrive.
