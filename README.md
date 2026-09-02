# Prototipo de agente para emails de viajes

Este prototipo lee emails guardados como `.txt`, extrae datos de viaje y genera tres salidas para revisión operativa:

- Excel: `../../outputs/viajes_demo.xlsx`
- CSV: `../../outputs/viajes_demo.csv`
- JSON: `../../outputs/viajes_demo.json`
- Excel por evento: `../../outputs/eventos/`

Si existe `NN_TEMPLATE_PATH`, el Excel se genera usando la plantilla Novo Nordisk/MICE y vuelca los datos en la hoja `Totales`.
El Excel global se mantiene como consulta y, además, se genera un Excel independiente por evento.

## Campos extraídos

`email_id`, `nombre`, `apellidos`, `dni`, `email`, `telefono`, `evento`, `origen`, `destino`, `fecha_viaje`, `hora_preferida`, `preferencia`, `observaciones`, `estado`, `dudas`.

## Validación

Los campos obligatorios son:

- `nombre`
- `dni`
- `origen`
- `destino`
- `fecha_viaje`

Si falta alguno, si el DNI tiene formato dudoso o si la fecha no está normalizada como `YYYY-MM-DD`, la fila queda con `estado = pendiente_revision`.

## Uso sin API key

Ejecuta:

```bash
cd /Users/raulmartinez/Documents/Codex/2026-07-27/si/work/email-agent-demo
python3 process_emails.py
```

Sin `OPENAI_API_KEY`, el script usa reglas locales simples como fallback.

## Uso con OpenAI API

La opción recomendada es guardar la clave en un archivo local `.env`.

Copia la plantilla:

```bash
cp .env.example .env
```

Edita `.env` y pega tu clave real:

```text
OPENAI_API_KEY=tu_api_key_real
OPENAI_MODEL=gpt-4.1-mini
EMAIL_AGENT_PORT=8765
AUTO_UPDATE_ENABLED=1
AUTO_UPDATE_BRANCH=main
NN_TEMPLATE_PATH=/ruta/a/LISTADO PARA VOLCAR LOS DATOS NN.xlsx
EVENT_OUTPUT_DIR=/ruta/a/OneDrive/EMPRESAS/Novo Nordisk
PROCESSED_IDS_PATH=/ruta/a/OneDrive/EMPRESAS/Novo Nordisk/_bot_processed_message_ids.json
WATCH_INTERVAL_SECONDS=300
```

Para la primera fase con OneDrive no hace falta que el bot tenga claves de OneDrive: escribe en la carpeta local sincronizada y OneDrive sube los cambios. Los listados por evento se crean con este formato:

```text
NO ENVIAR ---- NOMBRECONGRESO FECHAINICIOCONGRESO.xlsx
```

Ejemplo:

```text
NO ENVIAR ---- CONGRESO IMS 28-09-26.xlsx
```

El bot guarda los identificadores reales de los emails procesados en `_bot_processed_message_ids.json`. Si varios equipos comparten la misma carpeta de OneDrive, todos pueden consultar el mismo registro para reducir duplicados.

Después ejecuta el script normalmente:

```bash
python3 process_emails.py
```

### Actualizaciones desde GitHub

Si la carpeta instalada viene de un repositorio GitHub privado, el lanzador ejecuta `update_before_start.py` antes de abrir el panel. Si hay una version nueva en la rama configurada, la descarga y despues arranca el bot normalmente. Si no hay internet, no hay cambios o GitHub falla, arranca la version local.

El archivo `.env` no se sube a GitHub y se conserva en cada ordenador.

Variables opcionales:

```text
AUTO_UPDATE_ENABLED=1
AUTO_UPDATE_BRANCH=main
```

También puedes seguir usando variable de entorno si lo prefieres:

```bash
export OPENAI_API_KEY="tu_api_key_real"
python3 process_emails.py
```

Con API key, el script envía cada email a OpenAI con un prompt de extracción en español y pide una respuesta JSON con esquema fijo. La validación local se aplica igualmente después de la extracción.

## Demo en Zoom con email nuevo

La forma más cómoda para presentar por Zoom es abrir el panel local. En Mac puedes hacer doble clic en:

```text
abrir_panel_mac.command
```

O abrirlo desde terminal:

```bash
cd /Users/raulmartinez/Documents/Codex/2026-07-27/si/work/email-agent-demo
python3 process_emails.py --dashboard
```

Después abre:

```text
http://127.0.0.1:8765
```

En el panel puedes:

- Enviar manualmente 3 correos de prueba desde cuentas distintas al buzón configurado.
- Dejar esos correos como no leídos en Gmail/Outlook.
- Pulsar `Probar conexión` para verificar IMAP.
- Pulsar `Detectar no leídos` para contar correos nuevos sin marcarlos como leídos.
- Pulsar `Procesar al Excel` para importar los correos, extraer los datos, actualizar Excel/CSV/JSON y marcarlos como leídos.
- Pulsar `Abrir Excel` para abrir el listado generado.
- Ver la actividad en el panel de logs.

El botón `ON` ejecuta una revisión inmediata y deja el bot activo. Por defecto revisa el buzón cada 5 minutos (`WATCH_INTERVAL_SECONDS=300`).

Algunos emails generados están incompletos a propósito para enseñar `pendiente_revision`.

### Configurar Gmail para la demo real

Usa una cuenta de demo, no una cuenta personal sensible.

En Gmail:

1. Activa la verificación en dos pasos.
2. Crea una contraseña de aplicación.
3. Activa IMAP en la configuración de Gmail.
4. Guarda la configuración en `.env`.

Ejemplo:

```text
OPENAI_API_KEY=tu_api_key_real
OPENAI_MODEL=gpt-4.1-mini
EMAIL_AGENT_PORT=8765

SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_USER=cuenta.demo@gmail.com
SMTP_PASSWORD=password_de_aplicacion
MAIL_FROM=cuenta.demo@gmail.com
MAIL_TO=cuenta.demo@gmail.com
ALERT_EMAIL_FROM=cuenta.demo@gmail.com
ALERT_EMAIL_TO=persona_que_recibe_avisos@micetravel.es

IMAP_HOST=imap.gmail.com
IMAP_PORT=993
IMAP_USER=cuenta.demo@gmail.com
IMAP_PASSWORD=password_de_aplicacion
IMAP_FOLDER=INBOX
```

Con esta configuración, la demo queda separada en dos fases visibles: detectar correos no leídos y volcar al Excel.
Si el bot encuentra un problema leyendo el correo, procesando los datos o escribiendo el Excel, enviará un aviso a `ALERT_EMAIL_TO` usando la misma configuración SMTP.

### Formato MICE/Novo Nordisk

- Los textos se vuelcan en MAYUSCULAS y sin tildes.
- Los emails se mantienen en minusculas.
- Las filas añadidas por el bot se escriben en azul para diferenciarlas de las ya revisadas manualmente.
- Las restricciones alimentarias reales quedan en negrita.
- `FECHA INICIO SERVICIO`, `IN` y `OUT` se guardan como fechas con formato `dd/mm/yy`.
- `MES SOLICITUD/FACTURACION` se guarda con formato `mm/yyyy`.
- La columna `GERENTE` se rellena con las iniciales detectadas en el `CC` del email.

También puedes enseñar que el sistema reacciona cuando "llega" un email nuevo usando el modo escucha:

```bash
cd /Users/raulmartinez/Documents/Codex/2026-07-27/si/work/email-agent-demo
python3 process_emails.py --watch
```

Con el programa abierto, copia un email de demo a la carpeta `emails/`:

```bash
cp demo_email_ok.txt emails/email_003.txt
```

El programa detectará el nuevo `.txt` y regenerará automáticamente:

- `../../outputs/viajes_demo.xlsx`
- `../../outputs/viajes_demo.csv`
- `../../outputs/viajes_demo.json`

También puedes enseñar un caso incompleto:

```bash
cp demo_email_pendiente.txt emails/email_004.txt
```

Ese email debería quedar como `pendiente_revision` porque no incluye todos los campos obligatorios.

## Preparación para Gmail, Outlook o n8n

Para una versión conectada:

1. Gmail, Outlook o n8n detecta un email nuevo.
2. El cuerpo del email se guarda como texto o se envía al extractor.
3. El extractor devuelve JSON estructurado.
4. La validación marca la fila como `ok` o `pendiente_revision`.
5. n8n, Make, Google Sheets o Excel/SharePoint añade la fila al fichero final.

Recomendación de fase 1: automatizar extracción y revisión, no compra de billetes. La compra automática debería tratarse como fase posterior por pagos, permisos y responsabilidad operativa.
