# Datos a pedir al cliente antes de instalar

## 1. Ordenador principal

Pedir:

- Que indiquen que PC sera el principal.
- Que ese PC tenga OneDrive sincronizado.
- Que ese PC sea el unico con el bot en `ON` al principio.

Donde se mete:

No va en `config/.env`; es una decision operativa.

## 2. Carpeta OneDrive de prueba o produccion

Pedir:

- Ruta local exacta de la carpeta OneDrive donde se crearan los Excel por evento.

Ejemplo Windows:

```text
C:\Users\Usuario\OneDrive - MICE TRAVEL\EMPRESAS\Novo Nordisk
```

Donde se mete en `config/.env`:

```text
EVENT_OUTPUT_DIR=C:\Users\Usuario\OneDrive - MICE TRAVEL\EMPRESAS\Novo Nordisk
PROCESSED_IDS_PATH=C:\Users\Usuario\OneDrive - MICE TRAVEL\EMPRESAS\Novo Nordisk\_bot_processed_message_ids.json
```

## 3. Carpeta para Excel global/logs

Pedir:

- Si quieren una carpeta para consulta global.
- Puede estar dentro de OneDrive o fuera.

Donde se mete:

```text
OUTPUT_DIR=C:\Users\Usuario\OneDrive - MICE TRAVEL\EMPRESAS\Novo Nordisk\_global
```

## 4. Plantilla Excel oficial

Pedir:

- Archivo Excel plantilla definitivo.
- Confirmar hoja donde se vuelca, normalmente `Totales`.
- Confirmar que no se han cambiado columnas.

Donde se mete:

```text
NN_TEMPLATE_PATH=C:\ruta\a\LISTADO PARA VOLCAR LOS DATOS NN.xlsx
```

## 5. Buzon que recibe solicitudes

Pedir:

- Email del buzon.
- Servidor IMAP.
- Puerto IMAP.
- Usuario.
- Password o password de aplicacion.
- Carpeta a revisar, normalmente `INBOX`.

Donde se mete:

```text
IMAP_HOST=
IMAP_PORT=993
IMAP_USER=
IMAP_PASSWORD=
IMAP_FOLDER=INBOX
```

## 6. Correo para avisos

Pedir:

- A que persona o grupo debe avisar el bot si hay un problema.

Donde se mete:

```text
ALERT_EMAIL_TO=
```

## 7. SMTP para enviar avisos

Pedir:

- Servidor SMTP.
- Puerto SMTP.
- Usuario.
- Password o password de aplicacion.
- Remitente.

Donde se mete:

```text
SMTP_HOST=
SMTP_PORT=587
SMTP_USER=
SMTP_PASSWORD=
MAIL_FROM=
MAIL_TO=
```

`MAIL_TO` se usa para demos; `ALERT_EMAIL_TO` para avisos reales.

## 8. OpenAI API

Pedir o decidir:

- Si usan tu API key o una API key del cliente.
- Si se cobra cuota mensual por uso/mantenimiento.

Donde se mete:

```text
OPENAI_API_KEY=
OPENAI_MODEL=gpt-4.1-mini
```

## 9. Frecuencia de revision

Pedir:

- Cada cuantos minutos quieren revisar correo.

Recomendado inicial:

```text
WATCH_INTERVAL_SECONDS=300
```

300 segundos son 5 minutos.

## 10. Nombre de Excel por evento

Confirmar formato:

```text
NO ENVIAR -----LISTADO NOMBRECONGRESO FECHAINICIOCONGRESO.xlsx
```

Ejemplo:

```text
NO ENVIAR -----LISTADO ESC MUNICH 28 AGO.xlsx
```

Donde se mete:

No va en `config/.env`; esta implementado en el bot.

## 11. Prueba de Excel abierto

Pedir:

- Que haya dos PCs con OneDrive.
- Que abran el mismo Excel.
- Que uno edite una celda y guarde.
- Que luego el bot intente escribir.

Resultado esperado:

- Si sincroniza bien, seguimos con version local OneDrive.
- Si bloquea o genera conflicto, fase 2 con Microsoft 365/SharePoint API.
