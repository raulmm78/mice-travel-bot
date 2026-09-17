# Guia rapida Windows - MICE Travel Bot

## Instalacion desde cero

Si usas el ZIP:

1. Boton derecho sobre el ZIP mas reciente `MICE Travel Bot_windows_cliente_TEST_BOT_*.zip`.
2. Pulsar `Extraer todo...`.
3. Entrar en la carpeta extraida `MICE Travel Bot`.
4. Ejecutar `instalar_windows.bat`.

No ejecutes `instalar_windows.bat` desde dentro del ZIP, porque Windows muestra los archivos como si estuvieran ahi pero el instalador no puede encontrar bien `app\requirements.txt`. Si el ZIP ya incluye `config/.env`, el instalador conserva esas claves y rutas.

## Activar actualizaciones en un PC ya instalado desde ZIP antiguo

El ZIP entregado antes de esta mejora no puede actualizarse solo porque su actualizador antiguo exigia Git. Una vez, extrae el ZIP `MICE_Travel_Bot_ACTUALIZAR_Windows_*.zip` y ejecuta `ACTUALIZAR_MICE_TRAVEL_BOT_windows.bat`. Cierra antes el bot con el acceso directo `Cerrar MICE Travel Bot`.

Alternativamente, en PowerShell del PC cliente:

```powershell
$script = Join-Path $env:TEMP "activar_actualizaciones_windows.ps1"
Invoke-WebRequest -UseBasicParsing "https://raw.githubusercontent.com/raulmm78/mice-travel-bot/main/activar_actualizaciones_windows.ps1" -OutFile $script
powershell -NoProfile -ExecutionPolicy Bypass -File $script
```

El script localiza la instalacion mediante el acceso directo del escritorio, activa el nuevo actualizador y descarga inmediatamente el codigo mas reciente desde GitHub. Conserva `config/.env`, correos, registros y Excels. Desde ese momento no hace falta enviar mas ZIPs para actualizar el programa.

## Instalacion desde GitHub

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
EVENT_TEMPLATE_PATH=
WATCH_INTERVAL_SECONDS=300
```

## Arranque

Usar el acceso directo:

```text
MICE Travel Bot
```

Al arrancar:

1. Busca actualizaciones en GitHub al abrir la app, sin necesitar Git en instalaciones desde ZIP.
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

## Test de seguridad del Excel

Antes de apuntar a un Excel real, puedes ejecutar:

```text
test_seguridad_excel_windows.bat
```

Debe mostrar:

```text
OK - Seguridad Excel verificada
```

Comprueba que:

- Se filtran correos que no son formulario.
- Se aceptan reenvios `RV:` o `FW:` si contienen el formulario.
- No se borran filas antiguas del Excel.
- No se duplica la misma solicitud al repetir el proceso.
- Los Excel existentes, tanto el global como los de evento, conservan sus filas.
- Los eventos nuevos esperan a que pulses `Elegir Excel` o `Crear Excel` en el panel.
- `Crear Excel` usa por defecto la plantilla NN del 17 de septiembre instalada localmente con el ZIP privado `MICE_Travel_Bot_PLANTILLA_NN_Windows_*.zip`. Es una copia `.xlsx` convertida del `.xls` facilitado por el cliente; el original de OneDrive no se modifica ni se publica en GitHub. Extrae el ZIP y ejecuta `INSTALAR_PLANTILLA_NN_windows.bat` una sola vez. `EVENT_TEMPLATE_PATH` permite seleccionar otra plantilla `.xlsx` mas adelante.
- Antes de actualizar un Excel existente, el bot guarda una copia en `_bot_backups` junto al archivo.
- Si el Excel cambia durante el proceso, el bot detiene ese guardado y deja el archivo original intacto.

## Recomendacion para varios ordenadores

Primera version:

- Solo un PC con el bot en `ON`.
- Todos pueden abrir/ver los Excel desde OneDrive.
- El registro `_bot_processed_message_ids.json` debe estar en la carpeta compartida de OneDrive.

Si varios PCs tienen el bot en `ON`, hay menos riesgo de duplicados por el `Message-ID`, pero sigue pudiendo haber conflicto si OneDrive tarda en sincronizar.
