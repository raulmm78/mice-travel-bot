# Guia Mac para ensayar antes del cliente

## Doble clic

Usar estos archivos en la raiz de la carpeta:

```text
instalar_mac.command
diagnostico_mac.command
abrir_panel_mac.command
cerrar_panel_mac.command
```

Tambien puedes usar la app del Escritorio:

```text
MICE Travel Bot.app
```

## Secuencia recomendada

1. Abrir `instalar_mac.command`.
2. Revisar `config/.env`.
3. Abrir `diagnostico_mac.command`.
4. Si el diagnostico sale OK, abrir `abrir_panel_mac.command` o `MICE Travel Bot.app`.
5. Enviar correos de prueba.
6. Pulsar `ON`.
7. Comprobar Excel global y Excel por evento.
8. Cerrar con `cerrar_panel_mac.command`.

## Si el diagnostico falla

- `OpenAI API key configurada`: falta `OPENAI_API_KEY`.
- `Configuracion IMAP`: faltan usuario o password de correo.
- `Conexion IMAP`: host, puerto, usuario, password o permisos IMAP no son correctos.
- `Plantilla Excel`: `NN_TEMPLATE_PATH` no apunta al Excel plantilla.
- `Carpeta global escribible`: `OUTPUT_DIR` no existe o no permite escribir.
- `Carpeta eventos escribible`: `EVENT_OUTPUT_DIR` no existe o no permite escribir.
- `Avisos por email configurados`: falta SMTP o `ALERT_EMAIL_TO`.

## Nota

La prueba Mac valida el funcionamiento del bot y la configuracion. La prueba que solo puede hacerse en el entorno del cliente es OneDrive con el Excel abierto en dos PCs.
