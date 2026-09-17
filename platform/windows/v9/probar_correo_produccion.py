from __future__ import annotations

import imaplib
import sys
from pathlib import Path


def production_settings(path: Path) -> dict[str, str]:
    if not path.is_file():
        raise ValueError("No se encuentra config/.env de la instalacion")
    values: dict[str, str] = {}
    for line in path.read_text(encoding="utf-8-sig").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        name, value = line.split("=", 1)
        if name.strip().startswith("PROD_IMAP_"):
            values[name.strip()] = value.strip().strip('"').strip("'")
    missing = [name for name in ("HOST", "USER", "PASSWORD") if not values.get(f"PROD_IMAP_{name}")]
    if missing:
        raise ValueError("Faltan en config/.env: " + ", ".join(f"PROD_IMAP_{name}" for name in missing))
    return values


def check_mail(path: Path) -> str:
    settings = production_settings(path)
    host = settings["PROD_IMAP_HOST"]
    port = int(settings.get("PROD_IMAP_PORT") or "993")
    user = settings["PROD_IMAP_USER"]
    password = settings["PROD_IMAP_PASSWORD"]
    folder = settings.get("PROD_IMAP_FOLDER") or "INBOX"
    mailbox = imaplib.IMAP4_SSL(host, port, timeout=10)
    with mailbox:
        mailbox.login(user, password)
        status, _ = mailbox.select(folder, readonly=True)
        if status != "OK":
            raise RuntimeError("No se pudo abrir la carpeta en modo lectura")
    return user


def main() -> int:
    if len(sys.argv) != 2:
        print("Falta la ruta de config/.env")
        return 1
    try:
        user = check_mail(Path(sys.argv[1]))
    except ValueError as exc:
        print(f"ERROR: {exc}")
        return 1
    except (OSError, imaplib.IMAP4.error, RuntimeError) as exc:
        print(f"ERROR: No se pudo conectar ({type(exc).__name__}). Comprueba servidor, clave e IMAP.")
        return 1
    print(f"OK: Conexion IMAP de produccion verificada para {user}.")
    print("No se han descargado ni marcado mensajes. No se ha abierto ningun Excel.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
