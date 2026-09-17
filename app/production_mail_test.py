from __future__ import annotations

import importlib.util
import tempfile
from pathlib import Path
from unittest.mock import MagicMock, patch


SCRIPT = Path(__file__).resolve().parents[1] / "platform" / "windows" / "v9" / "probar_correo_produccion.py"
spec = importlib.util.spec_from_file_location("production_mail_check", SCRIPT)
assert spec and spec.loader
checker = importlib.util.module_from_spec(spec)
spec.loader.exec_module(checker)


def main() -> None:
    with tempfile.TemporaryDirectory() as temp:
        config = Path(temp) / ".env"
        config.write_text(
            "IMAP_PASSWORD=test-secret\n"
            "PROD_IMAP_HOST=imap.example.com\n"
            "PROD_IMAP_PORT=993\n"
            "PROD_IMAP_USER=production@example.com\n"
            "PROD_IMAP_PASSWORD=production-secret\n"
            "PROD_IMAP_FOLDER=INBOX\n",
            encoding="utf-8",
        )
        before = config.read_bytes()
        mailbox = MagicMock()
        mailbox.__enter__.return_value = mailbox
        mailbox.select.return_value = ("OK", [b"0"])
        with patch.object(checker.imaplib, "IMAP4_SSL", return_value=mailbox) as connect:
            assert checker.check_mail(config) == "production@example.com"
            connect.assert_called_once_with("imap.example.com", 993, timeout=10)
            mailbox.login.assert_called_once_with("production@example.com", "production-secret")
            mailbox.select.assert_called_once_with("INBOX", readonly=True)
            mailbox.uid.assert_not_called()
            mailbox.fetch.assert_not_called()
            mailbox.store.assert_not_called()
        assert config.read_bytes() == before
        assert list(Path(temp).iterdir()) == [config]
    print("OK: Solo lectura IMAP; sin mensajes ni archivos modificados")


if __name__ == "__main__":
    main()
