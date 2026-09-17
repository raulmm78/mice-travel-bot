from __future__ import annotations

import io
import json
import tempfile
import zipfile
from pathlib import Path
from unittest.mock import patch

import update_before_start as updater


def main() -> None:
    with tempfile.TemporaryDirectory() as temp:
        root = Path(temp)
        (root / "app").mkdir()
        (root / "config").mkdir()
        (root / "data").mkdir()
        (root / "app" / "process_emails.py").write_text("version antigua", encoding="utf-8")
        (root / "config" / ".env").write_text("CLAVE_PRIVADA=sin_cambios\n", encoding="utf-8")
        (root / "data" / "listado.xlsx").write_bytes(b"excel-sin-cambios")

        archive_data = io.BytesIO()
        with zipfile.ZipFile(archive_data, "w") as archive:
            prefix = "mice-travel-bot-version/"
            archive.writestr(prefix + "app/process_emails.py", "version nueva")
            archive.writestr(prefix + "app/update_before_start.py", "actualizador nuevo")
            archive.writestr(prefix + "app/assets/plantilla_nn_2026_09_17.xlsx", b"plantilla-convertida")
            archive.writestr(prefix + "config/.env", "CLAVE_PRIVADA=robada\n")
            archive.writestr(prefix + "data/listado.xlsx", b"excel-borrado")

        sha = "a" * 40
        def fake_download(url: str) -> bytes:
            if "api.github.com" in url:
                return json.dumps({"sha": sha}).encode("utf-8")
            return archive_data.getvalue()

        with patch.object(updater, "PROJECT_DIR", root), \
             patch.object(updater, "VERSION_FILE", root / "config" / ".github_commit"), \
             patch.object(updater, "github_bytes", side_effect=fake_download), \
             patch.object(updater, "env_value", return_value="main"):
            updater.update_zip_installation()
            updater.update_zip_installation()

        assert (root / "app" / "process_emails.py").read_text(encoding="utf-8") == "version nueva"
        assert not (root / "app" / "assets" / "plantilla_nn_2026_09_17.xlsx").exists()
        assert (root / "config" / ".env").read_text(encoding="utf-8") == "CLAVE_PRIVADA=sin_cambios\n"
        assert (root / "data" / "listado.xlsx").read_bytes() == b"excel-sin-cambios"
        assert (root / "config" / ".github_commit").read_text(encoding="ascii").strip() == sha
    print("OK - Actualizacion GitHub sin Git; claves y Excel intactos")


if __name__ == "__main__":
    main()
