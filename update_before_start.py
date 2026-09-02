from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent


def env_value(name: str, default: str = "") -> str:
    env_path = BASE_DIR / ".env"
    if env_path.exists():
        for line in env_path.read_text(encoding="utf-8").splitlines():
            clean = line.strip()
            if not clean or clean.startswith("#") or "=" not in clean:
                continue
            key, value = clean.split("=", 1)
            if key.strip() == name:
                return value.strip().strip('"').strip("'")
    return os.getenv(name, default).strip()


def run_git(args: list[str]) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        ["git", "-C", str(BASE_DIR), *args],
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        timeout=45,
        check=False,
    )


def main() -> int:
    if env_value("AUTO_UPDATE_ENABLED", "1").lower() in {"0", "false", "no", "off"}:
        print("Autoactualizacion desactivada.")
        return 0

    if not (BASE_DIR / ".git").exists():
        print("Autoactualizacion omitida: esta carpeta aun no es un repositorio Git.")
        return 0

    branch = env_value("AUTO_UPDATE_BRANCH", "")
    current_branch = run_git(["rev-parse", "--abbrev-ref", "HEAD"]).stdout.strip()
    target_branch = branch or current_branch or "main"

    print(f"Buscando actualizaciones en GitHub ({target_branch})...")
    fetch = run_git(["fetch", "--quiet", "origin", target_branch])
    if fetch.returncode != 0:
        print("No se pudo consultar GitHub. Arrancando version local.")
        print(fetch.stdout.strip())
        return 0

    local = run_git(["rev-parse", "HEAD"]).stdout.strip()
    remote = run_git(["rev-parse", f"origin/{target_branch}"]).stdout.strip()
    if not local or not remote or local == remote:
        print("No hay actualizaciones. Arrancando version local.")
        return 0

    status = run_git(["status", "--porcelain"]).stdout.strip()
    if status:
        print("Hay cambios locales en el codigo. No se aplica la actualizacion automatica.")
        return 0

    pull = run_git(["pull", "--ff-only", "origin", target_branch])
    if pull.returncode == 0:
        print("Actualizacion descargada correctamente.")
    else:
        print("No se pudo aplicar la actualizacion. Arrancando version local.")
        print(pull.stdout.strip())
    return 0


if __name__ == "__main__":
    sys.exit(main())
