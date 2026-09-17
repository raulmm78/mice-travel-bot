from __future__ import annotations

import io
import json
import os
import subprocess
import sys
import tempfile
import urllib.request
import zipfile
from pathlib import Path


APP_DIR = Path(__file__).resolve().parent
PROJECT_DIR = APP_DIR.parent
REPO = "raulmm78/mice-travel-bot"
VERSION_FILE = PROJECT_DIR / "config" / ".github_commit"


def env_value(name: str, default: str = "") -> str:
    env_path = PROJECT_DIR / "config" / ".env"
    if not env_path.exists():
        env_path = PROJECT_DIR / ".env"
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
        ["git", "-C", str(PROJECT_DIR), *args],
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        timeout=45,
        check=False,
    )


def update_git_installation() -> None:
    branch = env_value("AUTO_UPDATE_BRANCH", "")
    current_branch = run_git(["rev-parse", "--abbrev-ref", "HEAD"]).stdout.strip()
    target_branch = branch or current_branch or "main"

    print(f"Buscando actualizaciones en GitHub ({target_branch})...")
    fetch = run_git(["fetch", "--quiet", "origin", target_branch])
    if fetch.returncode != 0:
        print("No se pudo consultar GitHub. Arrancando version local.")
        print(fetch.stdout.strip())
        return

    local = run_git(["rev-parse", "HEAD"]).stdout.strip()
    remote = run_git(["rev-parse", f"origin/{target_branch}"]).stdout.strip()
    if not local or not remote or local == remote:
        print("No hay actualizaciones. Arrancando version local.")
        return

    status = run_git(["status", "--porcelain"]).stdout.strip()
    if status:
        print("Hay cambios locales en el codigo. No se aplica la actualizacion automatica.")
        return

    pull = run_git(["pull", "--ff-only", "origin", target_branch])
    if pull.returncode == 0:
        print("Actualizacion descargada correctamente.")
    else:
        print("No se pudo aplicar la actualizacion. Arrancando version local.")
        print(pull.stdout.strip())


def github_bytes(url: str) -> bytes:
    request = urllib.request.Request(url, headers={"User-Agent": "MICE-Travel-Bot-Updater"})
    with urllib.request.urlopen(request, timeout=30) as response:
        return response.read()


def managed_file(relative: Path) -> bool:
    parts = relative.parts
    if len(parts) == 2 and parts[0] == "app":
        return relative.suffix in {".py", ".bat", ".command", ".txt"}
    if len(parts) == 3 and parts[:2] == ("app", "assets"):
        return relative.suffix.lower() in {".png", ".jpg", ".jpeg"}
    if len(parts) == 1:
        return relative.name in {
            "cerrar_panel_windows.bat", "diagnostico_windows.bat",
            "instalar_windows.bat", "test_seguridad_excel_windows.bat",
        }
    if len(parts) == 3 and parts[:2] == ("platform", "windows"):
        return relative.suffix == ".ps1"
    if len(parts) == 2 and parts[0] == "docs":
        return relative.suffix == ".md"
    return False


def update_zip_installation() -> None:
    branch = env_value("AUTO_UPDATE_BRANCH", "main") or "main"
    commit_info = json.loads(github_bytes(f"https://api.github.com/repos/{REPO}/commits/{branch}"))
    sha = str(commit_info["sha"])
    if VERSION_FILE.exists() and VERSION_FILE.read_text(encoding="ascii").strip() == sha:
        print("No hay actualizaciones. Arrancando version local.")
        return

    archive = zipfile.ZipFile(io.BytesIO(github_bytes(f"https://codeload.github.com/{REPO}/zip/{sha}")))
    updates: dict[Path, bytes] = {}
    for entry in archive.infolist():
        if entry.is_dir():
            continue
        parts = Path(entry.filename).parts
        if len(parts) < 2:
            continue
        relative = Path(*parts[1:])
        if managed_file(relative):
            updates[relative] = archive.read(entry)
    if Path("app/process_emails.py") not in updates or Path("app/update_before_start.py") not in updates:
        raise RuntimeError("La descarga de GitHub esta incompleta")

    requirements = updates.get(Path("app/requirements.txt"))
    current_requirements = PROJECT_DIR / "app" / "requirements.txt"
    if requirements is not None and (not current_requirements.exists() or current_requirements.read_bytes() != requirements):
        with tempfile.TemporaryDirectory(prefix="mice-requirements-") as temp:
            candidate = Path(temp) / "requirements.txt"
            candidate.write_bytes(requirements)
            subprocess.run([sys.executable, "-m", "pip", "install", "-r", str(candidate)], check=True, timeout=180)

    previous: dict[Path, bytes | None] = {}
    changed: list[Path] = []
    try:
        for relative, content in updates.items():
            target = PROJECT_DIR / relative
            before = target.read_bytes() if target.exists() else None
            if before == content:
                continue
            target.parent.mkdir(parents=True, exist_ok=True)
            candidate = target.with_name(f".{target.name}.update-{os.getpid()}")
            candidate.write_bytes(content)
            try:
                os.replace(candidate, target)
            finally:
                candidate.unlink(missing_ok=True)
            previous[target] = before
            changed.append(target)
        VERSION_FILE.parent.mkdir(parents=True, exist_ok=True)
        VERSION_FILE.write_text(sha + "\n", encoding="ascii")
    except Exception:
        for target in reversed(changed):
            original = previous[target]
            if original is None:
                target.unlink(missing_ok=True)
            else:
                target.write_bytes(original)
        raise
    print(f"Actualizacion de GitHub aplicada: {sha[:8]} ({len(changed)} archivos)")


def main() -> int:
    if env_value("AUTO_UPDATE_ENABLED", "1").lower() in {"0", "false", "no", "off"}:
        print("Autoactualizacion desactivada.")
        return 0

    try:
        if (PROJECT_DIR / ".git").exists():
            update_git_installation()
        else:
            update_zip_installation()
    except Exception as exc:
        print(f"No se pudo actualizar desde GitHub: {exc}. Arrancando version local.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
