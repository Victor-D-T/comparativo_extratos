import sys
import os
import json
import subprocess
import urllib.request


GITHUB_REPO = "Victor-D-T/comparativo_extratos"
API_URL = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"


def _parse_version(v):
    return tuple(int(x) for x in v.strip("v").split("."))


def check_for_update(current_version):
    """Verifica se há versão nova. Retorna (tag, download_url) ou None."""
    if not getattr(sys, 'frozen', False):
        return None
    try:
        req = urllib.request.Request(API_URL, headers={"User-Agent": "comparativo-extratos"})
        with urllib.request.urlopen(req, timeout=8) as response:
            data = json.loads(response.read())

        latest_tag = data.get("tag_name", "")
        if not latest_tag:
            return None

        latest = _parse_version(latest_tag)
        current = _parse_version(current_version)
        if latest <= current:
            return None

        exe_asset = next(
            (a for a in data.get("assets", []) if a["name"].endswith(".exe")),
            None
        )
        if exe_asset is None:
            return None

        return latest_tag, exe_asset["browser_download_url"]

    except Exception:
        return None


def download_and_apply(download_url, tag):
    """Baixa o novo exe e cria o script de substituição. Encerra o processo."""
    current_exe = sys.executable
    new_exe_path = current_exe + ".new"

    urllib.request.urlretrieve(download_url, new_exe_path)

    bat_content = (
        "@echo off\n"
        "timeout /t 2 /nobreak >nul\n"
        f'move /y "{new_exe_path}" "{current_exe}"\n'
        f'start "" "{current_exe}"\n'
        'del "%~f0"\n'
    )
    bat_path = os.path.join(os.path.dirname(current_exe), "_update.bat")
    with open(bat_path, "w", encoding="utf-8") as f:
        f.write(bat_content)

    subprocess.Popen(["cmd", "/c", bat_path])
    sys.exit(0)
