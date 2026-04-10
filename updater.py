import sys
import os
import json
import ssl
import base64
import subprocess
import urllib.request
import urllib.error
import certifi


GITHUB_REPO = "Victor-D-T/comparativo_extratos"
API_URL = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"


def _parse_version(v):
    return tuple(int(x) for x in v.strip("v").split("."))


def check_for_update(current_version):
    """Verifica se há versão nova. Retorna (tag, download_url) ou None se já atualizado.
    Lança Exception com mensagem clara em caso de problema."""
    if not getattr(sys, 'frozen', False):
        return None

    ctx = ssl.create_default_context(cafile=certifi.where())
    req = urllib.request.Request(API_URL, headers={"User-Agent": "comparativo-extratos"})
    try:
        with urllib.request.urlopen(req, timeout=8, context=ctx) as response:
            data = json.loads(response.read())
    except urllib.error.HTTPError as e:
        if e.code == 404:
            raise Exception("Nenhum release estável encontrado no GitHub (404). O release pode estar marcado como 'pre-release'.")
        raise Exception(f"Erro na API do GitHub: HTTP {e.code}")

    latest_tag = data.get("tag_name", "")
    if not latest_tag:
        raise Exception("GitHub não retornou nenhuma versão de release.")

    latest = _parse_version(latest_tag)
    current = _parse_version(current_version)
    if latest <= current:
        return None  # já na versão mais recente

    exe_asset = next(
        (a for a in data.get("assets", []) if a["name"].endswith(".exe")),
        None
    )
    if exe_asset is None:
        raise Exception(f"Versão {latest_tag} disponível, mas o executável ainda não foi gerado. Tente novamente em alguns minutos.")

    return latest_tag, exe_asset["browser_download_url"]


def download_and_apply(download_url, tag):
    """Baixa o novo exe e cria o script de substituição. Encerra o processo."""
    current_exe = sys.executable
    new_exe_path = current_exe + ".new"

    ctx = ssl.create_default_context(cafile=certifi.where())
    req = urllib.request.Request(download_url, headers={"User-Agent": "comparativo-extratos"})
    try:
        with urllib.request.urlopen(req, context=ctx) as response:
            with open(new_exe_path, "wb") as f:
                f.write(response.read())
    except urllib.error.HTTPError as e:
        raise Exception(f"Erro ao baixar executável: HTTP {e.code}\nURL: {download_url}")

    # PowerShell com -EncodedCommand evita problemas com caminhos que contêm
    # parênteses (ex: "programa (1).exe") — o CMD falha nesse caso mesmo com aspas.
    ps_script = (
        f'Start-Sleep -Seconds 2; '
        f'Move-Item -Force "{new_exe_path}" "{current_exe}"; '
        f'Start-Process "{current_exe}"'
    )
    encoded_cmd = base64.b64encode(ps_script.encode("utf-16-le")).decode("ascii")

    subprocess.Popen(
        [
            "powershell",
            "-ExecutionPolicy", "Bypass",
            "-WindowStyle", "Hidden",
            "-EncodedCommand", encoded_cmd,
        ],
        creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP,
    )
    os._exit(0)
