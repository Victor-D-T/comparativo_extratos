import sys
import os
import json
import ssl
import base64
import subprocess
import urllib.request
import urllib.error
import urllib.parse
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
    current_dir = os.path.dirname(current_exe)

    # Usa o nome canônico do asset (da URL do GitHub) como destino.
    # Assim o novo exe é instalado sem "(1)" no nome, mesmo que o exe atual tenha.
    asset_name = os.path.basename(urllib.parse.urlparse(download_url).path)
    target_exe = os.path.join(current_dir, asset_name)
    new_exe_path = target_exe + ".new"

    ctx = ssl.create_default_context(cafile=certifi.where())
    req = urllib.request.Request(download_url, headers={"User-Agent": "comparativo-extratos"})
    try:
        with urllib.request.urlopen(req, context=ctx) as response:
            with open(new_exe_path, "wb") as f:
                f.write(response.read())
    except urllib.error.HTTPError as e:
        raise Exception(f"Erro ao baixar executável: HTTP {e.code}\nURL: {download_url}")

    current_pid = os.getpid()
    ps_lines = [
        # Aguarda o processo pai morrer de fato antes de tentar mover o arquivo
        f'$p = Get-Process -Id {current_pid} -ErrorAction SilentlyContinue',
        f'if ($p) {{ $p.WaitForExit(15000) }}',
        # Remove Mark of the Web (bloqueio de arquivo baixado da internet)
        f'Unblock-File -Path "{new_exe_path}" -ErrorAction SilentlyContinue',
        # Tenta mover até 5 vezes (Defender pode travar brevemente o arquivo)
        f'$ok = $false',
        f'for ($i = 0; $i -lt 5; $i++) {{',
        f'    try {{',
        f'        Move-Item -Force "{new_exe_path}" "{target_exe}" -ErrorAction Stop',
        f'        $ok = $true; break',
        f'    }} catch {{ Start-Sleep -Seconds 1 }}',
        f'}}',
        # SÓ relança o programa se o move funcionou — evita loop com exe antigo
        f'if ($ok) {{ Start-Process "{target_exe}" }}',
    ]
    ps_script = "\n".join(ps_lines)
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
