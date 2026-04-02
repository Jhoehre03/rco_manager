import os
import sys
import zipfile
import threading
import requests
from ui.version import VERSION, GITHUB_REPO

_API_URL = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"


def verificar_atualizacao():
    try:
        r = requests.get(_API_URL, timeout=5, headers={"Accept": "application/vnd.github+json"})
        r.raise_for_status()
        data = r.json()

        tag = data.get("tag_name", "").lstrip("v")
        if not tag:
            return {"disponivel": False}

        if _versao_maior(tag, VERSION):
            # Pega URL do primeiro asset .zip
            url_download = ""
            for asset in data.get("assets", []):
                if asset.get("name", "").endswith(".zip"):
                    url_download = asset["browser_download_url"]
                    break

            return {
                "disponivel": True,
                "versao": tag,
                "url_download": url_download,
                "descricao": data.get("body", ""),
            }

        return {"disponivel": False}

    except Exception as e:
        return {"disponivel": False, "erro": str(e)}


def baixar_e_instalar(url_download, callback_progresso=None):
    """
    Baixa o .zip, descompacta sobre a pasta atual do executável e
    cria/executa um update.bat que substitui os arquivos e reinicia.
    """
    try:
        pasta_update = os.path.join(
            os.environ.get("LOCALAPPDATA", os.path.expanduser("~")),
            "RCOManager", "update"
        )
        os.makedirs(pasta_update, exist_ok=True)
        zip_path = os.path.join(pasta_update, "rco_update.zip")

        # Download com progresso
        r = requests.get(url_download, stream=True, timeout=60)
        r.raise_for_status()
        total = int(r.headers.get("content-length", 0))
        baixado = 0

        with open(zip_path, "wb") as f:
            for chunk in r.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
                    baixado += len(chunk)
                    if callback_progresso and total:
                        callback_progresso(int(baixado / total * 100))

        if callback_progresso:
            callback_progresso(100)

        # Pasta de destino = pasta do executável (ou cwd em dev)
        if getattr(sys, "frozen", False):
            pasta_destino = os.path.dirname(sys.executable)
        else:
            pasta_destino = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

        # Descompacta e detecta subpasta raiz no zip
        with zipfile.ZipFile(zip_path, "r") as zf:
            zf.extractall(pasta_update)

            # Descobre se o zip tem uma subpasta raiz (ex: "RCO Manager/")
            nomes = zf.namelist()
            raiz = nomes[0].split("/")[0] if nomes else ""
            pasta_extraida = os.path.join(pasta_update, raiz) if raiz else pasta_update
            # Confirma que é realmente uma subpasta (não um arquivo)
            if not os.path.isdir(pasta_extraida):
                pasta_extraida = pasta_update

        # Cria update.bat
        bat_path = os.path.join(pasta_update, "update.bat")
        exe_path = sys.executable if getattr(sys, "frozen", False) else ""

        log_path = os.path.join(pasta_update, "update_log.txt")
        conteudo_bat = f"""@echo off
echo Aguardando encerramento do app... > "{log_path}"
timeout /t 3 /nobreak >nul
echo Copiando arquivos de: {pasta_extraida} >> "{log_path}"
echo Para: {pasta_destino} >> "{log_path}"
xcopy /E /Y /I "{pasta_extraida}\\*" "{pasta_destino}\\" >> "{log_path}" 2>&1
echo Xcopy retornou: %ERRORLEVEL% >> "{log_path}"
del /Q "{zip_path}" 2>nul
"""
        if exe_path:
            conteudo_bat += f'echo Reiniciando app... >> "{log_path}"\n'
            conteudo_bat += f'start "" "{exe_path}"\n'
        conteudo_bat += f'echo Concluido. >> "{log_path}"\n'
        conteudo_bat += "del /Q \"%~f0\"\n"

        with open(bat_path, "w", encoding="utf-8") as f:
            f.write(conteudo_bat)

        # Executa o bat e encerra o app atual
        import subprocess as _sp
        _sp.Popen(["cmd", "/c", bat_path], creationflags=0x08000000)  # CREATE_NO_WINDOW
        sys.exit(0)

    except Exception as e:
        raise RuntimeError(f"Falha na atualização: {e}")


def _versao_maior(nova, atual):
    """Retorna True se nova > atual (comparação semântica simples)."""
    def partes(v):
        try:
            return tuple(int(x) for x in v.split("."))
        except Exception:
            return (0,)
    return partes(nova) > partes(atual)


def verificar_em_background(callback):
    """Verifica atualização após 3 segundos em thread separada."""
    def _run():
        import time
        time.sleep(3)
        resultado = verificar_atualizacao()
        callback(resultado)

    t = threading.Thread(target=_run, daemon=True)
    t.start()
