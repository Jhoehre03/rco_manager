"""
Verifica se o .NET 6.0 Desktop Runtime está instalado.
Se não estiver, baixa e instala silenciosamente.
"""
import os
import subprocess
import sys
import tempfile
import urllib.request


_DOTNET_URL = (
    "https://download.visualstudio.microsoft.com/download/pr/"
    "8a1e6a00-b3cc-4f79-b5b2-edcd96f48e17/90f5c7f3b2bdc0af8c24f1aa89e5f3de/"
    "windowsdesktop-runtime-6.0.36-win-x64.exe"
)


def _dotnet6_instalado():
    """Retorna True se o .NET 6.0 Desktop Runtime estiver instalado."""
    try:
        result = subprocess.run(
            ["dotnet", "--list-runtimes"],
            capture_output=True, text=True, timeout=5
        )
        return "Microsoft.NETCore.App 6." in result.stdout or \
               "Microsoft.WindowsDesktop.App 6." in result.stdout
    except Exception:
        pass

    # Fallback: verifica no registro do Windows
    try:
        import winreg
        key_path = r"SOFTWARE\dotnet\Setup\InstalledVersions\x64\sharedfx\Microsoft.WindowsDesktop.App"
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
            i = 0
            while True:
                try:
                    name, _, _ = winreg.EnumValue(key, i)
                    if name.startswith("6."):
                        return True
                    i += 1
                except OSError:
                    break
    except Exception:
        pass

    return False


def garantir_dotnet():
    """
    Verifica e instala o .NET 6.0 Desktop Runtime se necessário.
    Exibe uma janela de progresso simples via tkinter (disponível no Windows).
    Retorna True se está disponível, False se falhou.
    """
    if _dotnet6_instalado():
        return True

    # Mostra diálogo de confirmação
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        resposta = messagebox.askyesno(
            "RCO Manager — Requisito necessário",
            "O Microsoft .NET 6.0 Runtime não está instalado.\n\n"
            "Este componente é necessário para o RCO Manager funcionar.\n\n"
            "Deseja instalar agora? (requer conexão com a internet)\n"
            "Tamanho: ~55 MB",
            icon='warning'
        )
        root.destroy()

        if not resposta:
            return False

        # Mostra janela de progresso
        root = tk.Tk()
        root.title("RCO Manager — Instalando .NET 6.0")
        root.geometry("380x100")
        root.resizable(False, False)
        root.eval('tk::PlaceWindow . center')

        label = tk.Label(root, text="Baixando .NET 6.0 Runtime...", pady=10)
        label.pack()

        progress_var = tk.StringVar(value="0%")
        progress_label = tk.Label(root, textvariable=progress_var, fg="#3b82f6")
        progress_label.pack()

        root.update()

        # Baixa o instalador
        tmp = tempfile.mktemp(suffix=".exe")

        def reporthook(count, block_size, total_size):
            if total_size > 0:
                pct = int(count * block_size * 100 / total_size)
                progress_var.set(f"{min(pct, 100)}%")
                root.update()

        urllib.request.urlretrieve(_DOTNET_URL, tmp, reporthook)

        label.config(text="Instalando .NET 6.0 Runtime...")
        progress_var.set("Aguarde...")
        root.update()

        # Instala silenciosamente
        subprocess.run([tmp, "/install", "/quiet", "/norestart"], check=True)

        os.remove(tmp)
        root.destroy()

        return True

    except Exception as e:
        try:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror(
                "RCO Manager — Erro",
                f"Falha ao instalar o .NET 6.0:\n{e}\n\n"
                "Instale manualmente em:\nhttps://dotnet.microsoft.com/download/dotnet/6.0"
            )
            root.destroy()
        except Exception:
            pass
        return False
