import sys
import os

# Quando empacotado pelo PyInstaller, sys._MEIPASS aponta para a pasta temporária
# com os arquivos bundled. Precisamos mudar o diretório de trabalho para que
# database.py, token.json e dados.json sejam lidos/escritos na pasta do executável.
if getattr(sys, 'frozen', False):
    # Pasta onde o executável está (dist/RCO Manager/)
    BASE_EXEC = os.path.dirname(sys.executable)
    os.chdir(BASE_EXEC)

    # Cria dados.json a partir do template se não existir
    if not os.path.exists('dados.json'):
        template = os.path.join(sys._MEIPASS, 'dados_inicial.json')
        if os.path.exists(template):
            import shutil
            shutil.copy(template, 'dados.json')
        else:
            import json
            with open('dados.json', 'w', encoding='utf-8') as f:
                json.dump({"escolas": [], "ultima_atualizacao": None}, f,
                          ensure_ascii=False, indent=2)

from ui.dotnet_check import garantir_dotnet
from ui.app import iniciar

if __name__ == "__main__":
    if not garantir_dotnet():
        sys.exit(1)
    iniciar()
