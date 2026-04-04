# -*- mode: python ; coding: utf-8 -*-
import os
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

# Verifica se o ícone existe
icon_path = os.path.join('ui', 'icon.ico')
icon = icon_path if os.path.exists(icon_path) else None

# Coleta dados extras das bibliotecas que precisam deles
datas = [
    ('ui/index.html', 'ui'),
    ('ui/modal_lancar.js', 'ui'),
    ('ui/modal_editar.js', 'ui'),
]

# dados_inicial.json — template para novos usuários (main.py copia para dados.json se não existir)
if os.path.exists('dados_inicial.json'):
    datas.append(('dados_inicial.json', '.'))

# oauth_credentials.json é copiado pelo build.bat diretamente para dist/RCO Manager/
# ao lado do .exe, pois o app o lê do cwd (pasta do executável), não do _MEIPASS.

# Coleta arquivos de dados das bibliotecas que precisam deles
datas += collect_data_files('webview')
datas += collect_data_files('gspread')
datas += collect_data_files('google_auth_oauthlib')
datas += collect_data_files('selenium')

hidden_imports = collect_submodules('selenium') + [
    # pywebview internals (Windows — Edge Chromium)
    'webview',
    'webview.platforms.edgechromium',
    'webview.platforms.mshtml',
    'webview.guilib',
    'webview.http',
    'webview.dom',
    'webview.dom.dom',
    'webview.dom.element',
    'webview.dom.event',
    'webview.dom.classlist',
    'webview.dom.propsdict',
    'webview.window',
    'webview.screen',
    'webview.menu',
    'webview.models',
    'webview.util',
    'webview.errors',
    'webview.event',
    'webview.localization',
    'webview.state',

    # Selenium
    'selenium',
    'selenium.webdriver',
    'selenium.webdriver.chrome',
    'selenium.webdriver.chrome.options',
    'selenium.webdriver.chrome.service',
    'selenium.webdriver.chrome.webdriver',
    'selenium.webdriver.common.by',
    'selenium.webdriver.common.action_chains',
    'selenium.webdriver.common.desired_capabilities',
    'selenium.webdriver.common.options',
    'selenium.webdriver.remote.webdriver',
    'selenium.webdriver.remote.command',
    'selenium.webdriver.support.ui',
    'selenium.webdriver.support.expected_conditions',

    # webdriver_manager
    'webdriver_manager',
    'webdriver_manager.chrome',

    # Google / gspread
    'gspread',
    'gspread.auth',
    'google.oauth2',
    'google.oauth2.credentials',
    'google.auth',
    'google.auth.transport',
    'google.auth.transport.requests',
    'google_auth_oauthlib',
    'google_auth_oauthlib.flow',
    'googleapiclient',
    'googleapiclient.discovery',
    'googleapiclient.http',

    # Módulos internos do projeto
    'ui.app',
    'rco.auth',
    'rco.api_client',
    'rco.consultas',
    'rco.rate_limiter',
    'rco.exceptions',
    'rco.escolas',
    'rco.notas',
    'sheets.gerador',
    'sheets.sincronizar',
    'calendario.calendario_pr',

    # stdlib às vezes não detectados
    'unicodedata',
    'json',
    'subprocess',

    # pycparser — dependência indireta do cffi; módulos opcionais que evitam warnings
    'pycparser',
]

a = Analysis(
    ['main.py'],
    pathex=['.'],
    binaries=[],
    datas=datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter',
        'matplotlib',
        'numpy',
        'pandas',
        'PIL',
        'PyQt5',
        'wx',
        'pythonnet',
        'clr',
        'clr_loader',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='RCO Manager',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon=icon,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='RCO Manager',
)
