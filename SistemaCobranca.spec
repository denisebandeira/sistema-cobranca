# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['app_desktop.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('modelo_comunicado.xlsx', '.'),
        ('modelo_importacao_contatos.xlsx', '.'),
    ],
    # A importacao direta ja e detectada pelo PyInstaller. Mantemos a entrada
    # explicita para documentar e garantir a inclusao do modulo local no .exe.
    hiddenimports=[
        'banco_dados', 'cargas_dados',
        'gerador_comunicados', 'modulo_comunicacoes', 'modulo_consultas',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='SistemaCobranca',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='SistemaCobranca',
)
app = BUNDLE(
    coll,
    name='SistemaCobranca.app',
    icon=None,
    bundle_identifier=None,
)
