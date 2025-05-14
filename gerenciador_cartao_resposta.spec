# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['gerenciador_cartao_resposta.py'],
    pathex=[os.path.abspath('.')],  # Caminho absoluto para o diretório do projeto
    binaries=[('pyzbar/libzbar-64.dll', 'pyzbar'),
              ('pyzbar/libiconv.dll', 'pyzbar')],  # Incluindo bibliotecas específicas do pyzbar
    datas=[('tesseract', 'tesseract'),
           ('pyzbar', 'pyzbar'),  # Incluindo a pasta pyzbar inteira
           ('colaboradores.xlsx', '.'),
           ('pops_gabaritos.json', '.'),
           ('ultima_pasta.json', '.')],  # Incluindo outros arquivos necessários
    hiddenimports=['reportlab.graphics.barcode.code128',
                   'reportlab.graphics.barcode.code93',
                   'reportlab.graphics.barcode.code39',
                   'reportlab.graphics.barcode.usps',
                   'reportlab.graphics.barcode.usps4s',
                   'reportlab.graphics.barcode.ecc200datamatrix',
                   'reportlab.graphics.barcode.eanbc',
                   'reportlab.graphics.barcode.i2of5',
                   'reportlab.graphics.barcode.qr'],
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
    a.binaries,
    a.datas,
    [],
    name='gerenciador_cartao_resposta',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
