# -*- mode: python -*-

block_cipher = None

import os
a = Analysis(['pyment\\__main__.py'],
             pathex=['D:\\Users\\Gysco\\Documents\\Git\\SSWD'],
             binaries=[],
             datas=[('.\\rsrc\\img\\pyment_splashart@300x100.png', '.')],
             hiddenimports=['xlrd', 'openpyxl', 'xlsxwriter'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name=os.environ['EXE_NAME'],
          debug=False,
          strip=False,
          upx=True,
          console=False , icon='rsrc\\img\\pyment.ico' )
