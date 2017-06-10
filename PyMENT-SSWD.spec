# -*- mode: python -*-

block_cipher = None


a = Analysis(['pyment/__main__.py'],
             pathex=['/Users/gysco/EPITECH/IRSN/SSWD'],
             binaries=[],
             datas=[],
             hiddenimports=[],
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
          name='PyMENT-SSWD',
          debug=False,
          strip=False,
          upx=True,
          console=False , icon='pyment.icns')
app = BUNDLE(exe,
             name='PyMENT-SSWD.app',
             icon='pyment.icns',
             bundle_identifier='fr.irsn.pyment.sswd')
