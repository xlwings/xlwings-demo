# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['demo.py'],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             # TODO: make sure no package appears here that you're using
             excludes=['scipy', 'numpy', 'pandas', 'Tkinter', 'Tkconstants', 'pydoc',
                       'tcl', 'tk', 'matplotlib', 'PIL', 'nose', 'setuptools', 'xlrd', 'xlwt', 'PyQt5',
                       'markdown', 'IPython', 'docutils'],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          [],
          exclude_binaries=True,
          name='demo',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=True )
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='demo')
