# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['flasktest.py'],
             pathex=['C:\\Users\\tom_b\\.virtualenvs\\final-DIRW199q', 'C:\\Users\\tom_b\\Desktop\\Documents\\Python\\vcoding\\final'],
             binaries=[],
             datas=[],
             hiddenimports=['xlrd', 'pkg_resources.py2_warn'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
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
          name='flasktest',
          debug=True,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=True)
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='flasktest')
