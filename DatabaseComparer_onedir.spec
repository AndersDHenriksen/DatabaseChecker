# -*- mode: python -*-

block_cipher = None

added_files = [('*.ui', '.'), ('.\platforms', 'platforms')]

a = Analysis(['DatabaseComparer.py'],
             pathex=['C:\\Users\\Anders\\PycharmProjects\\DatabaseChecker'],
             binaries=[],
             datas=added_files,
             hiddenimports=['PyQt5', 'pandas._libs.tslibs.timedeltas'],
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
          exclude_binaries=True,
          name='DatabaseComparer',
          debug=False,
          strip=False,
          upx=True,
          console=False , icon='database_refresh.ico')
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               name='DatabaseComparer')
