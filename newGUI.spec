# -*- mode: python -*-

block_cipher = None

added_files = [
('sources/', './sources')
]

a = Analysis(['C:\\Users\\dlangowski\\PycharmProjects\\ACG\\Benchmarking\\newGUI.py'],
             pathex=['C:\\Users\\dlangowski\\PycharmProjects\\ACG\\Benchmarking'],
             binaries=[],
             datas = added_files,
             hiddenimports=[],
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
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='Industry Reporting Gen',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          runtime_tmpdir= None,
          icon = 'sources/ACGiGen.ico',
          console=False)
