# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['SupplierListManagerProject.pyw'],
<<<<<<< HEAD
             pathex=['C:\\Users\\Andrei Bancila\\Desktop\\Andrei Bancila\\Proiecte\\SupplierListManagerProject'],
=======
             pathex=['C:\\Users\\Andrei Bancila\\Desktop\\SupplierListManagerProject'],
>>>>>>> 15b89f049af001062626f79c23a8dde749b3e834
             binaries=[],
             datas=[],
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
          name='SupplierListManagerProject',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False )
