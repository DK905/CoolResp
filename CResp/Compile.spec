# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

added_files = [
			   ( 'src\\About.ico', '.' ),
			   ( 'src\\CheckMark.ico', '.' ),
			   ( 'src\\CResp.ico', '.' ),
			   ( 'src\\Git.ico', '.' ),
			   ( 'src\\Help.ico', '.' ),
			   ( 'src\\Short_Help.png', '.' )
			  ]

a = Analysis(['main_PC.py'],
             pathex=['F:\\CoolResp\\CResp'],
             binaries=[],
             datas=added_files,
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
          name='CoolResp',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False , icon='src\\CResp.ico', version='info.txt')
