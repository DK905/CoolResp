# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

added_files = [
			   ( 'CoolRespProject/src/About.ico', '.' ),
			   ( 'CoolRespProject/src/CheckMark.ico', '.' ),
			   ( 'CoolRespProject/src/CResp.ico', '.' ),
			   ( 'CoolRespProject/src/Git.ico', '.' ),
			   ( 'CoolRespProject/src/Help.ico', '.' ),
			   ( 'CoolRespProject/src/Short_Help.png', '.' )
			  ]

a = Analysis(['main.py'],
             pathex=['E:/CoolResp/CoolRespProject'],
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
          console=False , icon='CoolRespProject/src/CResp.ico', version='info.txt')
