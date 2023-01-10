# -*- mode: python ; coding: utf-8 -*-

# 在main.spec文件所在目录下启动管理员权限cmd，运行：pyinstaller main.spec

block_cipher = None

py_files = [
    'main.py',
    'SH_stock.py',
    'SH_fund.py',
    'SH_fund.py',
    'SZ.py',
    'writeXL.py',
]

added_files = [
    ('docs\\config.cfg', 'docs'),
    ('docs\\2023年市场交易量统计表_python.xls', 'docs'),
]

a = Analysis(py_files,
             pathex=[r'E:\python_code\Trade_amt'],
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
          [],
          exclude_binaries=True,
          name='Trade_amt',
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
               name='Trade_amt')
