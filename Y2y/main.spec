# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

py_files = [
    'main.py',
    'convert.py',
    'yg_fl.py',
    'jjr_fl.py',
    'sb_fl.py',
    'gjj_fl.py',
]

added_files = [
    ('docs\\python_logo.gif', 'docs'),
    ('docs\\config.cfg', 'docs'),
    ('docs\\template.xlsx', 'docs'),
    ('docs\\结算信息.xlsx', 'docs'),
    ('docs\\结算信息--原版备份.xlsx', 'docs'),
    ('docs\\员工工资科目代码.txt', 'docs'),
    ('docs\\经纪人工资科目代码.txt', 'docs'),
    ('docs\\readme.txt', 'docs'),
]

a = Analysis(py_files,
             pathex=['D:\\Users\\wms-office\\python-Projects\\y2y'],
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
          name='y2y',
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
               name='y2y')
