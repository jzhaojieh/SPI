# -*- mode: python -*-

block_cipher = None


a = Analysis(['prac2.py'],
             pathex=['C:\\Users\\huangjos\\Desktop\\PurchaseReq'],
             binaries=[],
             datas=[],
             hiddenimports=['win32timezone'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
for d in a.datas:
    if 'pyconfig' in d[0]:
        a.datas.remove(d)
        break

a.datas += [('spi.ico','C:\\Users\\huangjos\\Desktop\\PurchaseReq\\spi.ico','Data')]

pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='PurchaseReq',
          debug=False,
          strip=False,
          upx=True,
          console=False, icon='C:\\Users\\huangjos\\Desktop\\PurchaseReq\\spi.ico' )
