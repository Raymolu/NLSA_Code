<<<<<<< HEAD
# -*- mode: python -*-

block_cipher = None


a = Analysis(['NordicLamSplitAnalysisInterface.py'],
             pathex=['C:\\Nordic\\Projets\\ProgramationDivers\\NordicLamSplitAnalysis\\Code'],
             binaries=[],
             datas=[('NordicLamSpecs.xlsx', '.')],
             hiddenimports=['pandas._libs.tslibs.timedeltas', 'pandas._libs.tslibs.np_datetime', 'pandas._libs.tslibs.nattype', 'pandas._libs.skiplist'],
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
          a.binaries,
          a.zipfiles,
          a.datas,
          name='NLSAnalyzer_6.04.2',
          debug=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=True )
=======
# -*- mode: python -*-

block_cipher = None


a = Analysis(['NordicLamSplitAnalysisInterface.py'],
             pathex=['C:\\Nordic\\Projets\\ProgramationDivers\\NordicLamSplitAnalysis\\Code'],
             binaries=[],
             datas=[('NordicLamSpecs.xlsx', '.')],
             hiddenimports=['pandas._libs.tslibs.timedeltas', 'pandas._libs.tslibs.np_datetime', 'pandas._libs.tslibs.nattype', 'pandas._libs.skiplist'],
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
          a.binaries,
          a.zipfiles,
          a.datas,
          name='NLSAnalyzer_6.04.2',
          debug=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=True )
>>>>>>> 3be3cb1c1744d355d05ff23bf1db2074bf1a4bcf
