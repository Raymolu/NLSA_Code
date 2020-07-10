# -*- mode: python -*-

block_cipher = None


a = Analysis(['NordicLamSplitAnalysisInterface.py'],
             pathex=['C:\\Nordic\\Projets\\ProgramationDivers\\NordicLamSplitAnalysis\\NLSA_Code'],
             binaries=[],
             datas=[('templates/*xlsx','templates'),('NordicLamSpecs.xlsx', '.'), ('NLRepairSpecs.xlsx', '.'),\
			 ('nordic_N.ico','.'),('NLSA_Reference.pdf','.'),('reports/*','reports')],
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
          [],
          exclude_binaries=True,
          name='NLSAnalyzer_0.00.0',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
		  icon='nordic_N.ico',
          console=True)
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='NLSAnalyzer_0.00.0')
