import gooey
gooey_root = os.path.dirname(gooey.__file__)
gooey_languages = Tree(os.path.join(gooey_root, 'languages'), prefix = 'gooey/languages')
gooey_images = Tree(os.path.join(gooey_root, 'images'), prefix = 'gooey/images')

a = Analysis(['VMIQuoteGen.py'],
             hiddenimports = ['pandas._libs.tslibs.timedeltas',
             'pandas._libs.tslibs.np_datetime', 'pandas._libs.tslibs.nattype',
             'pandas._libs.skiplist'],
             hookspath=None,
	     datas=[('./product_data.csv', 'data'), ('./config.json',
             'config'), ('./logos/PSSI Horz Logo.png', '/images')],
             runtime_hooks=None
             )
pyz = PYZ(a.pure, a.zipped_data)

options = [('u', None, 'OPTION')]

exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          options,
          gooey_languages, # Add them in to collected files
          gooey_images, # Same here.
          name='VMIQuoteGen',
          debug=False,
          strip=None,
          upx=True,
          console=False,
          icon=os.path.join(gooey_root, 'images', 'program_icon.ico'))

coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=None,
               upx=True,
               name='resources')
