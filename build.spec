import os

import gooey

gooey_root = os.path.dirname(gooey.__file__)

gooey_languages = Tree(
    os.path.join(gooey_root, 'languages'), prefix='gooey/languages')

gooey_images = Tree(os.path.join(gooey_root, 'images'), prefix='gooey/images')

a = Analysis(
    ['VMIQuoteGen.py'],
    hiddenimports=[
        'pandas._libs.tslibs.timedeltas', 'pandas._libs.tslibs.np_datetime',
        'pandas._libs.tslibs.nattype', 'pandas._libs.skiplist'
    ],
    hookspath=None,
    runtime_hooks=None)

pyz = PYZ(a.pure, a.zipped_data)

options = [('u', None, 'OPTION')]

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    options,
    gooey_languages,  # Add them in to collected files
    gooey_images,  # Same here.
    name='VMIQuoteGen',
    debug=False,
    strip=None,
    windowed=True,
    upx=True,
    console=False,
    icon=os.path.join('images', 'VMIQuoteGen.icns'))

app = BUNDLE(
    exe,
    name='VMIQuoteGen.app',
    info_plist={'NSHighResolutionCapable': 'True'},
    icon=os.path.join('images', 'VMIQuoteGen.icns'))
