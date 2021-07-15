# -*- mode: python -*-
from pathlib import Path

import ahkunwrapped

a = Analysis(
    ['example.py'],
    datas=[
        (Path(ahkunwrapped.__file__).parent / 'lib', 'lib'),
        ('example.ahk', '.'),
    ]
)
pyz = PYZ(a.pure)

# for onefile
exe = EXE(pyz, a.scripts, a.binaries, a.datas, name='my-example', upx=True, console=False)
# for onedir
# exe = EXE(pyz, a.scripts, exclude_binaries=True, name='my-example', upx=True, console=False)
# dir = COLLECT(exe, a.binaries, a.datas, name='my-example-folder')
