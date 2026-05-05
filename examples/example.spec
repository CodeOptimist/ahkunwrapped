# SPDX-License-Identifier: 0BSD

from PyInstaller.utils.hooks import collect_data_files

import ahkunwrapped

a = Analysis(
    ['example.py'],                               # Python file
    datas=collect_data_files('ahkunwrapped') + [
        ('example.ahk', '.'),                     # AutoHotkey script (if using `Script.from_file()`)
    ]
)
pyz = PYZ(a.pure)

name = 'my-example'                               # used below

# for onefile
#exe = EXE(pyz, a.scripts, a.binaries, a.datas, name=name, upx=True, console=False)
# for onedir
exe = EXE(pyz, a.scripts, exclude_binaries=True, name=name, upx=True, console=False)
dir = COLLECT(exe, a.binaries, a.datas, name=name)
