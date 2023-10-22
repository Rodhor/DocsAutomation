from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need
# fine tuning.
build_options = {'packages': [], 'excludes': ["gi", "gtk", "PyQt4", "PyQt5", "PyQt6", "PySide2", "PySide6", "shiboken2", "shiboken6", "wx"]}

import sys
base = 'Win32GUI' if sys.platform=='win32' else None

executables = [
    Executable('Viona.py', base=base, target_name = 'App')
]
include_files =[("logo.png")]

setup(name='App',
      version = '1',
      description = 'A viona app',
      options = {'build_exe': build_options},
      executables = executables)
