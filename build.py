"""Build the single-file Windows exe for the Email Archive Tool.

Usage:  python build.py
Output: dist/EmailArchiveTool.exe (self-contained, no Python needed)
"""

import os
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(HERE, 'dist')
WORK_DIR = os.path.join(HERE, 'build')


def main():
    for path in (OUT_DIR, WORK_DIR):
        if os.path.exists(path):
            shutil.rmtree(path)
    code = subprocess.call([
        sys.executable, '-m', 'PyInstaller', '--noconfirm', '--clean',
        '--onefile', '--windowed',
        '--distpath', OUT_DIR,
        '--workpath', os.path.join(WORK_DIR, 'work'),
        '--specpath', WORK_DIR,
        '--name', 'EmailArchiveTool',
        '--version-file', os.path.join(HERE, 'version_info.txt'),
        os.path.join(HERE, 'gui.py')])
    if code != 0:
        raise SystemExit('PyInstaller failed')
    exe = os.path.join(OUT_DIR, 'EmailArchiveTool.exe')
    print('built:', exe, os.path.getsize(exe), 'bytes')


if __name__ == '__main__':
    sys.exit(main() or 0)