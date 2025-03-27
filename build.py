import PyInstaller.__main__

PyInstaller.__main__.run([
    'main.py',
    '--name=ExcelMerger',
    '--windowed',
    '--onefile',
    '--clean',
    '--add-data=README.md:.'
]) 