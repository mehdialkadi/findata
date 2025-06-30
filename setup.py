from setuptools import setup

APP = ['app.py']
DATA_FILES = [
    ('templates', ['templates/index.html']),
    ('script_dependencies', ['script_dependencies/script.py']),
]
OPTIONS = {
    'argv_emulation': False,
    'iconfile': 'icon.icns',  # Optional
    'packages': ['flask', 'pandas', 'numpy', 'pdfplumber', 'pytesseract', 'pymupdf', 'Pillow'],
    'excludes': ['PyQt5', 'PyQt6', 'PySide2', 'PySide6', 'tkinter'],
}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
