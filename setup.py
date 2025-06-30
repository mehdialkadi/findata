from setuptools import setup

APP = ['app.py']
DATA_FILES = ['icon.icns', 'templates', 'script_dependencies']
OPTIONS = {
    'argv_emulation': True,
    'iconfile': 'icon.icns',
    'packages': ['fitz', 'pytesseract', 'pandas', 'numpy', 'openpyxl', 'pdfplumber'],
    'includes': ['PyQt5', 'PyMuPDF'],
    'resources': ['templates', 'script_dependencies']
}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
