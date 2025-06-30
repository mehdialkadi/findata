from setuptools import setup

APP = ['app.py']
DATA_FILES = [
    ('templates', ['templates/index.html']),
    ('script_dependencies', [
        'script_dependencies/script.py',
        # add any other needed files
    ])
]
OPTIONS = {
    'argv_emulation': True,
    'packages': ['flask', 'pandas', 'numpy', 'openpyxl', 'pdfplumber', 'pytesseract', 'fitz', 'PIL', 'xlsxwriter'],
    'iconfile': 'icon.icns',  # optional
}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)

