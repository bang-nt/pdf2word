import sys
from cx_Freeze import setup, Executable

shortcut_table = [
    (
        "DesktopShortcut",        # Shortcut
        "DesktopFolder",          # Directory_
        "PDF and Image to Word",    # Name
        "TARGETDIR",              # Component_
        "[TARGETDIR]pdf2word.exe",# Target
        None,                     # Arguments
        None,                     # Description
        None,                     # Hotkey
        None,                     # Icon
        None,                     # IconIndex
        None,                     # ShowCmd
        'TARGETDIR',              # WkDir
    ),
    (
        "StartMenuShortcut",
        "ProgramMenuFolder",
        "PDF and Image to Word",
        "TARGETDIR",
        "[TARGETDIR]pdf2word.exe",
        None,
        None,
        None,
        None,
        None,
        None,
        'TARGETDIR',
    ),
]

msi_data = {
    'Shortcut': shortcut_table,
}

build_options = {
    'packages': [],
    # 'includes': ["socks"],
    'excludes': ["sqlite3", "tkinter", "unittest"],
    'zip_includes': [],
    'no_compress': True
}

bdist_msi_options = {
    'data': msi_data
}

base = 'gui'

executables = [
    Executable('main.py', base=base, target_name='pdf2word.exe'),
]

setup(
    name='PDFImageToWordConverter',
    version='1.0',
    description="PDF & Image to Word Converter",
    options={
        'build_exe': build_options,
        'bdist_msi': bdist_msi_options
    },
    executables=executables
)

# Run the setup script
# python setup.py build
# To create an MSI installer, run the following command:
# python setup.py bdist_msi