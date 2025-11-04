import PyInstaller.__main__
import os

def build_executable():
    PyInstaller.__main__.run([
        'src/main.py',
        '--name=SparConverter',
        '--onefile',
        '--windowed',
        '--add-data=src/converter.py;.',
        '--add-data=src/utils.py;.',
        '--icon=icon.ico'  # Opzionale: aggiungi un'icona
    ])

if __name__ == "__main__":
    build_executable()
