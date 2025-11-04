import PyInstaller.__main__
import os
import sys

def build_executable():
    # Determina il percorso dello script principale
    main_script = os.path.join('src', 'main.py')
    
    # Parametri per PyInstaller
    params = [
        main_script,
        '--name=SparConverter',
        '--onefile',
        '--windowed',
        '--hidden-import=pandas',
        '--hidden-import=openpyxl',
        '--hidden-import=tkinter',
        '--collect-all=pandas',
        '--collect-all=openpyxl',
    ]
    
    # Aggiungi icona se presente
    if os.path.exists('icon.ico'):
        params.append('--icon=icon.ico')
    
    PyInstaller.__main__.run(params)

if __name__ == "__main__":
    build_executable()
