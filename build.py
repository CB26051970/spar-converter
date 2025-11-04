import PyInstaller.__main__
import os
import sys

def build_executable():
    # Parametri per PyInstaller
    params = [
        'src/main.py',
        '--name=SparConverter',
        '--onefile',
        '--windowed',
        '--hidden-import=pandas',
        '--hidden-import=openpyxl',
        '--hidden-import=tkinter',
        '--hidden-import=openpyxl.workbook',
        '--collect-all=pandas',
        '--collect-all=openpyxl',
    ]
    
    PyInstaller.__main__.run(params)

if __name__ == "__main__":
    build_executable()
