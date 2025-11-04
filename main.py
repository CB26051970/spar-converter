import tkinter as tk
from tkinter import filedialog, messagebox
from converter import SparConverter
import os

def select_file(title, file_types):
    """Seleziona un file tramite dialog"""
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title=title, filetypes=file_types)
    root.destroy()
    return file_path

def main():
    # Seleziona il file di conversione
    conversion_file = select_file(
        "Seleziona il file SPAR CONVERSION.xlsm",
        [("Excel files", "*.xlsm"), ("All files", "*.*")]
    )
    
    if not conversion_file:
        return
    
    # Seleziona il file da convertire
    input_file = select_file(
        "Seleziona il file Excel da convertire",
        [("Excel files", "*.xlsx"), ("All files", "*.*")]
    )
    
    if not input_file:
        return
    
    # Esegue la conversione
    converter = SparConverter(conversion_file, input_file)
    converter.convert()

if __name__ == "__main__":
    main()
