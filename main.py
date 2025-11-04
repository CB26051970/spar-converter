import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import pandas as pd
import openpyxl
import os

class SparConverter:
    def __init__(self, conversion_file, input_file):
        self.conversion_file = conversion_file
        self.input_file = input_file
        self.ws = None
        self.start_row = None
        
    def load_workbook(self):
        """Carica il file Excel di input"""
        try:
            self.wb = openpyxl.load_workbook(self.input_file)
            self.ws = self.wb.active
            return True
        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile caricare il file: {str(e)}")
            return False
    
    def pre_processing(self):
        """Esegue il pre-processing: rimuove merge, wrap text, etc."""
        # Rimuovi tutti i merge
        for merged_range in list(self.ws.merged_cells.ranges):
            self.ws.unmerge_cells(str(merged_range))
        
        # Rimuovi wrap text
        for row in self.ws.iter_rows():
            for cell in row:
                cell.alignment = openpyxl.styles.Alignment(wrap_text=False)
        
        # Imposta altezza uniforme delle righe
        for row in range(1, self.ws.max_row + 1):
            self.ws.row_dimensions[row].height = 15
        
        # Auto-adatta le colonne
        for column in self.ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2) * 1.2
            self.ws.column_dimensions[column_letter].width = adjusted_width
    
    def get_start_row(self):
        """Chiede all'utente la riga di partenza"""
        root = tk.Tk()
        root.withdraw()
        
        user_input = simpledialog.askstring(
            "Riga di Partenza", 
            "Inserisci il numero della riga di partenza (es. 5 o 6):", 
            initialvalue="6"
        )
        
        if user_input is None:
            return None
        
        try:
            return int(user_input)
        except ValueError:
            messagebox.showerror("Errore", "Inserisci un numero valido!")
            return None

    def convert(self):
        """Esegue la conversione"""
        if not self.load_workbook():
            return False
        
        self.pre_processing()
        self.start_row = self.get_start_row()
        
        if self.start_row is None:
            return False
            
        messagebox.showinfo("Completato", f"Conversione completata!\nRiga di partenza: {self.start_row}")
        return True

def select_file(title, file_types):
    """Seleziona un file tramite dialog"""
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title=title, filetypes=file_types)
    root.destroy()
    return file_path

def main():
    conversion_file = select_file(
        "Seleziona il file SPAR CONVERSION.xlsm",
        [("Excel files", "*.xlsm"), ("All files", "*.*")]
    )
    
    if not conversion_file:
        return
    
    input_file = select_file(
        "Seleziona il file Excel da convertire",
        [("Excel files", "*.xlsx"), ("All files", "*.*")]
    )
    
    if not input_file:
        return
    
    converter = SparConverter(conversion_file, input_file)
    converter.convert()

if __name__ == "__main__":
    main()
