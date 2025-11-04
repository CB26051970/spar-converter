import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import openpyxl
import pandas as pd
import os
import pdfplumber
from pathlib import Path

class PDFConverter:
    def __init__(self, pdf_file):
        self.pdf_file = pdf_file
        
    def extract_data_from_pdf(self):
        """Estrae i dati dall'ordine PDF"""
        try:
            with pdfplumber.open(self.pdf_file) as pdf:
                all_data = []
                
                for page in pdf.pages:
                    # Estrai le tabelle dalla pagina
                    tables = page.extract_tables()
                    
                    for table in tables:
                        for row in table:
                            # Filtra righe vuote e header
                            if row and len(row) >= 3:
                                # Cerca righe con dati numerici (articoli)
                                if (row[0] and row[0].strip() and 
                                    any(char.isdigit() for char in str(row[0])) and
                                    row[1] and row[2]):
                                    all_data.append(row)
                
                return all_data
                
        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile leggere il PDF: {str(e)}")
            return None
    
    def pdf_to_excel(self):
        """Converte il PDF in un file Excel temporaneo"""
        data = self.extract_data_from_pdf()
        if not data:
            return None
        
        try:
            # Crea un nuovo workbook
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Order Data"
            
            # Intestazioni
            headers = ["Article Ref", "Cases Ordered", "Unit Qty"]
            ws.append(headers)
            
            # Aggiungi i dati
            for row in data:
                if len(row) >= 3:
                    # Pulisci i dati
                    clean_row = []
                    for cell in row[:3]:  # Prendi solo le prime 3 colonne
                        if cell:
                            # Rimuovi spazi extra e converte virgole in punti
                            clean_cell = str(cell).strip().replace(',', '.')
                            clean_row.append(clean_cell)
                        else:
                            clean_row.append("")
                    
                    # Assicurati che ci siano 3 colonne
                    while len(clean_row) < 3:
                        clean_row.append("")
                    
                    ws.append(clean_row)
            
            # Salva il file Excel temporaneo
            temp_file = os.path.join(os.path.dirname(self.pdf_file), 
                                   f"temp_conversion_{os.path.basename(self.pdf_file).replace('.pdf', '.xlsx')}")
            wb.save(temp_file)
            wb.close()
            
            return temp_file
            
        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile convertire PDF in Excel: {str(e)}")
            return None

class SparConverter:
    def __init__(self, conversion_file, input_file):
        self.conversion_file = conversion_file
        self.input_file = input_file
        self.wb = None
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
        # 1. Rimuovi tutti i merge
        merged_ranges = list(self.ws.merged_cells.ranges)
        for merged_range in merged_ranges:
            self.ws.unmerge_cells(str(merged_range))
        
        # 2. Rimuovi wrap text da tutte le celle
        for row in self.ws.iter_rows():
            for cell in row:
                if cell.alignment:
                    cell.alignment = openpyxl.styles.Alignment(
                        wrap_text=False,
                        vertical=cell.alignment.vertical,
                        horizontal=cell.alignment.horizontal
                    )
        
        # 3. Imposta altezza uniforme di 15 per tutte le righe
        for row in range(1, self.ws.max_row + 1):
            self.ws.row_dimensions[row].height = 15
        
        # 4. Auto-adatta la larghezza di tutte le colonne
        for col in range(1, self.ws.max_column + 1):
            max_length = 0
            col_letter = openpyxl.utils.get_column_letter(col)
            for row in range(1, self.ws.max_row + 1):
                try:
                    cell_value = self.ws.cell(row=row, column=col).value
                    if cell_value:
                        max_length = max(max_length, len(str(cell_value)))
                except:
                    pass
            adjusted_width = (max_length + 2)
            self.ws.column_dimensions[col_letter].width = adjusted_width
    
    def get_start_row(self):
        """Chiede all'utente la riga di partenza"""
        root = tk.Tk()
        root.withdraw()
        
        user_input = simpledialog.askstring(
            "Riga di Partenza", 
            "Inserisci il numero della riga di partenza (es. 5 o 6):", 
            initialvalue="2"
        )
        
        root.destroy()
        
        if user_input is None or user_input == "":
            return None
        
        try:
            return int(user_input)
        except ValueError:
            messagebox.showerror("Errore", "Inserisci un numero valido!")
            return None
    
    def load_conversion_table(self):
        """Carica la tabella di conversione dal file SPAR CONVERSION.xlsm"""
        try:
            conversion_wb = openpyxl.load_workbook(self.conversion_file)
            conversion_ws = conversion_wb['Sheet1']
            
            # Crea un dizionario per la conversione (colonna B -> colonna C)
            conversion_dict = {}
            for row in range(1, 131):  # Da riga 1 a 130
                key_cell = conversion_ws[f'B{row}']
                value_cell = conversion_ws[f'C{row}']
                if key_cell.value is not None and value_cell.value is not None:
                    conversion_dict[key_cell.value] = value_cell.value
            
            conversion_wb.close()
            return conversion_dict
            
        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile caricare la tabella di conversione: {str(e)}")
            return None
    
    def apply_vlookup(self, conversion_dict):
        """Applica l'equivalente di VLOOKUP nella colonna C"""
        last_row = self.ws.max_row
        
        for row in range(self.start_row, last_row + 1):
            try:
                lookup_value = self.ws[f'A{row}'].value
                if lookup_value is not None:
                    # Converti a intero se possibile
                    try:
                        lookup_int = int(lookup_value)
                        result = conversion_dict.get(lookup_int, 0)
                    except (ValueError, TypeError):
                        result = 0
                else:
                    result = 0
                
                self.ws[f'C{row}'] = result
            except Exception as e:
                self.ws[f'C{row}'] = 0
    
    def insert_column_and_apply_formula(self):
        """Inserisce una colonna tra C e D e applica la formula IF"""
        # Inserisce colonna D (dopo C)
        self.ws.insert_cols(4)
        
        last_row = self.ws.max_row
        
        # Codici speciali per le moltiplicazioni (come nel VBA originale)
        multiply_4_codes = [11005101, 11005102, 11005111, 11005112, 11005107, 11005113]
        multiply_3_codes = [11005382, 11005387]
        multiply_2_codes = [11004140, 11004141]
        
        for row in range(self.start_row, last_row + 1):
            try:
                code = self.ws[f'C{row}'].value
                value_e = self.ws[f'E{row}'].value
                
                # Gestione valori None
                if value_e is None:
                    value_e = 0
                else:
                    try:
                        value_e = float(value_e)
                    except (ValueError, TypeError):
                        value_e = 0
                
                # Applica le moltiplicazioni come nel VBA originale
                if code in multiply_4_codes:
                    result = value_e * 4
                elif code in multiply_3_codes:
                    result = value_e * 3
                elif code in multiply_2_codes:
                    result = value_e * 2
                else:
                    result = value_e * 1
                
                self.ws[f'D{row}'] = result
                
            except Exception as e:
                self.ws[f'D{row}'] = 0
    
    def delete_zero_rows(self):
        """Elimina le righe con 0 nella colonna C"""
        rows_to_delete = []
        
        # Trova le righe da eliminare
        for row in range(self.start_row, self.ws.max_row + 1):
            try:
                cell_value = self.ws[f'C{row}'].value
                if cell_value == 0 or cell_value == "0":
                    rows_to_delete.append(row)
            except:
                pass
        
        # Elimina le righe dalla fine per evitare problemi con gli indici
        deleted_count = 0
        for row in sorted(rows_to_delete, reverse=True):
            self.ws.delete_rows(row)
            deleted_count += 1
        
        return deleted_count
    
    def convert(self, is_pdf_conversion=False):
        """Esegue l'intero processo di conversione"""
        if not self.load_workbook():
            return False
        
        # PRE-STEP: Formattazione iniziale
        self.pre_processing()
        
        # INPUT: Chiedi all'utente la riga di partenza
        self.start_row = self.get_start_row()
        if self.start_row is None:
            return False
        
        # Verifica che la riga di partenza sia valida
        if self.start_row > self.ws.max_row:
            messagebox.showerror("Errore", "La riga di partenza è oltre l'ultima riga con dati!")
            return False
        
        # Carica la tabella di conversione
        conversion_dict = self.load_conversion_table()
        if conversion_dict is None:
            return False
        
        # PRIMO STEP: Applica VLOOKUP nella colonna C
        self.apply_vlookup(conversion_dict)
        
        # SECONDO STEP: Inserisce una colonna tra C e D
        self.insert_column_and_apply_formula()
        
        # TERZO STEP: Elimina righe con 0 nella colonna C
        deleted_rows = self.delete_zero_rows()
        
        # QUARTO STEP: Ri-applica auto-fit alle colonne
        for col in range(1, self.ws.max_column + 1):
            max_length = 0
            col_letter = openpyxl.utils.get_column_letter(col)
            for row in range(1, self.ws.max_row + 1):
                try:
                    cell_value = self.ws.cell(row=row, column=col).value
                    if cell_value:
                        max_length = max(max_length, len(str(cell_value)))
                except:
                    pass
            adjusted_width = (max_length + 2)
            self.ws.column_dimensions[col_letter].width = adjusted_width
        
        # Salva il file convertito
        base_name = os.path.splitext(self.input_file)[0]
        if is_pdf_conversion:
            # Rimuovi "temp_conversion_" dal nome del file
            clean_name = base_name.replace("temp_conversion_", "")
            output_file = f"{clean_name}_CONVERTITO.xlsx"
        else:
            output_file = f"{base_name}_CONVERTITO.xlsx"
        
        try:
            self.wb.save(output_file)
            self.wb.close()
            
            # Messaggio di completamento
            messagebox.showinfo(
                "Automazione Completata!",
                f"Conversione terminata con successo!\n\n"
                f"Riga di partenza: {self.start_row}\n"
                f"Righe eliminate: {deleted_rows}\n"
                f"File salvato come: {os.path.basename(output_file)}\n\n"
                f"Operazioni completate:\n"
                f"• Rimozione merge cells\n"
                f"• Formattazione uniforme\n"
                f"• Applicazione VLOOKUP\n"
                f"• Inserimento colonna calcolo\n"
                f"• Eliminazione righe con zero"
            )
            return True
            
        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile salvare il file: {str(e)}")
            return False

def select_file(title, file_types):
    """Seleziona un file tramite dialog"""
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title=title, filetypes=file_types)
    root.destroy()
    return file_path

def main():
    try:
        # Seleziona il file di conversione SPAR
        conversion_file = select_file(
            "Seleziona il file SPAR CONVERSION.xlsm",
            [("Excel files", "*.xlsm"), ("All files", "*.*")]
        )
        
        if not conversion_file:
            return
        
        # Chiedi all'utente se vuole convertire PDF o usare Excel
        root = tk.Tk()
        root.withdraw()
        
        choice = messagebox.askquestion(
            "Tipo di File",
            "Vuoi convertire un file PDF o un file Excel?\n\n"
            "• Sì = Converti PDF\n"
            "• No = Usa file Excel esistente",
            icon='question'
        )
        root.destroy()
        
        input_file = None
        is_pdf_conversion = False
        
        if choice == 'yes':
            # Conversione PDF
            pdf_file = select_file(
                "Seleziona il file PDF da convertire",
                [("PDF files", "*.pdf"), ("All files", "*.*")]
            )
            
            if pdf_file:
                messagebox.showinfo("Conversione PDF", "Sto convertendo il PDF in Excel...")
                pdf_converter = PDFConverter(pdf_file)
                input_file = pdf_converter.pdf_to_excel()
                is_pdf_conversion = True
                
                if not input_file:
                    return
        else:
            # File Excel esistente
            input_file = select_file(
                "Seleziona il file Excel da convertire",
                [("Excel files", "*.xlsx"), ("Excel files", "*.xls"), ("All files", "*.*")]
            )
        
        if not input_file:
            return
        
        # Esegue la conversione SPAR
        converter = SparConverter(conversion_file, input_file)
        success = converter.convert(is_pdf_conversion)
        
        # Pulisci file temporaneo se era una conversione PDF
        if is_pdf_conversion and input_file and os.path.exists(input_file):
            try:
                os.remove(input_file)
            except:
                pass  # Ignora errori di cancellazione
        
        if not success:
            messagebox.showerror("Errore", "La conversione non è stata completata.")
            
    except Exception as e:
        messagebox.showerror("Errore Critico", f"Si è verificato un errore: {str(e)}")

if __name__ == "__main__":
    main()
