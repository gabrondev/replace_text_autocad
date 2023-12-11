import os
from pyautocad import Autocad
import win32com.client
import os
import glob
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

def find_and_replace(dwg_files, old_text, new_text):
    acad = win32com.client.Dispatch("AutoCAD.Application") 
    acad = Autocad() 

    for file in dwg_files:
        acad.app.Documents.open(file) # Open current file in dwg_files
        print(acad.doc.Name) # Print curent file name
        doc = acad.ActiveDocument # Referencing current active AutoCAD document
        
        # Going trough all objects on drawing, modifying the texts that contains the "old_text"
        for entity in acad.ActiveDocument.ModelSpace:
            name = entity.EntityName
            if name == 'AcDbMText':
                # Checking multiline texts
                if old_text in entity.TextString:
                    text = str(entity.TextString)
                    text = text.replace(old_text, new_text)
                    entity.TextString = text

                # Checking singleline texts
                elif name == 'AcDbText': 
                    if old_text in entity.TextString:
                        text = str(entity.TextString)
                        text = text.replace (old_text, new_text)
                        entity.TextString = text
                    
        doc.Save()
        doc.Close()

def get_dwg_files(parent_folder):
    dwg_files = [] # List where all the DWG files will be stored
    
    # Walk trough all subfolders in the parent folder and storing the DWG files
    for parent_folder, subfolders, files in os.walk(parent_folder):
        dwg_files.extend(glob.glob(os.path.join(parent_folder, '*.dwg'))) # Filter all the DWGs files using the extension .dwg

    return dwg_files

def select_folder(entry_path):
    selected_folder = filedialog.askdirectory()
    entry_path.config(state='normal')
    entry_path.delete(0, tk.END)
    entry_path.insert(0, selected_folder)
    entry_path.config(state='readonly')

def get_parameters(entry_path, entry_old_text, entry_new_text):
    parent_folder = entry_path.get()
    old_text = entry_old_text.get()
    new_text = entry_new_text.get()

    start(parent_folder, old_text, new_text)

def start(parent_folder, old_text, new_text):
    dwg_files = get_dwg_files(parent_folder)
    find_and_replace(dwg_files, old_text, new_text)
    messagebox.showinfo("Processo concluído", "Todas as ocorrências encontradas foram substituídas.")

def main():
    window = tk.Tk()
    window.title("Substituir Texto no AutoCAD")

    # ENTRY AND BUTTON FOR SELECTING THE PARENT FOLDER
    label = tk.Label(window, text="Escolha o diretório pai:")
    label.grid(row=0, column=0, padx=10, sticky="w")

    entry_path = tk.Entry(window, state='readonly', width=50)
    entry_path.grid(row=1, column=0, padx=10, pady=10)

    select_button = tk.Button(window, text="Selecionar pasta", command=lambda: select_folder(entry_path))
    select_button.grid(row=1, column=1, padx=10, pady=10)

    #ENTRYS FOR TYPING THE NEW AND OLD TEXT
    label = tk.Label(window, text="Digite o texto a ser substituído:")
    label.grid(row=2, column=0, padx=10, sticky="w")

    entry_old_text = tk.Entry(window, state='normal', width=50)
    entry_old_text.grid(row=3, column=0, padx=10, pady=10, sticky="ew", columnspan=2)

    label = tk.Label(window, text="Digite o novo texto:")
    label.grid(row=4, column=0, padx=10, sticky="w")

    entry_new_text = tk.Entry(window, state='normal', width=50)
    entry_new_text.grid(row=5, column=0, padx=10, pady=10, sticky="ew", columnspan=2)

    replace_button = tk.Button(window, text="Localizar e substituir", command=lambda: get_parameters(entry_path, entry_old_text, entry_new_text))
    replace_button.grid(row=6, column=0, padx=10, pady=10, columnspan=2)

    window.mainloop()

main()