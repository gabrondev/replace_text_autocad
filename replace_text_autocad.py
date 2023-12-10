import os
from pyautocad import Autocad
import win32com.client
import os
import glob

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

def main():
    parent_folder = input("Digite o diretório pai: ")
    old_text = input("Digite o texto antigo: ")
    new_text = input("Digite o texto novo: ")
    dwg_files = get_dwg_files(parent_folder)
    find_and_replace(dwg_files, old_text, new_text)

main()