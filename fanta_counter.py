#Assicurati di aver installato le librerie occorrenti
import os
import openpyxl

def cerca_stringa_in_excel(file_path, stringa_da_cercare):
    cont = 0
    # Espandi il percorso se contiene la tilde (~)
    #Se usi la tilde nel tuo path del file.xlsx lascia la riga seguente altrimenti puoi eliminarla
    file_path = os.path.expanduser(file_path)
    
    # Carica il file Excel
    workbook = openpyxl.load_workbook(file_path)
    

    for sheet in workbook.worksheets:
        
        
        
        for row in sheet.iter_rows():
            for cell in row:
                
                if cell.value is not None:
                   
                    
                    # Rimuovi spazi vuoti e ignora maiuscole/minuscole
                    valore_cella = str(cell.value).strip().lower()
                    stringa_normalizzata = stringa_da_cercare.strip().lower()
                    
                    
                    if valore_cella == stringa_normalizzata:
                        cont += 1
                       
                        
    if cont == 0:
        print(f'Nessuno ha "{stringa_da_cercare}".')
    else:
        print(f'\u001b[31m {cont} squadre hanno "{stringa_da_cercare.upper()}".')


file_path = ''   #Qui inserisci il path di dove e' situato il tuo file.xlsx
stringa_da_cercare = input("Inserisci il giocatore da cercare: ") 

cerca_stringa_in_excel(file_path, stringa_da_cercare)
