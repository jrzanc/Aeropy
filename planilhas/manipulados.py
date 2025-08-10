
import tkinter as tk
from tkinter import filedialog

class planilha:
    def get_address(self):
        # Cria uma janela oculta
        root = tk.Tk()
        root.withdraw()
        # Abre a janela para selecionar o arquivo
        try:
            caminho_arquivo = filedialog.askopenfilename(
                title="Selecione um arquivo",
                filetypes=[("Arquivos Excel", "*.xls"),("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]
            )
            return caminho_arquivo
        except Exception as error:
            print(f"error type:{error}")
            
    def get_new_address(self,address): 
        address_splited=address.split("/")
        old_name_splited=address_splited[-1].split(".")
        new_name=old_name_splited[0]+"_MODIFIED_SPREADSHEET.xlsx"
        new_address = "/".join(address_splited[:-1]) + "/" + new_name
        return new_address

    def copy_spreadsheet(self,old_address,new_address):
        with open(old_address, 'rb') as origem:
            conteudo = origem.read()
        try:
            with open(new_address, 'wb') as destino:
                destino.write(conteudo)
            print("Arquivo Excel copiado com sucesso!")
        except Exception as error:
            print(f"error type:{error}")
   
    def flow(self):
        g_address=self.get_address()
        n_address=self.get_new_address(g_address)
        self.copy_spreadsheet(g_address,n_address)
        return n_address