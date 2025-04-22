import pandas as pd
import os
import re
from bank_collection import BankCollection

class ExtratosCashFlow:
    def __init__(self) -> None:
        self.banks_dict = {"cora rd": "cora unidade 1 e 2",
                        "caixa rd": "caixa unidade 1 e 2",
                        "cef irmãs vitória": "caixa unidade 3",
                        "cora iv": "cora unidade 3",
                        "itaú iv": "itau unidade 3",
                        "itaú rd": "itau unidade 1 e 2",
                        "sicredi rd": "sicredi unidade 1 e 2",
                        "sicredi iv": "sicredi unidade 3"}
        
        self.bancos = ["cora iv", "cora rd", "itau iv", "itau rd", "caixa rd", "sicredi iv", "sicredi rd"]
        
        self.folder_path = os.path.join(os.getcwd(),'Extratos')
        self.files = []

        for f in os.listdir(self.folder_path):
            complete_file_path = os.path.join(self.folder_path, f)
            self.files.append(complete_file_path)
        

    def main(self, final_dict):
        self.final_dict = final_dict 
        for file in self.files:
            clean_name = os.path.splitext(os.path.basename(file))[0].lower().replace('_', '')

            is_banco = False
            for bank in self.banks_dict.keys():
                if bank in clean_name:
                    banco = self.banks_dict[bank]
                    is_banco = True

            if is_banco is False:
                continue
            

            bank_collection = BankCollection(file, banco, self.final_dict)
            bank_collection.main()

            return self.final_dict
        
