import pandas as pd
import os
import re
from bank_collection import BankCollection

class ExtratosCashFlow:
    def __init__(self) -> None:
        self.banks_dict = {"cora rd": "cora unidade 1 e 2",
                        "caixa": "caixa",
                        "cef 20-3":"caixa",
                        "cora iv": "cora unidade 3",
                        "itaú iv": "itau unidade 3",
                        "itaú rd": "itau unidade 1 e 2",
                        "sicredi rd": "sicredi unidade 1 e 2",
                        "sicredi iv": "sicredi unidade 3"}
                
        self.folder_path = os.path.join(os.getcwd(),'Extratos')
        self.files = []

        for f in os.listdir(self.folder_path):
            if f.startswith('~$'):
                continue
            complete_file_path = os.path.join(self.folder_path, f)
            self.files.append(complete_file_path)
        

    def main(self, final_dict):
        self.final_dict = final_dict 

        for file in self.files:
            clean_name = os.path.splitext(os.path.basename(file))[0].lower().replace('_', '')
            for bank in self.banks_dict.keys():
                if bank in clean_name:
                    banco = self.banks_dict[bank]
                    bank_collection = BankCollection(file, banco, self.final_dict)
                    self.final_dict = bank_collection.main()

        return self.final_dict
        
