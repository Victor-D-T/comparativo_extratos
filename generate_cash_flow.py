import pandas as pd
import os
import json

class GenerateCashFlow:
    def __init__(self) -> None:
        self.sophia_folder_path = os.path.join(os.getcwd(),'Sophia')

    def main(self):
        self.sophias_cash_flow()
        self.create_excel()

        return self.final_dict

    def extract_sophias_transactions_data(self, file):  
        df = pd.read_excel(file, engine='xlrd')
        if df.get("CLASSIFIC_COD") is not None:
            df = df[["CLASSIFIC_COD", "CLASSIFIC_DESC", "VALOR_RECEB"]]

            self.pagas = df.groupby(["CLASSIFIC_COD", "CLASSIFIC_DESC"])["VALOR_RECEB"].sum().reset_index()
            self.pagas = self.pagas.rename(columns={"CLASSIFIC_COD": "Conta", "CLASSIFIC_DESC": "Descricao_conta", "VALOR_RECEB":"Valor"}).reset_index(drop=True)

        elif df.get("PLANO_CONTAS") is not None:
            df = df[["PLANO_CONTAS", "PGTO_CLASSFIC"]]

            df[["Conta", "Descricao_conta"]] = df["PLANO_CONTAS"].str.split(" - ", n=1, expand=True)

            self.recebidas = df.groupby(["Conta", "Descricao_conta"])["PGTO_CLASSFIC"].sum().reset_index()
            self.recebidas = self.recebidas.rename(columns={"PGTO_CLASSFIC": "Valor"}).reset_index(drop=True)

    def create_excel(self):
        self.result = pd.concat([self.pagas, self.recebidas], ignore_index=True).reset_index(drop=True)
        self.result = self.result.sort_values("Conta")
        with pd.ExcelWriter(os.path.join(os.getcwd(), "fluxo_de_caixa.xlsx"), engine = "xlsxwriter" ) as writer:
            self.result.to_excel(writer, "fluxo de caixa", index=False )

    def sophias_cash_flow(self):
        files = []

        for f in os.listdir(self.sophia_folder_path):
            complete_file_path = os.path.join(self.sophia_folder_path, f)
            files.append(complete_file_path)

        for file in files:        
            if not file.lower().endswith(('.xls')):
                continue
            
            else:
                self.final_dict = self.extract_sophias_transactions_data(file)      
                

        return self.final_dict

