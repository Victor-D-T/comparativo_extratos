import pandas as pd
import os

class SophiasCashflow:
    def __init__(self) -> None:
        self.sophia_folder_path = os.path.join(os.getcwd(),'Sophia')
        self.banks_dict = {"banco cora": "cora unidade 1 e 2",
                        "caixa escola": "caixa",
                        "caixa unidade 2": "caixa",
                        "cef 20-3": "caixa",
                        "cora infantil": "cora unidade 3",
                        "itaú irmãs vitória": "itau unidade 3",
                        "itaú rd": "itau unidade 1 e 2",
                        "itaú rd (2)": "itau unidade 1 e 2",
                        "itaú rd (3)": "itau unidade 3",
                        "sicredi und 1": "sicredi unidade 1 e 2",
                        "sicredi und 3": "sicredi unidade 3"}

    def main(self, final_dict):
        self.final_dict = final_dict
        self.sophias_cash_flow()

        return self.final_dict        

    def extract_sophias_transactions_data(self, file):  
        engine = 'openpyxl' if file.lower().endswith('.xlsx') else 'xlrd'
        df = pd.read_excel(file, engine=engine)
        if df.get("DATA_EFETIVA") is not None:
            df["DATA_EFETIVA"] = pd.to_datetime(df["DATA_EFETIVA"])
            df["DATA_EFETIVA"] = df["DATA_EFETIVA"].dt.strftime('%Y-%m-%d')
            for idx, value in df["CONTA"].items():
                encountered = False
                for name in self.banks_dict.keys():
                    if name in value.lower():
                        value = name
                        encountered = True
                
                if encountered == False:
                    error_msg = f"O arquivo: '{value.lower()}' não foi encontrado na base de arquivos aceitos.\nNomes válidos: {list(self.banks_dict.keys())}"
    
                    print("ERRO CRÍTICO:", error_msg)
                    input("Pressione Enter para sair...")

                new_value = self.banks_dict[value.lower()]
                df.at[idx, "CONTA"] = new_value

            self.pagas = df.groupby(["CONTA", "DATA_EFETIVA"])["VALOR_RECEB"].sum().reset_index()
            for _, row in self.pagas.iterrows():
                if self.final_dict.get(row["CONTA"]) is None:
                    self.final_dict[row["CONTA"]] = {
                        "recebidas": {
                            "sophia": {},
                            "extrato": {}
                        },
                        "pagas": {
                            "sophia": {},
                            "extrato": {}
                        }
                    }

                self.final_dict[row["CONTA"]]["pagas"]["sophia"][str(row["DATA_EFETIVA"])] = -row["VALOR_RECEB"]
        elif df.get("RECEBTO") is not None:
            df["RECEBTO"] = pd.to_datetime(df["RECEBTO"])
            df["RECEBTO"] = df["RECEBTO"].dt.strftime('%Y-%m-%d')
            for idx, value in df["DESC_CONTA_DESTINO"].items():
                new_value = self.banks_dict[value.lower()]
                df.at[idx, "DESC_CONTA_DESTINO"] = new_value
                            
            self.recebidas = df.groupby(["DESC_CONTA_DESTINO", "RECEBTO"])["VALREC"].sum().reset_index()
            

            for _, row in self.recebidas.iterrows():
                if self.final_dict.get(row["DESC_CONTA_DESTINO"]) is None:
                    self.final_dict[row["DESC_CONTA_DESTINO"]] = {
                        "recebidas": {
                            "sophia": {},
                            "extrato": {}
                        },
                        "pagas": {
                            "sophia": {},
                            "extrato": {}
                        }
                    }
        
                self.final_dict[row["DESC_CONTA_DESTINO"]]["recebidas"]["sophia"][str(row["RECEBTO"])] = row["VALREC"]

            


        return self.final_dict
            
    def sophias_cash_flow(self):
        files = []

        for f in os.listdir(self.sophia_folder_path):
            complete_file_path = os.path.join(self.sophia_folder_path, f)
            files.append(complete_file_path)

        for file in files:
            if not file.lower().endswith(('.xls')):
                continue

            else:
                arquivo = os.path.basename(file)
                try:
                    self.final_dict = self.extract_sophias_transactions_data(file)
                except Exception as e:
                    raise Exception(f"Erro ao ler arquivo do Sophia '{arquivo}': {e}") from e

        return self.final_dict

