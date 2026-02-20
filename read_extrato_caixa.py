import pandas as pd

class ReadExtratoCaixa:
    def __init__(self, file, final_dict) -> None:
        self.file = file
        self.final_dict = final_dict

    def read_extrato(self, bank):
        engine = 'openpyxl' if self.file.lower().endswith('.xlsx') else 'xlrd'
        df = pd.read_excel(self.file, skiprows=1, engine=engine)
        df["Data Lançamento"] = pd.to_datetime(df["Data Lançamento"], dayfirst=True)
        df["Data Lançamento"] = df["Data Lançamento"].dt.strftime('%Y-%m-%d')
        recebidas = df[df["Valor Lançamento"]>0]
        recebidas = recebidas.groupby("Data Lançamento")["Valor Lançamento"].sum().reset_index()
        
        pagas = df[df["Valor Lançamento"]<0]
        pagas = pagas.groupby("Data Lançamento")["Valor Lançamento"].sum().reset_index()

        if self.final_dict.get(bank) is None:
            self.final_dict[bank] = {
                "recebidas":{
                    "extrato":{},
                    "sophia":{}
                },
                "pagas":{
                    "extrato":{},
                    "sophia":{}
                }
                    
            }

        for _, row in recebidas.iterrows():
            self.final_dict[bank]["recebidas"]["extrato"][str(row["Data Lançamento"])] = row["Valor Lançamento"]
        
        for _, row in pagas.iterrows():
            self.final_dict[bank]["pagas"]["extrato"][str(row["Data Lançamento"])] = row["Valor Lançamento"]
        

        return self.final_dict

