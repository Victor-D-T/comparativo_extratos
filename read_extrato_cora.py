import pandas as pd

class ReadExtratoCora:
    def __init__(self, file, final_dict) -> None:
        self.file = file
        self.final_dict = final_dict

    def read_extrato(self, bank):
        df = pd.read_csv(self.file)
        df["Data"] = pd.to_datetime(df["Data"], dayfirst=True)
        df["Data"] = df["Data"].dt.strftime('%Y-%m-%d')
        recebidas = df[df["Valor"]>0]
        recebidas = recebidas.groupby("Data")["Valor"].sum().reset_index()
        
        pagas = df[df["Valor"]<0]
        pagas = pagas.groupby("Data")["Valor"].sum().reset_index()

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
            self.final_dict[bank]["recebidas"]["extrato"][str(row["Data"])] = row["Valor"]
        
        for _, row in pagas.iterrows():
            self.final_dict[bank]["pagas"]["extrato"][str(row["Data"])] = row["Valor"]
        

        return self.final_dict

