import pandas as pd

class ReadExtratoItau:
    def __init__(self, file, final_dict) -> None:
        self.file = file
        self.final_dict = final_dict

    def read_extrato(self, bank):
        df = pd.read_excel(self.file, sheet_name="Lançamentos", skiprows=9, usecols="A:E")
        df.columns = df.columns.str.lower()
        df["data"] = pd.to_datetime(df["data"], dayfirst=True)
        df["data"] = df["data"].dt.strftime('%Y-%m-%d')
        recebidas = df[df["valor (r$)"]>0]
        recebidas = recebidas.groupby("data")["valor (r$)"].sum().reset_index()
        
        pagas = df[df["valor (r$)"]<0]
        pagas = pagas.groupby("data")["valor (r$)"].sum().reset_index()

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
            self.final_dict[bank]["recebidas"]["extrato"][str(row["data"])] = row["valor (r$)"]
        
        for _, row in pagas.iterrows():
            self.final_dict[bank]["pagas"]["extrato"][str(row["data"])] = row["valor (r$)"]
        
        return self.final_dict

