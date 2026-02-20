VERSION = "1.3.1"

from openpyxl import Workbook
import pandas as pd
import os
import re
from sophias_cashflow import SophiasCashflow
from extratos_cashflow import ExtratosCashFlow
from generate_cash_flow import GenerateCashFlow

class CashFlowComparative:
    def __init__(self, progress_callback=None) -> None:
        self.final_dict = {}
        self.sophias_cashflow = SophiasCashflow()
        self.extratos_cashflow = ExtratosCashFlow()
        self.generate_cashflow = GenerateCashFlow()
        self._progress = progress_callback or (lambda pct, msg: None)

        self.main()

    def main(self):
        self._progress(15, "Lendo arquivos Sophia...")
        self.sophias_cashflow.main(self.final_dict)

        self._progress(45, "Lendo extratos bancários...")
        self.extratos_cashflow.main(self.final_dict)

        self._progress(72, "Gerando comparativo de caixa...")
        self.__compare_sophia_and_extratos()

        self._progress(90, "Gerando fluxo de caixa...")
        self.generate_cashflow.main()

    
    def __compare_sophia_and_extratos(self):
        data = self.final_dict
        if not data:
            raise Exception("Nenhum extrato bancário foi lido. Verifique se a pasta Extratos/ contém os arquivos corretos.")
        final_df = pd.DataFrame()

        for bank, bank_data in data.items():
    
            dates = set()
            for source in ['sophia', 'extrato']:
                dates.update(bank_data['recebidas'][source].keys())
                dates.update(bank_data['pagas'][source].keys())
            dates = sorted([d for d in dates if d != '1970-01-01']) 
           
            recebidas_df = pd.DataFrame(index=dates, columns=['sophia', 'extrato', 'diff'])
            pagas_df = pd.DataFrame(index=dates, columns=['sophia', 'extrato', 'diff'])
            
            for date in dates:
                recebidas_df.loc[date, 'sophia'] = bank_data['recebidas']['sophia'].get(date, 0)
                recebidas_df.loc[date, 'extrato'] = bank_data['recebidas']['extrato'].get(date, 0)
                recebidas_df.loc[date, 'diff'] = recebidas_df.loc[date, 'sophia'] - recebidas_df.loc[date, 'extrato']
            
            for date in dates:
                pagas_df.loc[date, 'sophia'] = bank_data['pagas']['sophia'].get(date, 0)
                pagas_df.loc[date, 'extrato'] = bank_data['pagas']['extrato'].get(date, 0)
                pagas_df.loc[date, 'diff'] = pagas_df.loc[date, 'sophia'] - pagas_df.loc[date, 'extrato']
            
            recebidas_sum = recebidas_df.sum()
            pagas_sum = pagas_df.sum()
            
            bank_df = pd.concat([
                recebidas_df.rename(columns=lambda x: f"recebidas_{x}"),
                pagas_df.rename(columns=lambda x: f"pagas_{x}")
            ], axis=1)
            
            bank_df.loc['soma'] = {
                'recebidas_sophia': recebidas_sum['sophia'],
                'recebidas_extrato': recebidas_sum['extrato'],
                'recebidas_diff': round(recebidas_sum['diff'],2),
                'pagas_sophia': pagas_sum['sophia'],
                'pagas_extrato': pagas_sum['extrato'],
                'pagas_diff': round(pagas_sum['diff'],2)
            }
            
            bank_df['bank'] = bank
            
            final_df = pd.concat([final_df, bank_df])

        final_df = final_df.reset_index().rename(columns={'index': 'date'})
        final_df = final_df[['bank', 'date'] + [col for col in final_df.columns if col not in ['bank', 'date']]]

        self.__write_to_excel(final_df)

    def __write_to_excel(self, final_df):
                    
        with pd.ExcelWriter(os.path.join(os.getcwd(), 'comparativo_de_caixa.xlsx'), engine = "xlsxwriter" ) as writer:
            for bank, group in final_df.groupby('bank'):
                sheet_name = bank.replace(' ', '_')[:31] 
                group_without_bank = group.drop('bank', axis=1)
                group_without_bank.to_excel(writer, sheet_name=sheet_name, index=False)

                workbook = writer.book
                worksheet = writer.sheets[sheet_name]

                border_format = workbook.add_format({'border': 1})  # Basic border for all cells
                header_format = workbook.add_format({
                    'bold': True,
                    'bg_color': '#B8CCE4',
                    'border': 1,
                    'align': 'center'
                })
                soma_format = workbook.add_format({
                    'bold': True,
                    'bg_color': '#DCE6F1',
                    'border': 1,
                    'align': 'center'
                })
                soma_diff_format = workbook.add_format({
                    'bold': True,
                    'bg_color': '#DCE6F1',
                    'font_color': '#FF0000',
                    'border': 1,
                    'align': 'center'
                })



                num_rows, num_cols = group_without_bank.shape
                for row in range(num_rows + 1):  # +1 to include header row
                    for col in range(num_cols):
                        if row == 0:  # Header row
                            worksheet.write(row, col, group_without_bank.columns[col], header_format)
                        elif row == num_rows:  # Soma row (last row)
                            worksheet.write(row, col, group_without_bank.iloc[row-1, col], soma_format)
                            if col == 3 or col == 6:
                                 worksheet.write(row, col, group_without_bank.iloc[row-1, col], soma_diff_format)
                        else:  # Regular data cells
                            worksheet.write(row, col, group_without_bank.iloc[row-1, col], border_format)
                            
                # Auto-adjust column width based on the content

                for col_num, col in enumerate(group.drop('bank', axis=1).columns.values):
                    # Find the maximum length of content in the column (including header)
                    max_length = max(group[col].astype(str).map(len).max(), len(col))
                    worksheet.set_column(col_num, col_num, max_length + 2)  # Add a little padding
        

if __name__ == "__main__":
    from gui import App
    app = App()
    app.mainloop()