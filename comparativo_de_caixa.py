VERSION = "1.3.4"

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
        self.generate_cashflow.main(self.final_dict)

    
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

    def __build_resumo(self, final_df):
        soma_rows = final_df[final_df['date'] == 'soma'].copy()
        soma_rows = soma_rows.set_index('bank')

        rows = []
        for bank in soma_rows.index:
            r = soma_rows.loc[bank]
            rec_s  = float(r['recebidas_sophia'])
            rec_e  = float(r['recebidas_extrato'])
            pag_s  = abs(float(r['pagas_sophia']))
            pag_e  = abs(float(r['pagas_extrato']))
            rows.append({
                'Banco':             bank,
                'Rec. Sophia':       rec_s,
                'Rec. Extrato':      rec_e,
                'Dif. Recebidas':    round(rec_s - rec_e, 2),
                'Pag. Sophia':       pag_s,
                'Pag. Extrato':      pag_e,
                'Dif. Pagas':        round(pag_s - pag_e, 2),
                'Result. Sophia':    round(rec_s - pag_s, 2),
                'Result. Extrato':   round(rec_e - pag_e, 2),
            })

        totals = {
            'Banco':           'TOTAL',
            'Rec. Sophia':     sum(r['Rec. Sophia']    for r in rows),
            'Rec. Extrato':    sum(r['Rec. Extrato']   for r in rows),
            'Dif. Recebidas':  round(sum(r['Dif. Recebidas'] for r in rows), 2),
            'Pag. Sophia':     sum(r['Pag. Sophia']    for r in rows),
            'Pag. Extrato':    sum(r['Pag. Extrato']   for r in rows),
            'Dif. Pagas':      round(sum(r['Dif. Pagas']     for r in rows), 2),
            'Result. Sophia':  round(sum(r['Result. Sophia']  for r in rows), 2),
            'Result. Extrato': round(sum(r['Result. Extrato'] for r in rows), 2),
        }
        rows.append(totals)
        return rows

    def __write_resultados_gerais(self, writer, resumo_rows):
        workbook  = writer.book
        worksheet = workbook.add_worksheet("Resultados Gerais")
        writer.sheets["Resultados Gerais"] = worksheet

        hdr_fmt   = workbook.add_format({'bold': True, 'bg_color': '#B8CCE4', 'border': 1, 'align': 'center'})
        bank_fmt  = workbook.add_format({'bold': False, 'bg_color': '#EBF1DE', 'border': 1})
        num_fmt   = workbook.add_format({'bg_color': '#EBF1DE', 'border': 1, 'num_format': '#,##0.00'})
        diff_fmt  = workbook.add_format({'bg_color': '#EBF1DE', 'border': 1, 'num_format': '#,##0.00', 'font_color': '#C00000'})
        total_bank_fmt = workbook.add_format({'bold': True, 'bg_color': '#DCE6F1', 'border': 1})
        total_num_fmt  = workbook.add_format({'bold': True, 'bg_color': '#DCE6F1', 'border': 1, 'num_format': '#,##0.00'})
        total_diff_fmt = workbook.add_format({'bold': True, 'bg_color': '#DCE6F1', 'border': 1, 'num_format': '#,##0.00', 'font_color': '#C00000'})

        headers = ['Banco', 'Rec. Sophia', 'Rec. Extrato', 'Dif. Recebidas',
                   'Pag. Sophia', 'Pag. Extrato', 'Dif. Pagas',
                   'Result. Sophia', 'Result. Extrato']

        for col, h in enumerate(headers):
            worksheet.write(0, col, h, hdr_fmt)

        diff_cols = {headers.index('Dif. Recebidas'), headers.index('Dif. Pagas'),
                     headers.index('Result. Sophia'), headers.index('Result. Extrato')}

        for row_idx, row_data in enumerate(resumo_rows, start=1):
            is_total = row_data['Banco'] == 'TOTAL'
            for col, h in enumerate(headers):
                val = row_data[h]
                if col == 0:
                    worksheet.write(row_idx, col, val, total_bank_fmt if is_total else bank_fmt)
                elif col in diff_cols:
                    worksheet.write(row_idx, col, val, total_diff_fmt if is_total else diff_fmt)
                else:
                    worksheet.write(row_idx, col, val, total_num_fmt if is_total else num_fmt)

        col_widths = [24, 14, 14, 16, 14, 14, 12, 16, 16]
        for col, w in enumerate(col_widths):
            worksheet.set_column(col, col, w)

    def __build_consolidado_diario(self, final_df):
        daily = final_df[final_df['date'] != 'soma'].copy()
        for col in ['recebidas_sophia', 'recebidas_extrato', 'pagas_sophia', 'pagas_extrato']:
            daily[col] = pd.to_numeric(daily[col], errors='coerce').fillna(0)

        grouped = daily.groupby('date')[
            ['recebidas_sophia', 'recebidas_extrato', 'pagas_sophia', 'pagas_extrato']
        ].sum().reset_index().sort_values('date')

        rows = []
        for _, r in grouped.iterrows():
            rec_s = float(r['recebidas_sophia'])
            rec_e = float(r['recebidas_extrato'])
            pag_s = abs(float(r['pagas_sophia']))
            pag_e = abs(float(r['pagas_extrato']))
            rows.append({
                'Data':           r['date'],
                'Rec. Sophia':    rec_s,
                'Rec. Extrato':   rec_e,
                'Dif. Rec.':      round(rec_s - rec_e, 2),
                'Pag. Sophia':    pag_s,
                'Pag. Extrato':   pag_e,
                'Dif. Pag.':      round(pag_s - pag_e, 2),
                'Result. Sophia': round(rec_s - pag_s, 2),
                'Result. Extrato':round(rec_e - pag_e, 2),
            })

        rows.append({
            'Data':            'TOTAL',
            'Rec. Sophia':     round(sum(r['Rec. Sophia']     for r in rows), 2),
            'Rec. Extrato':    round(sum(r['Rec. Extrato']    for r in rows), 2),
            'Dif. Rec.':       round(sum(r['Dif. Rec.']       for r in rows), 2),
            'Pag. Sophia':     round(sum(r['Pag. Sophia']     for r in rows), 2),
            'Pag. Extrato':    round(sum(r['Pag. Extrato']    for r in rows), 2),
            'Dif. Pag.':       round(sum(r['Dif. Pag.']       for r in rows), 2),
            'Result. Sophia':  round(sum(r['Result. Sophia']  for r in rows), 2),
            'Result. Extrato': round(sum(r['Result. Extrato'] for r in rows), 2),
        })
        return rows

    def __write_consolidado_diario(self, writer, rows):
        workbook  = writer.book
        worksheet = workbook.add_worksheet("Consolidado Diário")
        writer.sheets["Consolidado Diário"] = worksheet

        hdr_fmt    = workbook.add_format({'bold': True, 'bg_color': '#B8CCE4', 'border': 1, 'align': 'center'})
        date_fmt   = workbook.add_format({'border': 1})
        num_fmt    = workbook.add_format({'border': 1, 'num_format': '#,##0.00'})
        diff_fmt   = workbook.add_format({'border': 1, 'num_format': '#,##0.00', 'font_color': '#C00000'})
        total_date_fmt = workbook.add_format({'bold': True, 'bg_color': '#DCE6F1', 'border': 1})
        total_num_fmt  = workbook.add_format({'bold': True, 'bg_color': '#DCE6F1', 'border': 1, 'num_format': '#,##0.00'})
        total_diff_fmt = workbook.add_format({'bold': True, 'bg_color': '#DCE6F1', 'border': 1, 'num_format': '#,##0.00', 'font_color': '#C00000'})

        headers = ['Data', 'Rec. Sophia', 'Rec. Extrato', 'Dif. Rec.',
                   'Pag. Sophia', 'Pag. Extrato', 'Dif. Pag.',
                   'Result. Sophia', 'Result. Extrato']
        diff_cols = {headers.index('Dif. Rec.'), headers.index('Dif. Pag.'),
                     headers.index('Result. Sophia'), headers.index('Result. Extrato')}

        for col, h in enumerate(headers):
            worksheet.write(0, col, h, hdr_fmt)

        for row_idx, row_data in enumerate(rows, start=1):
            is_total = row_data['Data'] == 'TOTAL'
            for col, h in enumerate(headers):
                val = row_data[h]
                if col == 0:
                    worksheet.write(row_idx, col, val, total_date_fmt if is_total else date_fmt)
                elif col in diff_cols:
                    worksheet.write(row_idx, col, val, total_diff_fmt if is_total else diff_fmt)
                else:
                    worksheet.write(row_idx, col, val, total_num_fmt if is_total else num_fmt)

        col_widths = [14, 14, 14, 12, 14, 14, 12, 16, 16]
        for col, w in enumerate(col_widths):
            worksheet.set_column(col, col, w)

    def __write_to_excel(self, final_df):
        resumo_rows     = self.__build_resumo(final_df)
        consolidado_rows = self.__build_consolidado_diario(final_df)

        with pd.ExcelWriter(os.path.join(os.getcwd(), 'comparativo_de_caixa.xlsx'), engine = "xlsxwriter" ) as writer:
            self.__write_resultados_gerais(writer, resumo_rows)
            self.__write_consolidado_diario(writer, consolidado_rows)
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