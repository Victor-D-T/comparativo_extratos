import pandas as pd
import os

class GenerateCashFlow:

    SYNTHETIC_ACCOUNTS = [
        ("100.000",       "RECEITAS",                        ["1"]),
        ("110.000",       "MENSALIDADES",                    ["101","112","113","114"]),
        ("120.000",       "OUTRAS RECEITAS",                 ["121","122","123","124","125","126","127","128"]),
        ("121.000",       "ALUGUEL",                         ["121"]),
        ("122.000",       "MAT. DIDÁTICO/PEDÁGOGICO",        ["122"]),
        ("123.000",       "CURSOS LIVRES",                   ["123"]),
        ("124.000",       "REVENDAS",                        ["124"]),
        ("125.000",       "EVENTOS",                         ["125"]),
        ("126.000",       "FINANCEIRAS",                     ["126"]),
        ("127.000",       "BANCÁRIA",                        ["127"]),
        ("128.000",       "PARCERIAS",                       ["128"]),
        ("130.000",       "ACORDOS JUDICIAIS",               ["130"]),
        ("200.000",       "DESPESAS",                        ["2"]),
        ("201.000",       "SERVIÇOS BÁSICOS",                ["201"]),
        ("202.000",       "DESPESAS GERAIS",                 ["202"]),
        ("2020103 e 104", "Apostilas Anglo e Complementares",["202103","202104"]),
        ("203.000",       "SERVIÇOS CONTRATADOS",            ["203"]),
        ("204.000",       "SALÁRIOS E ENCARGOS",             ["204"]),
        ("205.000",       "BENEFÍCIOS",                      ["205"]),
        ("206.000",       "ESTAGIARIOS",                     ["206"]),
        ("207.000",       "FINANCEIRAS",                     ["207"]),
        ("212.000",       "TRIBUTÁRIAS",                     ["212"]),
        ("214.000",       "SEGUROS",                         ["214"]),
        ("216.000",       "REMATRICULA",                     ["216"]),
        ("217.000",       "SÓCIAS",                          ["217"]),
        ("220.000",       "CUSTO REVENDAS DIVERSAS",         ["220"]),
        ("230.000",       "EVENTOS",                         ["230","231"]),
        ("240.000",       "IMOBILIZADO",                     ["240"]),
    ]

    def __init__(self) -> None:
        self.sophia_folder_path = os.path.join(os.getcwd(),'Sophia')
        self.final_dict = None
        _empty = pd.DataFrame(columns=["Conta", "Descricao_conta", "Valor"])
        self.pagas     = _empty.copy()
        self.recebidas = _empty.copy()

    def main(self, final_dict=None):
        self.sophias_cash_flow()
        self.create_excel(final_dict)

        return self.final_dict

    def extract_sophias_transactions_data(self, file):
        engine = 'openpyxl' if file.lower().endswith('.xlsx') else 'xlrd'
        df = pd.read_excel(file, engine=engine)
        if df.get("CLASSIFIC_COD") is not None:
            df = df[["CLASSIFIC_COD", "CLASSIFIC_DESC", "VALOR_CLASS"]]
            new_pagas = df.groupby(["CLASSIFIC_COD", "CLASSIFIC_DESC"])["VALOR_CLASS"].sum().reset_index()
            new_pagas = new_pagas.rename(columns={"CLASSIFIC_COD": "Conta", "CLASSIFIC_DESC": "Descricao_conta", "VALOR_CLASS": "Valor"}).reset_index(drop=True)
            self.pagas = pd.concat([self.pagas, new_pagas], ignore_index=True)

        elif df.get("PLANO_CONTAS") is not None:
            df = df[["PLANO_CONTAS", "PGTO_CLASSFIC"]]
            df[["Conta", "Descricao_conta"]] = df["PLANO_CONTAS"].str.split(" - ", n=1, expand=True)
            new_recebidas = df.groupby(["Conta", "Descricao_conta"])["PGTO_CLASSFIC"].sum().reset_index()
            new_recebidas = new_recebidas.rename(columns={"PGTO_CLASSFIC": "Valor"}).reset_index(drop=True)
            self.recebidas = pd.concat([self.recebidas, new_recebidas], ignore_index=True)

    @staticmethod
    def _normalize(code: str) -> str:
        return str(code).replace(".", "")

    def _find_parent(self, account_code: str) -> str:
        norm = self._normalize(account_code)
        best_code, best_len, best_idx = None, 0, -1
        for idx, (code, desc, prefixes) in enumerate(self.SYNTHETIC_ACCOUNTS):
            for p in prefixes:
                if norm.startswith(p) and (len(p) > best_len or (len(p) == best_len and idx > best_idx)):
                    best_code, best_len, best_idx = code, len(p), idx
        return best_code

    def create_excel(self, final_dict=None):
        analytical = pd.concat([self.pagas, self.recebidas], ignore_index=True).reset_index(drop=True)

        # Pre-compute parent for each analytical row
        analytical["_parent"] = analytical["Conta"].apply(self._find_parent)

        # Pre-compute synthetic totals
        synth_totals = {}
        for code, desc, prefixes in self.SYNTHETIC_ACCOUNTS:
            mask = analytical["Conta"].apply(
                lambda c: any(self._normalize(c).startswith(p) for p in prefixes)
            )
            synth_totals[code] = analytical.loc[mask, "Valor"].sum()

        # Build ordered output rows: (is_synthetic, Conta, Descricao_conta, Valor)
        output_rows = []
        for code, desc, prefixes in self.SYNTHETIC_ACCOUNTS:
            output_rows.append((True, code, desc, synth_totals[code]))
            children = analytical[analytical["_parent"] == code].sort_values("Conta")
            for _, row in children.iterrows():
                output_rows.append((False, row["Conta"], row["Descricao_conta"], row["Valor"]))

        # Write to Excel with per-row formatting
        output_path = os.path.join(os.getcwd(), "fluxo_de_caixa.xlsx")
        with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
            workbook = writer.book
            worksheet = workbook.add_worksheet("fluxo de caixa")
            writer.sheets["fluxo de caixa"] = worksheet

            synth_fmt       = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1})
            synth_num_fmt   = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1, 'num_format': '#,##0.00'})
            anal_fmt        = workbook.add_format({'border': 1})
            anal_num_fmt    = workbook.add_format({'border': 1, 'num_format': '#,##0.00'})
            header_fmt      = workbook.add_format({'bold': True, 'bg_color': '#B8CCE4', 'border': 1})
            conf_header_fmt = workbook.add_format({'bold': True, 'bg_color': '#F4B942', 'border': 1})
            conf_num_fmt    = workbook.add_format({'bold': True, 'bg_color': '#FFF2CC', 'border': 1, 'num_format': '#,##0.00'})
            conf_diff_fmt   = workbook.add_format({'bold': True, 'bg_color': '#FFF2CC', 'border': 1, 'num_format': '#,##0.00', 'font_color': '#FF0000'})

            # Header row
            for col, name in enumerate(["Conta", "Descricao_conta", "Valor"]):
                worksheet.write(0, col, name, header_fmt)

            worksheet.set_column(0, 0, 18)
            worksheet.set_column(1, 1, 45)
            worksheet.set_column(2, 2, 16)
            worksheet.set_column(3, 3, 16)
            worksheet.set_column(4, 4, 16)

            for row_idx, (is_synth, conta, desc, valor) in enumerate(output_rows, start=1):
                text_fmt = synth_fmt if is_synth else anal_fmt
                num_fmt  = synth_num_fmt if is_synth else anal_num_fmt
                worksheet.write(row_idx, 0, conta, text_fmt)
                worksheet.write(row_idx, 1, desc,  text_fmt)
                worksheet.write(row_idx, 2, valor, num_fmt)

            if final_dict:
                total_rec_extrato = sum(
                    v for bd in final_dict.values()
                    for v in bd['recebidas']['extrato'].values()
                )
                total_pag_extrato = abs(sum(
                    v for bd in final_dict.values()
                    for v in bd['pagas']['extrato'].values()
                ))
                sophia_rec = synth_totals.get("100.000", 0)
                sophia_pag = synth_totals.get("200.000", 0)

                next_row = len(output_rows) + 2
                for col, h in enumerate(["Conta", "Descrição", "Sophia", "Extrato", "Diferença"]):
                    worksheet.write(next_row, col, h, conf_header_fmt)

                for i, (code, label, sophia_val, extrato_val) in enumerate([
                    ("100.000", "RECEITAS", sophia_rec, total_rec_extrato),
                    ("200.000", "DESPESAS", sophia_pag, total_pag_extrato),
                ]):
                    r = next_row + 1 + i
                    worksheet.write(r, 0, code,                               conf_num_fmt)
                    worksheet.write(r, 1, label,                              conf_num_fmt)
                    worksheet.write(r, 2, sophia_val,                         conf_num_fmt)
                    worksheet.write(r, 3, extrato_val,                        conf_num_fmt)
                    worksheet.write(r, 4, round(sophia_val - extrato_val, 2), conf_diff_fmt)

    def sophias_cash_flow(self):
        files = []

        for f in os.listdir(self.sophia_folder_path):
            complete_file_path = os.path.join(self.sophia_folder_path, f)
            files.append(complete_file_path)

        for file in files:
            if not file.lower().endswith(('.xls', '.xlsx')):
                continue
            self.extract_sophias_transactions_data(file)

        if not self.pagas.empty:
            self.pagas = self.pagas.groupby(["Conta", "Descricao_conta"])["Valor"].sum().reset_index()
        if not self.recebidas.empty:
            self.recebidas = self.recebidas.groupby(["Conta", "Descricao_conta"])["Valor"].sum().reset_index()

        return self.final_dict
