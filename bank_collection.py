from read_extrato_cora import ReadExtratoCora
from read_extrato_itau import ReadExtratoItau
from read_extrato_sicredi import ReadExtratoSicredi
from read_extrato_caixa import ReadExtratoCaixa


class BankCollection:
    def __init__(self, file, bank, final_dict) -> None:
        self.file = file
        self.bank = bank
        self.final_dict = final_dict

    def main(self):
        if self.bank in ["cora unidade 3", "cora unidade 1 e 2"]:
            read_extrato_cora = ReadExtratoCora(self.file, self.final_dict)
            self.final_dict = read_extrato_cora.read_extrato(self.bank)

        elif self.bank in ["itau unidade 1 e 2",  "itau unidade 3"]:
            read_extrato_itau = ReadExtratoItau(self.file, self.final_dict)
            self.final_dict = read_extrato_itau.read_extrato(self.bank)

        elif self.bank in ["sicredi unidade 1 e 2",  "sicredi unidade 3"]:
            read_extrato_sicredi = ReadExtratoSicredi(self.file, self.final_dict)
            self.final_dict = read_extrato_sicredi.read_extrato(self.bank)

        elif self.bank in ["caixa unidade 1 e 2",  "caixa unidade 3"]:
            read_extrato_caixa = ReadExtratoCaixa(self.file, self.final_dict)
            self.final_dict = read_extrato_caixa.read_extrato(self.bank)

        else:
            raise Exception("no file")
        
        return self.final_dict
        