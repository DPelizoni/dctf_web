import os
from datetime import datetime

from openpyxl import load_workbook
from tqdm import tqdm


class Planilha:

    def __init__(self, mes, arquivo, planilha, col_nome, col_cnpj, col_status, col_inicio, col_fim):
        self.mes = mes
        self.arquivo = arquivo
        self.excel = load_workbook(self.arquivo)
        self.planilha = self.excel[planilha]
        self.col_nome = col_nome
        self.col_cnpj = col_cnpj
        self.col_status = col_status
        self.col_inicio = col_inicio
        self.col_fim = col_fim

    @staticmethod
    def tempo():
        """Marcação do tempo atual, para medição do processo"""
        return datetime.now().strftime('%d/%m/%Y %H:%M:%S')

    @staticmethod
    def cnpj_completo(cnpj):
        """Tamanho do CNPJ 14 dígitos, completa com 0 no início o CNPJ com quantidade inferior"""
        tamanho = len(str(cnpj))

        if tamanho < 14:
            diferenca = 14 - tamanho
            complemento = '0' * diferenca
            return f'{complemento}{cnpj}'
        else:
            return cnpj

    def coluna(self, coluna):
        """Coluna da planilha ativada"""
        return self.planilha[coluna]

    def celula(self, celula_col, celula_lin):
        """Mostra a referência da celula"""
        return self.planilha[f'{celula_col}{celula_lin}']

    def celula_valor(self, celula_col, celula_lin):
        """Mostra o valor contido na celula"""
        return self.planilha[f'{celula_col}{celula_lin}'].value

    def quantidade_preenchidas(self, coluna):
        """Totaliza quantidade de células preenchidas da coluna informada"""
        return len(self.planilha[coluna])

    def quantidade_status_ok(self):
        """Totaliza a quantidade de células marcadas com Ok"""
        contagem = 0
        for celula in self.planilha[self.col_status]:
            if celula.value == 'Ok':
                contagem += 1
        return contagem

    def preencher_celula(self, celula_col, celula_lin, valor):
        """Coloca o valor na célula informada"""
        celula = self.planilha
        celula[f'{celula_col}{celula_lin}'] = valor
        return celula

    def parametro_listagem(self):
        """Base dos processos a serem realizados"""
        status = self.quantidade_status_ok() + 2
        lista_cnpj = self.quantidade_preenchidas(self.col_cnpj) + 1
        return tqdm(range(status, lista_cnpj))

    def razao_social(self, linha):
        """Busca a razao social das células"""
        return self.celula_valor(self.col_nome, linha)

    def cnpj(self, linha):
        """Busca número do CNPJ das células"""
        return self.cnpj_completo(str(self.celula_valor(self.col_cnpj, linha)))

    def inicio_processo(self, linha):
        """Marcação inicial do processo"""
        return self.preencher_celula(self.col_inicio, linha, self.tempo())

    def fim_processo(self, linha):
        """Marcação final do processo"""
        return self.preencher_celula(self.col_fim, linha, self.tempo())

    def confirmacao_ok(self, linha):
        """Marcação Ok do processo concluído"""
        return self.preencher_celula(self.col_status, linha, 'Ok')

    def salvar_planilha(self):
        return self.excel.save(filename=self.arquivo)

    def abrir_planilha(self):
        os.startfile(self.arquivo)


if __name__ == '__main__':
    pass
