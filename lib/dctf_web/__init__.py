import pyautogui as pag

from lib.planilha import Planilha


class DCTFWeb:
    alterar_perfil_x = 729
    alterar_perfil_y = 217
    alterar_CNPJ_x = 332
    alterar_CNPJ_y = 520
    btn_alterar_perfil_x = 580
    btn_alterar_perfil_y = 522
    visualizar_recibo_x = 910
    visualizar_recibo_y = 631
    visualizar_guia_x = 832
    visualizar_guia_y = 640
    salvar_guia_x = 832
    salvar_guia_y = 640
    titular_vis_recibo_x = 911
    titular_vis_recibo_y = 694
    titular_vis_guia_x = 836
    titular_vis_guia_y = 696
    titular_salvar_guia_x = 830
    titular_salvar_guia_y = 642
    duration = 0.5

    def __init__(self):
        pass

    @staticmethod
    def click(x, y):
        """Simula a opção click do mouse"""
        pag.click(x, y, duration=DCTFWeb.duration)

    @staticmethod
    def escrever(text):
        """Será escrita a informação passada"""
        return pag.typewrite(text)

    @staticmethod
    def teclar(tecla):
        """Será precionada a tecla informada"""
        pag.press(tecla)

    @staticmethod
    def esperar(tempo=2):
        """Tempo de espera para uma nova ação"""
        pag.sleep(tempo)

    def alterar_perfil_acesso(self):
        """Clica na opção para alteração de acesso"""
        self.click(DCTFWeb.alterar_perfil_x, DCTFWeb.alterar_perfil_y)

    def alterar_perfil_cnpj(self, cnpj):
        """Clica no campo informando o CNPJ"""
        self.click(DCTFWeb.alterar_CNPJ_x, DCTFWeb.alterar_CNPJ_y)
        self.escrever(cnpj)

    def alterar_perfil_btn(self):
        """Clica para que seja efetuada a troca do perfil"""
        self.click(DCTFWeb.btn_alterar_perfil_x, DCTFWeb.btn_alterar_perfil_y)

    def acessar_perfil_acesso(self, cnpj):
        """Rotina de 3 passos para acessar a alteração do perfil"""
        self.alterar_perfil_acesso()
        self.esperar()
        self.alterar_perfil_cnpj(cnpj)
        self.esperar(1)
        self.alterar_perfil_btn()

    def visualizar_recibo(self, x, y):
        """Cordenadas para empresas com procuração"""
        self.click(x, y)

    def salvar_recibo(self, nome, mes):
        """Salva o arquivo com o nome, a descrição e o mês de referência"""
        pag.typewrite(f'{nome}_DCTF Web_Recibo_{mes}')
        self.teclar('enter')

    def visualizar_guia(self, x, y):
        """Clica para visualizar a guia"""
        self.click(x, y)

    def salvar_guia(self, nome, mes, x, y):
        """Salva o arquivo com o nome, a descrição e o mês de referência"""
        self.click(x, y)
        self.esperar()
        pag.typewrite(f'{nome}_DCTF Web_Guia_{mes}')
        self.esperar()
        self.teclar('enter')
        self.esperar()
        self.teclar('enter')

    def empresa(self, nome, mes, cnpj):
        """Seleciona o CNPJ para visualizar o recibo e emitir a guia"""
        self.acessar_perfil_acesso(cnpj)
        self.esperar()
        self.visualizar_recibo(DCTFWeb.visualizar_recibo_x, DCTFWeb.visualizar_recibo_y)
        self.esperar()
        self.salvar_recibo(nome, mes)
        self.esperar()
        self.visualizar_guia(DCTFWeb.visualizar_guia_x, DCTFWeb.visualizar_guia_y)
        self.esperar()
        self.salvar_guia(nome, mes, DCTFWeb.salvar_guia_x, DCTFWeb.salvar_guia_y)
        self.esperar()

    def empresa_titular(self, nome, mes):
        """Empresa do certificado digital, visualiza o recibo e emiti a guia"""
        self.visualizar_recibo(DCTFWeb.titular_vis_recibo_x, DCTFWeb.titular_vis_recibo_y)
        self.esperar()
        self.salvar_recibo(nome, mes)
        self.esperar()
        self.visualizar_guia(DCTFWeb.titular_vis_guia_x, DCTFWeb.titular_vis_guia_y)
        self.esperar()
        self.salvar_guia(nome, mes, DCTFWeb.titular_salvar_guia_x, DCTFWeb.titular_salvar_guia_y)
        self.esperar()

    def main(self, competencia, arq, planilha, coluna_nome, coluna_cnpj, coluna_status, coluna_inicio, coluna_fim):
        """Realiza o processo com base na planilha, ao término dá baixa informando o tempo inicial e final"""
        planilha = Planilha(
            mes=competencia,
            arquivo=arq,
            planilha=planilha,
            col_nome=coluna_nome,
            col_cnpj=coluna_cnpj,
            col_status=coluna_status,
            col_inicio=coluna_inicio,
            col_fim=coluna_fim,
        )

        for linha in planilha.parametro_listagem():
            planilha.inicio_processo(linha)
            self.empresa(nome=planilha.razao_social(linha), cnpj=planilha.cnpj(linha), mes=competencia)
            planilha.confirmacao_ok(linha)
            planilha.fim_processo(linha)
            planilha.salvar_planilha()


if __name__ == '__main__':
    pass
