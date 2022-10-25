import PySimpleGUI as sg

from lib.dctf_web import DCTFWeb


class Interface:

    planilha = 'DCTF_Web'
    col_nome = 'B'
    col_cnpj = 'D'
    col_status = 'G'
    col_inicio = 'H'
    col_fim = 'I'

    @staticmethod
    def input_col(chave, largura=5, just='c', default=None):
        """Padrão do input"""
        return sg.Input(k=chave, s=(largura, 1), justification=just, default_text=default)

    sg.set_options(font='_ 10')
    # sg.Print(font='_ 12', keep_on_top=True, size=(45, 20), location=(1461, 102))

    @staticmethod
    def layout_arquivo():
        """Layout que possibilita pegar o arquivo Excel"""
        return [[sg.Input(k='-ARQUIVO-'), sg.FilesBrowse('Excel', target='-ARQUIVO-')]]

    @staticmethod
    def layout_competencia():
        """Layout que possibilita a escolha do mês e ano"""
        layout = [[sg.Input(k='-MES-', justification='c', s=(9, 1), disabled=True),
                   sg.CalendarButton('Mês', target='-MES-', format='%m-%Y', no_titlebar=False, )]]
        return layout

    def layout_planilha(self):
        """Layout que identifica o nome da planilha ativada"""
        layout = [[self.input_col('-PLANILHA-', largura=26, just='l', default=Interface.planilha)]]
        return layout

    def layout_coluna_consulta(self):
        """Layout onde esta a coluna Nome e CNPJ"""
        layout = [[sg.T('Nome'), self.input_col('-NOME-', default=Interface.col_nome),
                   sg.T('CNPJ'), self.input_col('-CNPJ-', default=Interface.col_cnpj)]]
        return layout

    def layout_coluna_confirmacao(self):
        """Layout onde esta a coluna Status, Inicio e Fim para dar baixa no processo"""
        layout = [[sg.T('Status'), self.input_col('-STATUS-', default=Interface.col_status),
                   sg.T('Início'), self.input_col('-INICIO-', default=Interface.col_inicio),
                   sg.T('Fim'), self.input_col('-FIM-', default=Interface.col_fim)]]
        return layout

    @staticmethod
    def layout_titular():
        """Layout específica para a empresa titular"""
        layout = [[sg.Input(k='-TITULAR-'), sg.Button('Titular')]]
        return layout

    def layout_empresas(self):
        """Layout que encapsula os layouts para processo das empresas constantes na planilha"""
        layout = [[sg.Frame('Planilha', self.layout_planilha())],
                  [sg.Frame('Coluna Consulta', self.layout_coluna_consulta())],
                  [sg.Frame('Coluna Confirmação', self.layout_coluna_confirmacao())],
                  [sg.Frame('Arquivo', self.layout_arquivo())]]
        return layout

    def layout_main(self):
        layout = [[sg.Frame('Competência', self.layout_competencia())],
                  [sg.Frame('Empresas', self.layout_empresas())],
                  [sg.Button('Processar'), sg.Button('Cancelar')],
                  [sg.T('')],
                  [sg.HSeparator()],
                  [sg.Frame('Empresa Titular', self.layout_titular())],
                  [sg.T('')]]
        return layout

    def main(self):
        window = sg.Window('DCTF Web', self.layout_main(), keep_on_top=True, finalize=True)
        dctf_web = DCTFWeb()

        while True:
            event, values = window.read()
            if event in (sg.WINDOW_CLOSED, 'Cancelar'):
                break
            elif event == 'Processar':
                dctf_web.main(
                    arq=values['-ARQUIVO-'],
                    competencia=values['-MES-'],
                    planilha=values['-PLANILHA-'],
                    coluna_nome=values['-NOME-'],
                    coluna_cnpj=values['-CNPJ-'],
                    coluna_status=values['-STATUS-'],
                    coluna_inicio=values['-INICIO-'],
                    coluna_fim=values['-FIM-'],
                )
            elif event == 'Titular':
                dctf_web.empresa_titular(nome=values['-TITULAR-'], mes=values['-MES-'])

            # sg.Print(f'event => {event}', colors='white on green', erase_all=True)
            # sg.Print(*[f'    {k} => {values[k]}' for k in values], sep='\n')

        window.close()


if __name__ == '__main__':
    pass
