import PySimpleGUI as sg

from lib.analitico import ExtratoAnalitico
from lib.basico import ExtratoBasico
from lib.rescisorio import ExtratoRescisorio


class Interface:

    def __init__(self):
        self.extrato_analitico = ExtratoAnalitico()
        self.extrato_basico = ExtratoBasico()
        self.extrato_rescisorio = ExtratoRescisorio()

    @staticmethod
    def layout_base():
        layout = [[sg.I(k='-BASE-', s=(4, 1), justification='c', p=(4, 6))]]
        return layout

    @staticmethod
    def layout_arquivo():
        layout = [[sg.I(k='-ARQUIVO-', s=(37, 1)),
                   sg.FileBrowse('Excel', file_types=(('All files', '*.xlsx'),), target='-ARQUIVO-')]]
        return layout

    def layout_main(self):
        layout = [[sg.Frame('Base', self.layout_base(), vertical_alignment='top'),
                  sg.Frame('Arquivo', self.layout_arquivo(), vertical_alignment='top')],
                  [sg.B('Analítico', expand_x=True), sg.B('Básico', expand_x=True),
                   sg.B('Rescisório', expand_x=True), sg.B('Sair', expand_x=True)],
                  [sg.ProgressBar(100, k='-PROGBAR-', s=(35, 20))],
                  [sg.StatusBar(' ' * 100, k='-STATUS-', s=(35, 1))]]
        return layout

    def janela(self):
        return sg.Window('Extrato FGTS', self.layout_main(), keep_on_top=True, finalize=True, location=(1076, 372),
                         margins=(10, 10))

    def main(self):
        window = self.janela()
        analitico = 'Analítico'
        basico = 'Básico'
        rescisorio = 'Rescisório'
        sair = 'Sair'
        while True:
            event, values = window.read()
            if event in (sg.WINDOW_CLOSED, sair):
                break
            elif event == analitico:
                self.extrato_analitico.main(base=values['-BASE-'], arquivo=values['-ARQUIVO-'], window=window)
            elif event == basico:
                self.extrato_basico.main(base=values['-BASE-'], arquivo=values['-ARQUIVO-'], window=window)
            elif event == rescisorio:
                self.extrato_rescisorio.main(base=values['-BASE-'], arquivo=values['-ARQUIVO-'], window=window)
        window.close()


if __name__ == '__main__':
    pass
