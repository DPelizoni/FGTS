import pyautogui

from lib.arquivo import read_excel

pyautogui.FAILSAFE = True


class ExtratoRescisorio:

    def __init__(self):
        self.interacao = {'tela': {'x': 476, 'y': 263}}

    @staticmethod
    def solicitar(pis):
        pyautogui.write(pis)
        pyautogui.press('tab')
        pyautogui.press('space')
        pyautogui.hotkey('shift', 'tab')

    def base_conta(self, base):
        pyautogui.click(self.interacao['tela']['x'], self.interacao['tela']['y'])
        pyautogui.press('tab')
        pyautogui.write(base)
        pyautogui.press('tab')

    def main(self, base, arquivo, window):
        try:
            nome, lista_pis, quantidade = read_excel(arquivo)
            nivel_progresso = 100 / quantidade
            progresso = 100 / quantidade
            self.base_conta(base)

            for pis in lista_pis:
                self.solicitar(pis)
                window['-STATUS-'].update('Processando...', 'blue')
                window['-PROGBAR-'].update(progresso % 101)
                progresso += nivel_progresso
            pyautogui.press('enter')
            window['-STATUS-'].update('Concluído com sucesso!', 'green')
        except FileNotFoundError:
            window['-STATUS-'].update('Anexar a planilha.', 'red')


if __name__ == '__main__':
    pass
