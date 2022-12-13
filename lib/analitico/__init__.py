import pyautogui

from lib.arquivo import read_excel


class ExtratoAnalitico:

    def __init__(self):
        self.interacao = {'tela': {'x': 601, 'y': 254}}

    def solicitar(self, base, pis):
        pyautogui.click(self.interacao['tela']['x'], self.interacao['tela']['y'])
        pyautogui.press('tab')
        pyautogui.write(base)
        pyautogui.press('tab')
        pyautogui.write(pis)
        pyautogui.press('tab', presses=2)
        pyautogui.press('enter')

    def confirmar(self):
        pyautogui.click(self.interacao['tela']['x'], self.interacao['tela']['y'])
        pyautogui.press('tab', presses=2)
        pyautogui.press('enter')

    def main(self, base, arquivo, window):
        try:
            nome, lista_pis, quantidade = read_excel(arquivo)
            nivel_progresso = 100 / quantidade
            progresso = 100 / quantidade

            for pis in lista_pis:
                self.solicitar(base, pis)
                pyautogui.sleep(0.5)
                self.confirmar()
                pyautogui.sleep(0.5)
                window['-STATUS-'].update('Processando...', 'blue')
                window['-PROGBAR-'].update(progresso % 101)
                progresso += nivel_progresso
            pyautogui.press('enter')
            window['-STATUS-'].update('Concluído com sucesso!', 'green')
        except FileNotFoundError:
            window['-STATUS-'].update('Anexar a planilha.', 'red')


if __name__ == '__main__':
    pass
