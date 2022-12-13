import pyautogui

from lib.arquivo import read_excel

pyautogui.FAILSAFE = True


class ExtratoBasico:

    def __init__(self):
        self.interacao = {'tela': {'x': 598, 'y': 275}}

    def localizar_trabalhador(self, base, pis):
        pyautogui.click(self.interacao['tela']['x'], self.interacao['tela']['y'])
        pyautogui.press('tab')
        pyautogui.write(base)
        pyautogui.press('tab', presses=2, interval=0.2)
        pyautogui.write(pis)
        pyautogui.press('enter')

    def salvar_arquivo(self, nome):
        pyautogui.sleep(0.5)
        pyautogui.hotkey('ctrl', 'p')
        pyautogui.sleep(1)
        pyautogui.press('enter')
        pyautogui.sleep(1)
        pyautogui.write(f'{nome}_Extrato FGTS')
        pyautogui.press('enter')
        pyautogui.sleep(0.5)
        pyautogui.click(self.interacao['tela']['x'], self.interacao['tela']['y'])
        pyautogui.press('tab')
        pyautogui.press('enter')
        pyautogui.sleep(1)

    def main(self, base, arquivo, window):
        try:
            lista_nome, lista_pis, quantidade = read_excel(arquivo)
            nivel_progresso = 100 / quantidade
            progresso = 100 / quantidade

            for i in range(quantidade):
                self.localizar_trabalhador(base, lista_pis[i])
                self.salvar_arquivo(lista_nome[i])
                window['-STATUS-'].update('Processando...', 'blue')
                window['-PROGBAR-'].update(progresso % 101)
                progresso += nivel_progresso
            window['-STATUS-'].update('Concluído com sucesso!', 'green')
        except FileNotFoundError:
            window['-STATUS-'].update('Anexar a planilha.', 'red')


if __name__ == '__main__':
    pass
