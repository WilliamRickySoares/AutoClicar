import pyautogui
import time

print("Aguandando comando...")
time.sleep(5)
posicao_x = 1000
posicao_y = 450

posicao_x_play = 500
posicao_y_play = 450

x_fechar_janela = 1250
y_fechar_janela = 15


def clicar(x, y, t=0.5, q=True):
    print("Iniciando...")
    print("Movendo...")
    pyautogui.moveTo(x, y, duration = 0.5)
    print("Moveu")
    time.sleep(0.5)
    if q:
        pyautogui.click()
        print("clicou")
    print("Aguardando...")
    time.sleep(t)
    print("Encerrado!")


def minimizar():
    clicar(x_fechar_janela, y_fechar_janela, q = True, t = 1)


minimizar()

for a in range(0, 16):
    clicar(posicao_x_play, posicao_y_play, t=(60*60))
    clicar(posicao_x, posicao_y, t=3)
minimizar()
