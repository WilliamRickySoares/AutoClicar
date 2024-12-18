import pyautogui
import time

print("Iniciando...")
time.sleep(1)

# Check-box para selecionar todos os itens
check_all_x = 310
check_all_y = 283

# Terceiro icone da barra
terceiro_icon_x = 380
terceiro_icon_y = 245

# Pesquisar
pesquisar_x = 200
pesquisar_y = 590

try:
    while True:
        pyautogui.moveTo(check_all_x, check_all_y, duration = 0.5)
        pyautogui.click()
        time.sleep(0.5)

        pyautogui.moveTo(terceiro_icon_x, terceiro_icon_y, duration = 0.5)
        pyautogui.click()
        time.sleep(0.5)

        pyautogui.press('space')

        pyautogui.moveTo(pesquisar_x, pesquisar_y, duration = 0.5)
        time.sleep(5)
        pyautogui.click()
        time.sleep(0.5)

except KeyboardInterrupt:
    print("\nAutomação encerrada pelo usuário.")
