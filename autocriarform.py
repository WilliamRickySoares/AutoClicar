import pyautogui
import pyperclip
import time
import pandas as pd


def mover_clicar(x, y, clicar=True):
    pyautogui.moveTo(x, y, duration = 0.3)
    if clicar:
        pyautogui.click()


def copiar_colar(texto):
    pyperclip.copy(f'{texto}')
    pyautogui.hotkey("ctrl", "v")


def novo_form():
    """Clicar em Novo Responsivo (DEV)"""
    time.sleep(1)
    x_opcao_novo = 340
    y_opcao_novo = 230
    mover_clicar(x_opcao_novo, y_opcao_novo)
    x_novo_resp = 340
    y_novo_resp = 300
    mover_clicar(x_novo_resp, y_novo_resp)


def criar_form(id_form_, nome_form_, id_tab_):
    tipo_form = 'rh'
    nome_arquivo = f'{id_form_}.seformresp'
    novo_form()

    time.sleep(0.2)
    x_identificador = 450
    y_identificador = 285
    mover_clicar(x_identificador, y_identificador)
    copiar_colar(id_form_)  # ID formulario

    time.sleep(0.2)
    x_nome_form = 650
    y_nome_form = 285
    mover_clicar(x_nome_form, y_nome_form)
    pyperclip.copy(nome_form_)  # Nome formulario
    pyautogui.hotkey("ctrl", "v")

    """Clicar no tipo de form"""
    time.sleep(0.2)
    x_tipo_form = 450
    y_tipo_form = 350
    mover_clicar(x_tipo_form, y_tipo_form)
    pyautogui.click()
    time.sleep(0.2)
    pyperclip.copy(tipo_form)
    pyautogui.hotkey("ctrl", "v")

    """Clicar na opção selecionada"""
    time.sleep(0.2)
    x_selc_tab = 450
    y_selc_tab = 400
    pyautogui.moveTo(x_selc_tab+20, y_selc_tab, duration = 1)
    mover_clicar(x_selc_tab, y_selc_tab)

    """Procurar tabela"""
    time.sleep(0.2)
    x_proc_tab = 450
    y_proc_tab = 550
    mover_clicar(x_proc_tab, y_proc_tab)
    pyperclip.copy(id_tab_)  # Nome da tabela
    pyautogui.hotkey("ctrl", "v")

    """Selecionar tabela procurada"""
    time.sleep(0.3)
    x_clic_tab = 450
    y_clic_tab = 500
    pyautogui.moveTo(x_clic_tab+20, y_clic_tab, duration = 1)
    mover_clicar(x_clic_tab, y_clic_tab)
    pyautogui.click()

    """Salvar e Fechar o criador de form"""
    time.sleep(0.3)
    x_save_close = 850
    y_save_close = 620
    mover_clicar(x_save_close, y_save_close)

    """Clicar em "Importar" (Cinza)"""
    time.sleep(3.5)
    x_import = 1280
    y_import = 150
    mover_clicar(x_import, y_import)

    """Seleciona um arquivo e digita na procura"""
    time.sleep(0.2)
    pyperclip.copy(nome_arquivo)
    x_select_arquivo = 600
    y_select_arquivo = 230
    mover_clicar(x_select_arquivo, y_select_arquivo)
    time.sleep(0.7)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.5)
    pyautogui.hotkey("enter")
    time.sleep(0.2)

    # Clicar em "Importar" (Azul)
    time.sleep(0.2)
    x_clic_import = 900
    y_clic_import = 400
    mover_clicar(x_clic_import, y_clic_import)

    """Encerrar importação"""
    time.sleep(8)
    x_finalizar_import = 1000
    y_finalizar_import = 400
    mover_clicar(x_finalizar_import, y_finalizar_import)
    time.sleep(1)
    mover_clicar(x_finalizar_import, y_finalizar_import)

    # """ Minimizar a janela"""
    # x_close_nav = 1250
    # y_close_nav = 10
    # mover_clicar(x_close_nav, y_close_nav)
    # time.sleep(1)

    """ Fechar a janela"""
    x_close_nav = 1320
    y_close_nav = 10
    mover_clicar(x_close_nav, y_close_nav)
    time.sleep(1)


arquivo = "listagem_form.xls"
df = pd.read_excel(arquivo)
df = df.dropna(how='all')
print(f"Total de linhas: {len(df)}")

for index, row in df.iterrows():
    id_form = str(f"{row.tolist()[0]}").rstrip().lstrip()
    nome_form = str(f"{row.tolist()[1]}").rstrip().lstrip()
    id_tab = str(f"{row.tolist()[2]}").rstrip().lstrip()
    print(nome_form)
    criar_form(id_form, nome_form, id_tab)
