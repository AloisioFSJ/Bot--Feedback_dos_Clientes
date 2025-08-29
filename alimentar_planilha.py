from datetime import datetime
import pandas as pd
import pyautogui as pt
import pyperclip as pp
import time

repeticao = int(input('Quantas páginas de pedidos? - '))
cel = str(input('Em qual celula começar? - '))

print("""
    Padrão para inicar o bot:
      NO WINDOWS TAB:
      1ª - Chrome
      2ª - Excel

""")

input('Pressione "Enter" quando puder iniciar a automação... ')

# MUDAR PARA O CHROME =======================================
pt.PAUSE = 0.25
pt.hotkey('ctrl', 'win', 'right')
pt.hotkey('alt', 'tab')
pt.click(x=66, y=202)

# COPIAR OS DADOS DA TELA ===================================
pt.PAUSE = 0.5
pt.write(cel)
pt.press('enter')
pt.hotkey('alt', 'tab')

for i in range(repeticao):
    pt.click(x=1554, y=972)
    pt.scroll(200)
    
    # COPIAR FEEDBACKS DOS CLIENTES
    pt.moveTo(x=498, y=240)
    pt.mouseDown(button='left')
    pt.moveTo(x=1707, y=987, duration=0.5)

    # SALVAR NA AREA DE TRANFERENCIA
    pt.hotkey('ctrl', 'c')
    time.sleep(0.5)
    feedback = pp.paste()

    # COLAR NO ARQUIVO EM EXCEL
    pt.keyDown('alt')
    pt.press('tab')
    pt.keyUp('alt')

    pt.hotkey('ctrl', 'shift', 'v')
    pt.keyDown('ctrl')
    pt.press('down', presses=2)
    pt.press('up')
    pt.keyUp('ctrl')
    pt.press('down')
    
    # VOLTAR PARA O CHROME
    pt.keyDown('Alt')
    pt.press('Tab')
    pt.keyUp('Alt')

    # CLICAR NA PROXIMO PAGINA
    pt.scroll(-800)
    time.sleep(1)
    seta_vermelha = pt.locateCenterOnScreen('C:\_Codes\Feedback dos Clientes\Imagens\seta_vermelha.png', confidence=0.9)
    pt.moveTo(seta_vermelha, duration=1)
    pt.click()

    time.sleep(3)

pt.hotkey('ctrl', 'win', 'left')
print('✅ Bot finalizado, e Feedback copiado')
time.sleep(3)
