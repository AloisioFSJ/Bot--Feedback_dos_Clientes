import pyautogui as pt
import pyperclip as pp
from datetime import datetime
import time
import os

repeticao = int(input('Quantas páginas de pedidos? - '))
data_inicio = str(input('Qual a data de INICIO? (00/00/0000) - '))
data_final = datetime.now().strftime('%d/%m')
cel = str(input('Em qual celula começar a alimentar na planilha? - '))

print("""Padrão das Abas:
      
            (Atrás do cmd)
      ABA_01 - Chrome com a pagina do Ifood
      ABA_02 - Excel com a planilha aberta
""")

input('Pressione "Enter" quando puder começar... ')

pt.PAUSE = 0.5

# ABRIR PAGINA DE AVALIACOES ===================================================================
pt.keyDown('Alt')
pt.press('Tab')
pt.keyUp('Alt')
circulo_data = pt.locateCenterOnScreen('imagens\_avaliacoes.png', confidence=0.7) # ENCONTRAR O ICONE NA PAGINA
pt.click() # CLICA
time.sleep(1)
circulo_data = pt.locateCenterOnScreen('imagens\_filtros.png', confidence=0.7) # ENCONTRAR O ICONE DOS FILTROS
pt.click() # CLICA
time.sleep(1)

#FILTROS =======================================================================================
pt.click(x=1756, y=613) # MOSTRAR FILTROS

pt.click(x=550, y=764) # AVALIACAO
pt.click(x=501, y=825) # AVALIACAO

pt.tripleClick(x=1025, y=765) # MOVER PARA "DE" & SELECIONAR O CAMPO COMPLETO
pt.press('Delete')
pt.write(data_inicio)
pt.scroll(-60)
circulo_data = pt.locateCenterOnScreen('D:\Trabalho\Veloci\_Bots\Aloisio\Feedback de Clientes\imagens\circulo_data.png', confidence=0.7) # ENCONTRAR O ICONE NA PAGINA
pt.moveTo(circulo_data, duration=0.5) # MOVE PARA A DATA DENTRO DO CALENDARIO
pt.click() # CLICA

time.sleep(1.5)

pt.tripleClick(x=1408, y=700) # MOVER PARA "ATE" & SELECIONAR O CAMPO COMPLETO
pt.press('Delete')
pt.write(data_final)
circulo_data = pt.locateCenterOnScreen('D:\Trabalho\Veloci\_Bots\Aloisio\Feedback de Clientes\imagens\circulo_data.png', confidence=0.7) # ENCONTRAR O ICONE NA PAGINA
pt.moveTo(circulo_data, duration=0.5) # MOVE PARA A DATA DENTRO DO CALENDARIO
pt.click() # CLICA

pt.click(x=1357, y=800) # POSSUI COMENTARIOS
pt.click(x=1305, y=886) # SIM

pt.click(x=1640, y=826) # CONSULTAR

# COLOCAR O EXCEL NA SEGUNDA TELA ===============================================================
pt.keyDown('Alt')
pt.press('Tab', presses=2)
pt.keyUp('Alt')
pt.click(x=71, y=248)
pt.write(cel)
pt.press('Enter')

# AJUSTAR A PAGINA PARA INICIAR =================================================================
pt.keyDown('Alt')
pt.press('Tab')
pt.keyUp('Alt')
pt.click(x=1908, y=201)
pt.scroll(-700)

# VETOR PARA COPIA DOS COMENTARIOS ==============================================================
for i in range(repeticao):

# COPIAR NA PAGINA DO IFOOD =====================================================================
    pt.moveTo(x=412, y=251)
    pt.mouseDown(button='Left')
    pt.moveTo(x=1766, y=731, duration=0.5)
    pt.scroll(-800)
    pt.rightClick()
    pt.click(x=1611, y=768)
    pt.doubleClick(x=1163, y=953)
    pt.keyDown('Alt')
    pt.press('Tab')
    pt.keyUp('Alt')
    time.sleep(1)

# COLAR NO EXCEL ================================================================================
    pt.hotkey('Ctrl', 'Shift', 'V')
    pt.keyDown('Ctrl')
    pt.press('Down')
    pt.keyUp('Ctrl')
    pt.press('Down')

# PASSAR PARA A PROXIMA PAGINA ==================================================================
    pt.hotkey('Alt', 'Tab')
    pt.click(x=1163, y=953)
    seta_vermelha = pt.locateCenterOnScreen('D:\Trabalho\Veloci\_Bots\Aloisio\Feedback de Clientes\imagens\seta_vermelha.png', confidence=0.95) # ENCONTRAR O ICONE NA PAGINA
    pt.moveTo(seta_vermelha, duration=0.5) # CLICA NA SETA VERMELHA
    pt.click()
    pt.scroll(414)

