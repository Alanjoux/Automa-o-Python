import pyautogui as gui
import time

gui.alert('O código vai começar. Não use nada no seu computador enquanto o código estiver rodando')
gui.PAUSE = 0.8 # Quantidade de tempo que a automação vai levar entre a execução dos códigos.

# Abrir o google drive no meu computador.

gui.press('winleft') # press -> Comando usado para pressionar uma tecla.
gui.write('chrome') # write -> Comando usado para escrever um texto.
gui.press('enter') # press -> Comando usado para pressionar uma tecla.
time.sleep(0.9) # sleep -> Comando usado para informa quantos segundos sera executado a próxima linha de código.
gui.write('https://drive.google.com/drive/my-drive') # write -> Comando usado para escrever um texto.
gui.press('enter') # press -> Comando usado para pressionar uma tecla.


# Entrar na minha área de trabalho.

gui.hotkey('winleft', 'd') # hotkey ->Comando usado para fazer combinações de teclas.


# Clicar no arquivo que quero fazer o backup e arrasta-lo áte o google drive.
gui.moveTo(863, 334) # MoveTo -> Comando utilizado para mover o mouse.
gui.mouseDown() # MouseDown -> Comando utilizado para clicar em algo.
gui.moveTo(1123, 557) # MoveTo -> Comando utilizado para mover o mouse.
gui.hotkey('alt', 'tab') # hotkey ->Comando usado para fazer combinações de teclas.
time.sleep(1)

# Larga o arquivo no google drive.

gui.mouseUp() # mouseUp -> Camando utilizado para solta o click do mouse.

# espera 5 segundos.

time.sleep(5)
gui.alert('O código acabou de rodar. Pode usar seu computador novamente.')


