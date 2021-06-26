# Código para descobrir qual a posição de um item que queira clicar
# (a posição muda de tela para tela)

import pyautogui
import time

time.sleep(4)
print(pyautogui.position())
pyautogui.alert('Posição Registrada')