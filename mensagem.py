import pyautogui
import time
import pandas as pd

# ler o arquivo Excel com os números dos funcionários
df = pd.read_excel('C:/Users/campo/Documents/contatos.xlsx')
cont = 0

mensagem = str(input('Digite a mensagem que será enviada a todos os contatos: '))


pyautogui.press('winleft')
time.sleep(0.5)
pyautogui.typewrite('chrome')
pyautogui.press('enter')
time.sleep(1)
cont = 0
inicio = time.time()

for index, row in df.iterrows():
    pyautogui.hotkey('alt', 'home')
    time.sleep(0.5)
    numero = str(row['A'])
    time.sleep(1)
    pyautogui.typewrite('wa.me/')
    pyautogui.typewrite(numero[1:].split('-'))
    pyautogui.press('enter')
    time.sleep(1)
    # escrever e enviar a mensagem
    pyautogui.typewrite(mensagem)
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.hotkey('alt', 'tab')
    cont += 1

fim = time.time()
tempo = (fim - inicio)/60
pyautogui.alert('{} mensagens enviadas em {:.2} minutos!'.format(cont, tempo))




