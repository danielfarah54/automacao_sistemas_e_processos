import pandas as pd
import pyautogui
import pyperclip
import time

# Preencher seus dados
seu_nome = 'Daniel'
seu_gmail = 'daniel+diretoria@gmail.com'
link = 'https://drive.google.com/minha_pasta' # 'sistema' da empresa
arquivo_pasta_download = '/home/daniel/Downloads/Vendas-Dez.xlsx'

pyautogui.PAUSE = 3
pyautogui.alert('Vai começar, aperte OK e não mexa em nada')

# PASSO 1
# Abrir uma nova aba
pyautogui.hotkey('ctrl', 't')

# Entrar no link do sistema
pyperclip.copy(link)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
time.sleep(7)

# PASSO 2:
# Entrar na pasta 'exportar'
pyautogui.click(414, 271, clicks=2)
time.sleep(2)

# PASSO 3:
# Baixar planilha
pyautogui.click(407, 353)
pyautogui.click(1163, 153)
pyautogui.click(1052, 599)
time.sleep(10)

# PASSO 4:
# Ler planilha
df = pd.read_excel(arquivo_pasta_download)
display(df)

# Calcular faturamento e quantidade
faturamento = df['Valor Final'].sum()
qtde_produtos = df['Quantidade'].sum()

# PASSO 5:
# Abrir email
pyautogui.hotkey('ctrl', 't')
pyautogui.write('mail.google.com')
pyautogui.press('enter')
time.sleep(5)

# Clicar em escrever email
pyautogui.click(133, 183)

# Preencher as informações
pyautogui.write(seu_gmail)
pyautogui.press('tab')
pyautogui.press('tab')
assunto = 'Relatório de Vendas de Ontem'
pyperclip.copy(assunto)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')
texto = f'''
Prezados, bom dia

O faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade de produtos foi de: {qtde_produtos:,}

Att,
{seu_nome}'''
pyperclip.copy(texto)
pyautogui.hotkey("ctrl", 'v')

# Enviar e-mail
pyautogui.hotkey('ctrl', 'enter')

# Avisar que acabou
pyautogui.alert('Fim da Automação. Seu computador já voltou a ser seu')