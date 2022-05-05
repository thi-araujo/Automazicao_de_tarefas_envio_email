import pyautogui
import pyperclip
import time

#pyautogui -> automatização do seu mouse, teclado e tela
# !pip install pyautogui
# !pip install pyperclip

pyautogui.PAUSE = 1

# Passo 1: Entrar no sistema (no nosso caso, entrar no link)
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

# carregar o sistema (demorar 2 segundos)
time.sleep(2)

# Passo 2: Navegar até o local do relatório (entrar na pasta Exportar)
pyautogui.click(x=333, y=288, clicks=2)
time.sleep(2)

# Passo 3: Fazer o download do relatório
pyautogui.click(x=353, y=281) # cliclar no arquivo
pyautogui.click(x=1470, y=183) # clicar nos 3 pontos
pyautogui.click(x=1323, y=590)  #clicar em fazer download
time.sleep(5)

# Importar base de dados no Python
# pandas
# numpay
# openpyxl
import pandas as pd
tabela = pd.read_excel(r"C:\Users\Thiago\Downloads\Vendas - Dez.xlsx") # sempre que for passar um caminho do computador coloca um r antes do caminho
display(tabela)

# Calcular faturamento e quantidade de produtos vendidos (os indicadores)
Faturamento = tabela["Valor Final"].sum()
Qtde_Produtos = tabela["Quantidade"].sum()
display(Faturamento)
display(Qtde_Produtos)

# Abrir e-mail link e-mail: https://outlook.live.com/mail/0/inbox/id/AQQkADAwATZiZmYAZC1iZTdjLTBiZWYtMDACLTAwCgAQAGhaXoAshe5MkI0hdLBZHEo%3D
import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 1


pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://outlook.live.com/mail/0/inbox/id/AQQkADAwATZiZmYAZC1iZTdjLTBiZWYtMDACLTAwCgAQAGhaXoAshe5MkI0hdLBZHEo%3D")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

# carregar o sistema (demorar 5 segundos)
time.sleep(5)
# clicar no botão escrever
pyautogui.click(x=100, y=154)

# preencher o campo destino
pyautogui.write("thiago.araujo_@hotmail.com")
pyautogui.press("tab") # seleciona o e-mail
pyautogui.press("tab") # muda para o assunto

# escerver campo assunto
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab") # muda para o o campo corpo do e-mail
# escrever o e-mail
texto = f"""
Prezados, bom dia

O Faturamento de ontem foi de: R${Faturamento:,.2f}
A Quantidade de produtos foi de: {Qtde_Produtos:,}

Atenciosamente
Thiago Araujo
"""
pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")
# clicar em enviar
pyautogui.hotkey("ctrl", "enter")