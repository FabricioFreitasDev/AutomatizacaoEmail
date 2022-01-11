import pyautogui
import pyperclip
import time
import pandas as pd
import numpy
import openpyxl

#tempo/delay
pyautogui.PAUSE = 1

#comandos a serem executados de forma automática
pyautogui.press("win")
pyautogui.write("chrome")
pyautogui.press("enter")
#pyautogui.hotkey("ctrl", "t") #ABRE UMA NOVA ABA, QUANDO O NAVEGADOR ESTIVER ABERTO.
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing") #Copia o link
pyautogui.hotkey("ctrl", "v") #cola o link
pyautogui.press("enter")

#carregar o sistema(esperar 5 segundos)
time.sleep(5)

#Navegar no sistema até a pasta Exportar
pyautogui.click(x=347, y=264, clicks=2)
time.sleep(2)

#Fazer o Download da Base De Vendas
pyautogui.click(x=367, y=351) #clicar no arquivo
pyautogui.click(x=1156, y=154)#Clicar nos três Prontos
pyautogui.click(x=912, y=564)#Fazer o download

time.sleep(5) #Esperar a conclusão do downlod

tabela = pd.read_excel(r"C:\Users\fabri\Desktop\Vendas - Dez.xlsx") #Usar o (r) antes de cada diretorio
print(tabela)

#calcular o faturamento e a quantidade de produtos vendidos(os indicadores)
faturamento = tabela ["Valor Final"].sum() #Soma
quantidade_produtos = tabela["Quantidade"].sum()

#Enviar email para diretoria
link = "https://outlook.live.com/mail/0/"

#Abrir email
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://outlook.live.com/mail/0/")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

time.sleep(5)
#Clicar no botão escrever
pyautogui.click(x=144, y=130)
time.sleep(2)
#Preencher o destino
pyautogui.write("famaiden@hotmail.com")
pyautogui.press("tab") # Seleciona o email

#time.sleep(1)
#pyautogui.press("tab") #Muda para o campo do assunto

#Preencher o assunto
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v")
time.sleep(1)
pyautogui.press("tab")

#Escrever o email
texto = f"""
Prezados, Bom Dia!!!

O Faturamento de ontem foi de R$: {faturamento:,.2f}
A quantidade de produtos vendidas foi de: {quantidade_produtos:,}

Qualquer Dúvida Estou à Disposição.

Att: Fabricio Freitas
"""
pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")

#Clicar em enviar
pyautogui.hotkey("ctrl", "enter")