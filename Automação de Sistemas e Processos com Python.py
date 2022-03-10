#!/usr/bin/env python
# coding: utf-8

# # Automação de Sistemas e Processos com Python
# 
# ### Desafio:
# 
# Todos os dias, o nosso sistema atualiza as vendas do dia anterior.
# O seu trabalho diário, como analista, é enviar um e-mail para a empresa, assim que começar a trabalhar, com o faturamento e a quantidade de produtos vendidos no dia anterior
# 
# E-mail para quem você vai enviar : <br>
# Local onde o sistema disponibiliza as vendas do dia anterior: https://drive.google.com/drive/folders/1V9HCallexZCXAxkfSUAIXBLkN0AKx7bm?usp=sharing 
# 
# Para resolver isso, vamos usar o pyautogui, uma biblioteca de automação de comandos do mouse e do teclado
# 
# Comandos pyautogui: https://pyautogui.readthedocs.io/en/latest/quickstart.html
# 
# github: https://github.com/euhenri

# In[ ]:


# Como instalar
get_ipython().system('pip install pyautogui')
get_ipython().system('pip install pyperclip')


# In[70]:


import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 1

# Passo 1: Entrar no sistema (no nosso caso, entrar no link)
pyautogui.press('win')
pyperclip.copy('chrome')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
pyperclip.copy('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')


time.sleep(5)

# Passo 2: Navegar até o local do relatório (entrar na pasta Exportar)
pyautogui.click(x=359, y=250, clicks=2)
time.sleep(2)

# Passo 3: Fazer o download do relatório
pyautogui.click(x=370, y=259)
pyautogui.click(x=1157, y=162)
pyautogui.click(x=951, y=583)
time.sleep(4)

pyautogui.click(x=503, y=444,clicks=2)
time.sleep(6)


# ### Vamos agora ler o arquivo baixado para pegar os indicadores
# 
# - Faturamento
# - Quantidade de Produtos

# In[ ]:


# Passo 4: Calcular os indicadores
import pandas as pd

tabela = pd.read_excel(r'C:\Users\Henri\Downloads\Vendas - Dez.xlsx')
display(tabela)
faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()


# ### Vamos agora enviar um e-mail pelo gmail

# In[71]:


# Passo 5: Entrar no email
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(5)

# Passo 6: Enviar por e-mail o resultado
pyautogui.click(x=75, y=171)
time.sleep(5)

pyautogui.write("email para pessoa ou empresa que vai enviar")
pyautogui.press("tab") # seleciona o email
# escreve outro email
# tab
# escreve outro emailhttps://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing

# tab
pyautogui.press("tab") # pula pro campo de assunto
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v") # escrever o assunto
pyautogui.press("tab") #pular pro corpo do email

texto = f"""
Prezados, bom dia

O faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade de produtos foi de: {quantidade:,}

Abs
Luan"""

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")

# clicar no botão enviar

# apertar Ctrl Enter
pyautogui.hotkey("ctrl", "enter")


# #### Use esse código para descobrir qual a posição de um item que queira clicar
# 
# - Lembre-se: a posição na sua tela é diferente da posição na minha tela

# In[ ]:


time.sleep(5)
pyautogui.position()


# In[ ]:




