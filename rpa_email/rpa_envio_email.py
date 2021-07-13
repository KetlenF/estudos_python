#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pandas as pd
import pyperclip
import pyautogui
import time

df = pd.read_excel("\Exportar\Vendas - Dez.xlsx")
display(df)
faturamento = df['Valor Final'].sum()
qtde_produtos = df['Quantidade'].sum()


# In[ ]:




pyautogui.alert("Vai começar!")

pyautogui.PAUSE = 2
# opção 1 - abrir navegador novo e entrar no chrome
pyautogui.press("winleft")
pyautogui.write("opera")
pyautogui.press("enter")

# abrir aba gmail
#pyautogui.hotkey('ctrl', 't')
pyautogui.write("mail.google.com")
pyautogui.press('enter')
time.sleep(5)

# clicar em escrever email
pyautogui.click(107, 207)
time.sleep(5)

# preencher informações do e-mail
pyautogui.write('ketlen.fernandescs+teste@gmail.com')
pyautogui.press('tab')
pyautogui.press('tab')
assunto = "Relatório de Vendas de Ontem"
pyperclip.copy(assunto)
pyautogui.hotkey("ctrl", 'v')
pyautogui.press("tab")
texto = f"""
Prezados, bom dia

O faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade de produtos foi de: {qtde_produtos:,}
"""
pyperclip.copy(texto)
pyautogui.hotkey("ctrl", 'v')

# enviar e-mail
pyautogui.hotkey('ctrl', 'enter')

# avisar que acabou
pyautogui.alert("Fim da Automação. Seu computador já voltou a ser seu")


# In[ ]:





# In[ ]:




