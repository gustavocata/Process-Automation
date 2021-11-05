#!/usr/bin/env python
# coding: utf-8

# # Automação de Sistemas e Processos com Python
# 
# ### Desafio:
# 
# Todos os dias, o nosso sistema atualiza as vendas do dia anterior.
# O seu trabalho diário, como analista, é enviar um e-mail para a diretoria, assim que começar a trabalhar, com o faturamento e a quantidade de produtos vendidos no dia anterior
# 
# E-mail da diretoria: seugmail+diretoria@gmail.com<br>
# Local onde o sistema disponibiliza as vendas do dia anterior: https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing
# 
# Para resolver isso, vamos usar o pyautogui, uma biblioteca de automação de comandos do mouse e do teclado

# In[1]:


import pyautogui
import pyperclip
import time 
import pandas as pd

# entrar nesse link
pyautogui.PAUSE = 1
pyautogui.hotkey('ctrl', 't')
link = "https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing"
pyperclip.copy(link)
pyautogui.hotkey('ctrl', 'v')
time.sleep(3)
pyautogui.press('enter')

# entrar na pasta e fazer o download
time.sleep(4)
pyautogui.click(x=396, y=383, clicks=2)
time.sleep(2)
pyautogui.click(x=363, y=442)
time.sleep(2)
pyautogui.click(x=1721, y=183)
pyautogui.click(x=1548, y=587)
pyautogui.press('enter')

#abrir aquivo excel
tabela = pd.read_excel(r"C:\Users\Gustavo\Downloads\Vendas - Dez.xlsx")
faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()

#mandar por email
autogui.hotkey('ctrl', 't')
link = "https://mail.google.com/mail/u/0/#inbox"
pyperclip.copy(link)
time.sleep(3)
pyautogui.press('enter')
time.sleep(3)
pyautogui.click(x=48, y=204)
pyautogui.write("gugacatarino@hotmail.com")
pyautogui.press('tab')
pyautogui.write("vendas")
pyautogui.press('tab')
pyautogui.write(f"""

Vendas do dia 17/12/1998
Produtos de qualidade: {qualidade}
Total faturamento: {faturamento}

""")
autogui.hotkey('ctrl', 'enter')


# In[12]:


import pyautogui
import pyperclip
import time 
import pandas as pd

# entrar nesse link
pyautogui.PAUSE = 1
pyautogui.hotkey('ctrl', 't')
link = "https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing"
pyperclip.copy(link)
pyautogui.hotkey('ctrl', 'v')
time.sleep(3)
pyautogui.press('enter')

# entrar na pasta e fazer o download
time.sleep(4)
pyautogui.click(x=396, y=383, clicks=2)
time.sleep(2)
pyautogui.click(x=363, y=442)
time.sleep(2)
pyautogui.click(x=1721, y=183)
pyautogui.click(x=1548, y=587)
time.sleep(5)
pyautogui.press('enter')

#abrir aquivo excel
tabela = pd.read_excel(r"C:\Users\Gustavo\Downloads\Vendas - Dez.xlsx")
faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()

#mandar por email
pyautogui.hotkey('ctrl', 't')
link = "https://mail.google.com/mail/u/0/#inbox"
pyperclip.copy(link)
pyautogui.hotkey('ctrl', 'v')
time.sleep(3)
pyautogui.press('enter')
time.sleep(6)
pyautogui.click(x=48, y=204)
time.sleep(7)
pyautogui.write("gugacatarino@hotmail.com")
time.sleep(2)
pyautogui.press('tab')
time.sleep(2)
pyautogui.press('tab')
pyautogui.write("vendas")
pyautogui.press('tab')
pyautogui.write(f"""

Vendas do dia 17/12/1998
Total de quantidades: {quantidade}
Total faturamento: {faturamento}

""")
time.sleep(2)
autogui.hotkey('ctrl', 'enter')


# ### Vamos agora ler o arquivo baixado para pegar os indicadores
# 
# - Faturamento
# - Quantidade de Produtos

# In[9]:


#abrir aquivo excel
import pandas as pd

tabela = pd.read_excel(r"C:\Users\Gustavo\Downloads\Vendas - Dez.xlsx")
display(tabela)


# ### Vamos agora enviar um e-mail pelo gmail

# In[ ]:





# #### Use esse código para descobrir qual a posição de um item que queira clicar
# 
# - Lembre-se: a posição na sua tela é diferente da posição na minha tela

# In[4]:


import pyautogui
import time 
time.sleep(5)
pyautogui.position()


# In[ ]:




