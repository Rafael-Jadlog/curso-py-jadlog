#!/usr/bin/env python
# coding: utf-8

# In[ ]:





# IMPORTANDO BIBLIOTECAS

# In[1]:


import os
from datetime import datetime
from email.message import EmailMessage
import smtplib
import pywhatkit as kit
from time import sleep
import shutil 
import pyautogui
import win32com.client as win32


# ORGANIZANDO ARQUIVOS

# In[2]:


arquivo_excel = 'contas_processadas.xlsx'

pasta_base = 'base'

# Criando a pasta base dentro da pasta aula 04  


# VERIFICANDO SE O ARQUIVO EXISTE

# In[3]:


if not os.path.exists(arquivo_excel):
    raise FileNotFoundError('Arquivo não encontrado, verifique se o dowload esta concluído')

print('Arquivo encontrado')



# CRIANDO A PASTA BASE DENTRO DA PASTA AULA 04

# In[4]:


if not os.path.exists(pasta_base):
    os.makedirs(pasta_base)
    print('Pasta criada com sucesso')

else:
    print('Pasta base ja existe')


# MOVENDO ARQUIVO PARA DENTRO DA PASTA BASE

# In[5]:


destino = os.path.join(pasta_base,os.path.basename(arquivo_excel))

if not os.path.exists(destino):
    shutil.move(arquivo_excel,destino)
    print('Arquivo movido com sucesso!')

else:
    print('Arquivo ja existe na pasta base')    


# CRIANDO CORPO DO EMAIL

# In[6]:


hoje = datetime.now().strftime('%d/%m/%Y | %H:%M')

corpo_email = f'''
Bom dia!

Segue em anexo relatório fechamento de contas a pagar e receber
dados atualizado: 📅{hoje}🕛


'''

print(corpo_email)


# CONFIGURANDO OUTLOOK

# In[31]:


#pythoncom.coinitialize()  forcar inicialização

outlook = win32.Dispatch('Outlook.Application')

email = outlook.CreateItem(0)

email.to = 'leon4rdoalvess@gmail.com;leafarcardoso1@gmail.com'
email.Subject = 'Relatório de contas processadas de Janeiro'
email.Body = corpo_email

email.attachments.Add(
    r'C:\Users\rafael.cardoso\Desktop\Curso Python\Aula-04\base\contas_processadas.xlsx'
)

email.Send()

print('Email enviado com sucesso!')







# ALERTA WHATSAPP

# In[7]:


telefone = '+55119979714423'
telefone2 = '+5511946151206'

mensagem = 'Relatorio enviado no e-mail'

kit.sendwhatmsg_instantly(
    telefone2
    
    ,
    mensagem,
    wait_time = 20,
    tab_close = False
)

sleep(2)

pyautogui.press('enter')
print('mensagem enviado com sucesso')



# CRIANDO EXECUTAVEL

# In[9]:


#!pip install pyinstaller

get_ipython().system('jupyter nbconvert --to script projeto.ipynb')

#!pyinstaller --onefile projeto.py


