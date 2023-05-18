#!/usr/bin/env python
# coding: utf-8

# In[1]:


#Criando navegador
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

options = webdriver.ChromeOptions()
options.add_experimental_option("prefs", {
  "download.default_directory": r"C:\Users\vggam\anaconda3\AprendendoPython\Automação Web\Exercício - Nota Fiscal",
  "download.prompt_for_download": False,
  "download.directory_upgrade": True,
  "safebrowsing.enabled": True
})

navegador = webdriver.Chrome(chrome_options=options)


# In[2]:


#Criando caminho para o arquivo no navegador
import os

caminho = os.getcwd()
caminho = caminho + r'\login.html'
print(caminho)
navegador.get(caminho)


# ### Fazendo o login

# In[3]:


navegador.find_element(By.XPATH, '/html/body/div/form/input[1]').send_keys('vito@gmail.com')
navegador.find_element(By.XPATH, '/html/body/div/form/input[2]').send_keys('SenhaLegal')
navegador.find_element(By.XPATH, '/html/body/div/form/button').click()


# ### Importando tabela com clientes

# In[4]:


import pandas as pd

tabela_clientes = pd.read_excel('NotasEmitir.xlsx')
display(tabela_clientes.head())


# ### Preenchendo informações

# In[5]:


for linha in tabela_clientes.index:
    
    #Limpando linhas
    navegador.find_element(By.NAME, 'nome').clear()
    navegador.find_element(By.NAME, 'endereco').clear()
    navegador.find_element(By.NAME, 'bairro').clear()
    navegador.find_element(By.NAME, 'municipio').clear()
    navegador.find_element(By.NAME, 'cep').clear()
    navegador.find_element(By.NAME, 'cnpj').clear()
    navegador.find_element(By.NAME, 'inscricao').clear()
    navegador.find_element(By.NAME, 'descricao').clear()
    navegador.find_element(By.NAME, 'quantidade').clear()
    navegador.find_element(By.NAME, 'valor_unitario').clear()
    navegador.find_element(By.NAME, 'total').clear()

    #Colocando as informações
    navegador.find_element(By.NAME, 'nome').send_keys(tabela_clientes.loc[linha, "Cliente"])
    navegador.find_element(By.NAME, 'endereco').send_keys(tabela_clientes.loc[linha, "Endereço"])
    navegador.find_element(By.NAME, 'bairro').send_keys(tabela_clientes.loc[linha, "Bairro"])
    navegador.find_element(By.NAME, 'municipio').send_keys(tabela_clientes.loc[linha, "Municipio"])
    navegador.find_element(By.NAME, 'cep').send_keys(str(tabela_clientes.loc[linha, "CEP"]))
    navegador.find_element(By.NAME, 'uf').send_keys(tabela_clientes.loc[linha, "UF"])
    navegador.find_element(By.NAME, 'cnpj').send_keys(str(tabela_clientes.loc[linha, "CPF/CNPJ"]))
    navegador.find_element(By.NAME, 'inscricao').send_keys(str(tabela_clientes.loc[linha, "Inscricao Estadual"]))
    navegador.find_element(By.NAME, 'descricao').send_keys(tabela_clientes.loc[linha, "Descrição"])
    navegador.find_element(By.NAME, 'quantidade').send_keys(str(tabela_clientes.loc[linha, "Quantidade"]))
    navegador.find_element(By.NAME, 'valor_unitario').send_keys(str(tabela_clientes.loc[linha, "Valor Unitario"]))
    navegador.find_element(By.NAME, 'total').send_keys(str(tabela_clientes.loc[linha, "Valor Total"]))
    navegador.find_element(By.XPATH, '//*[@id="formulario"]/button').click()

print('Acabou!')

