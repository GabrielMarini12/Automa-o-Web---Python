from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

# 1º: abrir o navegador
navegador = webdriver.Chrome()

# Passo 1: Pegar a cotação do Dólar
navegador.get("https://www.google.com/") #abrindo o navegador
navegador.find_element(By.XPATH,'/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("cotação dólar")
navegador.find_element(By.XPATH,'/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
cotacao_dolar = navegador.find_element(By.XPATH, '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute("data-value")
print(cotacao_dolar)

# Passo 2: Pegar a cotação do Euro
navegador.get("https://www.google.com/")
navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("cotação euro")
navegador.find_element(By.XPATH,'/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
cotacao_euro = navegador.find_element(By.XPATH, '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute("data-value")
print(cotacao_euro)

# Passo 3: Pegar a cotação do Ouro
navegador.get("https://www.melhorcambio.com/ouro-hoje")
cotacao_ouro = navegador.find_element(By.XPATH, '//*[@id="comercial"]').get_attribute("value")
cotacao_ouro = cotacao_ouro.replace(",", ".")
print(cotacao_ouro)

navegador.quit() #fechando o navegador

import pandas as pd

tb = pd.read_excel("C:\\Users\\Usuario\\Desktop\\PROJECTS\\INTENSIVAO_PY\\Produtos.xlsx")
print(tb)

# Passo 5: Recalcular o preço de cada produto
# atualizar a cotação
# nas linhas onde na coluna "Moeda" = Dólar vai alterar o valor da cotação
tb.loc[tb["Moeda"] == "Dólar", "Cotação"] = float(cotacao_dolar)
tb.loc[tb["Moeda"] == "Euro", "Cotação"] = float(cotacao_euro)
tb.loc[tb["Moeda"] == "Ouro", "Cotação"] = float(cotacao_ouro)

# atualizar o preço base reais (preço base original * cotação)
tb["Preço de Compra"] = tb["Preço Original"] * tb["Cotação"]

# atualizar o preço final (preço base reais * Margem)
tb["Preço de Venda"] = tb["Preço de Compra"] * tb["Margem"]

# formatando o valor de venda para ficar bonitinho na tabela
tb["Preço de Venda"] = tb["Preço de Venda"].map("R${:.2f}".format)
print(tb)

# Passo 6: Salvar os novos preços dos produtos
tb.to_excel("C:\\Users\\Usuario\\Desktop\\PROJECTS\\INTENSIVAO_PY\\Produtos Ajustados.xlsx", index=False)