# Passo 1 Entrar no sistema da empresa

import time
import numpy as np
import pandas as pd
import openpyxl 
import pyautogui # pip install PyAutoGUI    Bone    2   25.0    11.0        

# Esperar 1 segundo para os códigos funcionar
pyautogui.PAUSE = 1

# aperta a tecla windows
pyautogui.press('win')

# Escreve o nome do navegador para abrir na barra do windows
pyautogui.write('firefox')

# Aperta Enter
pyautogui.press('enter')

# digitar o link do site na barra de endereço
link = "https://dlp.hashtagtreinamentos.com/python/intensivao/login"
pyautogui.write(link)

# Aperta Enter
pyautogui.press('enter')

# Esperar 5 segundos
time.sleep(5)


# Passo 2 Fazer login

# Pointclick do login
pyautogui.click(x=682, y=379)

# Digitar o campo de email
pyautogui.write('pythonimpressionador@gmail.com')

# Passar para o campo da senha
pyautogui.press('Tab')

# Digitar no campo da senha
pyautogui.write("minha senha")

# Clicar no botão "Enter"
pyautogui.click(x=676, y=539)
time.sleep(3)


# Passo 3 Importar a base de dados

# Importando tabela
tabela = pd.read_csv('produtos.csv')
print(tabela)

for linha in tabela.index:
    # Passo 4 Cadastrar produtos

    # Cadastrar produtos
    pyautogui.click(x=570, y=256)

    # código
    codigo = tabela.loc[linha, "codigo"]
    pyautogui.write(codigo)
    pyautogui.press("tab")

    # marca
    marca = tabela.loc[linha, "marca"]
    pyautogui.write(marca)
    pyautogui.press("tab")

    # tipo
    tipo = tabela.loc[linha, "tipo"]
    pyautogui.write(tipo)
    pyautogui.press("tab")

    # categoria
    categoria = str(tabela.loc[linha, "categoria"])
    pyautogui.write(categoria) # "1"
    pyautogui.press("tab")

    # preco_unitario
    preco_unitario = str(tabela.loc[linha, "preco_unitario"])
    pyautogui.write(preco_unitario)
    pyautogui.press("tab")

    # custo
    custo = str(tabela.loc[linha, "custo"])
    pyautogui.write(custo)
    pyautogui.press("tab")

    # obs
    obs = tabela.loc[linha, "obs"]
    if not pd.isna(obs):
        pyautogui.write(obs)

    pyautogui.write("")

    # enviar produto
    pyautogui.press("tab")
    pyautogui.press("enter")

    # Jogar a página web para o início / page up
    pyautogui.scroll(5000)


# Passo 5 Repetir isso até achar a base de dados