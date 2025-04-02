import os
import pyautogui
import time
from datetime import datetime

pyautogui.PAUSE = 1

# Fecha qualquer janela aberta da calculadora
os.system("taskkill /f /im calculator.exe")
time.sleep(2)

# Abrir o menu iniciar do Windows
pyautogui.press("win")

# Digitar Calculadora na pesquisa e aguardar abrir 
pyautogui.write("Calculadora", interval=0.1)
pyautogui.press("enter")
time.sleep(1)


# Capturar tela
nome_arquivo = datetime.now().strftime("minha_tela_%Y%m%d_%H%M%S.jpg")
screenshot = pyautogui.screenshot()
screenshot.save(nome_arquivo)
print("Imagem salva com sucesso!")  

#Fechar janela da calculadora
pyautogui.hotkey("alt", "f4")
time.sleep(1)
os.system("taskkill /f /im calculator.exe")
print(f"Automação concluída com sucesso! Arquivo salvo como: {nome_arquivo}")

# def verificarArquivo(nomeArquivo):
#     if os.path.exists(nomeArquivo):
#         pyautogui.alert("Arquivo encontrado")
#     else:
#         pyautogui.alert("Arquivo não encontrado")

# def verificarUsuario(senha):
#     if senha == "admin":
#         verificarArquivo("dados.csv")
#         local = "C:/Users/isadora.silva/OneDrive - TIMBRO TRADING/Cursos/FIT/TESTES"
#         listarArquivos(local)
#     else:
#         pyautogui.alert("Senha inválida! Finalizando processo...")

# def listarArquivos(diretorio):
#     for arquivo in os.listdir(diretorio):
#         print(f"Processando: {arquivo}")
#     print("Processo finalizado!")

# usuario = pyautogui.password("Informe a senha:")
# verificarUsuario(usuario)

# pyautogui.hotkey("alt", "f4") --> fecha janela