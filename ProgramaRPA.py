import os
import pyautogui
import time
from datetime import datetime

pyautogui.PAUSE = 1

def clicar_botao(imagem):
    botao = pyautogui.locateOnScreen(imagem, confidence=0.9)  # Aumenta a precisão com 'confidence'
    time.sleep(1)
    centro = pyautogui.center(botao)
    time.sleep(0.5) 
    print(centro)
    if centro:
        pyautogui.click(centro)
        print(f"Botão {imagem} encontrado e clicado!")
        return True
    else:
        print(f"Botão {imagem} não encontrado!")
        return False

# Abrir o menu iniciar do Windows
pyautogui.press("win")

# Pesquisar Word
pyautogui.write("word", interval=0.1)
pyautogui.press("enter")
time.sleep(2)

documento_em_branco = False
tentativas = 0

while not documento_em_branco and tentativas < 5:
    # Clica no botão "Documento em branco" (substitua pelo nome correto da imagem)
    documento_em_branco = clicar_botao("documento_em_branco.jpg")
    
    if not documento_em_branco:
        print("Tentando novamente...")
        tentativas += 1
        time.sleep(2)

if documento_em_branco:
    print("Documento em branco selecionado com sucesso!")

    # Pressionar Enter para garantir que o documento está em foco
    pyautogui.press("enter")
    time.sleep(1)

    # Digitar texto no documento
    texto = "Primeiro projeto de RPA realizado com sucesso pelas alunas Beatriz Bramont e Isadora Cristyne!"
    pyautogui.write(texto, interval=0.1)
    print("Texto digitado com sucesso!")

# Capturar tela
    nome_arquivo = datetime.now().strftime("minha_tela_%Y%m%d_%H%M%S.jpg")
    screenshot = pyautogui.screenshot()
    screenshot.save(nome_arquivo)
    print("Imagem salva com sucesso!")  

    #Fechar janela do Word
    pyautogui.hotkey("alt", "f4")
    time.sleep(1)
    os.system("taskkill /f /im winword.exe")
    print(f"Automação concluída com sucesso! Arquivo salvo como: {nome_arquivo}")
else:
    print("Não foi possível selecionar o documento em branco.")

