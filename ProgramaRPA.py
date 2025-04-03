import os
import pyautogui
import time
from datetime import datetime
from openpyxl import Workbook

# Definir a pausa entre os comandos
pyautogui.PAUSE = 1

# Função para clicar no botão (adaptado para mouse)
def clicar_botao(imagem):
    try:
        # Localizar o botão na tela usando a imagem
        botao = pyautogui.locateOnScreen(imagem, confidence=0.9)
        if botao:
            centro = pyautogui.center(botao)
            pyautogui.click(centro)  # Clicar no centro do botão
            print(f"Botão {imagem} encontrado e clicado!")
            return True
        else:
            print(f"Botão {imagem} não encontrado!")
            return False
    except Exception as e:
        print(f"Erro ao tentar clicar em {imagem}: {e}")
        return False

# Função para abrir o menu iniciar e pesquisar Word
def abrir_menu_iniciar():
    pyautogui.press("win")
    pyautogui.write("word", interval=0.1)
    pyautogui.press("enter")
    time.sleep(2)

# Função para selecionar o documento em branco
def selecionar_documento_em_branco():
    documento_em_branco = False
    tentativas = 0
    while not documento_em_branco and tentativas < 5:
        # Clica no botão "Documento em branco" (substitua pelo nome correto da imagem)
        documento_em_branco = clicar_botao("documento_em_branco.jpg")
        
        if not documento_em_branco:
            print("Tentando novamente...")
            tentativas += 1
            time.sleep(2)

    return documento_em_branco

# Função para digitar texto no documento
def digitar_texto():
    texto = "Primeiro projeto de RPA realizado com sucesso pelas alunas Beatriz Bramont e Isadora Cristyne!"
    pyautogui.write(texto, interval=0.1)
    print("Texto digitado com sucesso!")

# Função para capturar a tela
def capturar_tela():
    nome_arquivo = datetime.now().strftime("minha_tela_%Y%m%d_%H%M%S.jpg")
    screenshot = pyautogui.screenshot()
    screenshot.save(nome_arquivo)
    print("Imagem salva com sucesso!")
    return nome_arquivo

# Função para fechar a janela do Word
def fechar_janela_word():
    pyautogui.hotkey("alt", "f4")
    time.sleep(1)
    os.system("taskkill /f /im winword.exe")
    print("Janela do Word fechada com sucesso!")

# Função para gerar relatório em Excel
def gerar_relatorio(tarefas_executadas):
    wb = Workbook()
    ws = wb.active
    ws.append(["Tarefa", "Status", "Tempo Estimado"])
    
    for tarefa in tarefas_executadas:
        ws.append([tarefa['tarefa'], tarefa['status'], tarefa['tempo']])
    
    nome_relatorio = f"relatorio_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    wb.save(nome_relatorio)
    print(f"Relatório gerado com sucesso! Arquivo: {nome_relatorio}")

# Função principal para executar a automação
def executar_automacao():
    tarefas_executadas = []
    
    # Abrir o menu iniciar e pesquisar Word
    tarefas_executadas.append({"tarefa": "Abrir menu iniciar", "status": "Em progresso", "tempo": "1"})
    abrir_menu_iniciar()
    tarefas_executadas[-1]["status"] = "Sucesso"
    
    # Selecionar documento em branco
    tarefas_executadas.append({"tarefa": "Selecionar documento em branco", "status": "Em progresso", "tempo": "3"})
    documento_em_branco = selecionar_documento_em_branco()
    if documento_em_branco:
        tarefas_executadas[-1]["status"] = "Sucesso"
    else:
        tarefas_executadas[-1]["status"] = "Falha"

    if documento_em_branco:
        # Digitar texto no documento
        tarefas_executadas.append({"tarefa": "Digitar texto", "status": "Em progresso", "tempo": "2"})
        digitar_texto()
        tarefas_executadas[-1]["status"] = "Sucesso"
        
        # Capturar a tela
        tarefas_executadas.append({"tarefa": "Capturar tela", "status": "Em progresso", "tempo": "2"})
        nome_arquivo = capturar_tela()
        tarefas_executadas[-1]["status"] = "Sucesso"
        
        # Fechar janela do Word
        tarefas_executadas.append({"tarefa": "Fechar janela do Word", "status": "Em progresso", "tempo": "2"})
        fechar_janela_word()
        tarefas_executadas[-1]["status"] = "Sucesso"

    # Gerar relatório com o resumo das tarefas
    gerar_relatorio(tarefas_executadas)

# Chamada principal
if __name__ == "__main__":
    executar_automacao()
