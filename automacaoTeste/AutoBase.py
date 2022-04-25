import pyautogui as py
import openpyxl as xl
import shutil
import time

def localizationPositionMouse():
    py.alert("Você entrou no modo de Automação. "
                    "Te daremos 5 seconds para colocar o mouse na posição desejada!!")
    py.sleep(5)
    print(py.position())
# -------------------------------------------------- Google
def openGoogle():
    py.alert("Você entrou no modo de Automação. "
                    "Por favor não utilize o mouse nem o teclado!!")
    py.press("winleft")
    py.write("chrome")
    py.press("enter")
    py.sleep(4)

def openWhatsapChrome():
    openGoogle()
    py.write("https://web.whatsapp.com/")
    py.press("enter")

def openYoutubeChrome():
    openGoogle()
    py.moveTo(x=407, y=482)
    py.mouseDown()
    time.sleep(2)
    py.leftClick()

def openTrello():
    openGoogle()
    time.sleep(3)
    py.moveTo(x=58, y=77)
    py.leftClick()
    py.moveTo(x=82, y=112)
    py.leftClick()
# ---------------------------------------------------------- S.O
def toAreaDeTrabalho():
    py.hotkey("winleft", "d")

def openWPSOfiice():
    toAreaDeTrabalho()
    py.press("winleft")
    time.sleep(2)
    py.write("WPS OFFICE")
    py.press("enter")
    time.sleep(10)

def openExplorerFiles():
    # Abrindo o Explorador de Arquivos
    py.press("winleft")
    time.sleep(1.5)
    py.write("Explorador de Arquivos")
    time.sleep(5)
    py.press("enter")

# ---------------------------------------------------------- Excel

def criarPlanilha(nomePlanilha):
    planilha = xl.Workbook()
    salvarPlanilha(planilha, nomePlanilha)
    print("Planilha Criada!")

def carregarPlanilha(caminho):
    planilha = xl.load_workbook(caminho)
    print('planilha Carregada')
    return planilha

def salvarPlanilha(planilha, nome):
    planilha.save(nome)
    print("Planilha Salva!")

def mesclarCelulas(planilha, celulas, linhaInicial, linhaFinal, colunaInicial, colunaFinal):
    clls = planilha.active
    clls.merge_cells(celulas)
    clls.merge_cells(start_row=linhaInicial, start_column=colunaInicial, end_row=linhaFinal, end_column=colunaFinal)
    print('Células Mescladas')

def moverPlanilha(caminho, novoCaminho):
    shutil.move(caminho, novoCaminho)
    print("Planilha movida")
