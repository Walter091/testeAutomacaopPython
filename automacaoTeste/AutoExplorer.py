import pyautogui as py
from pyautogui import press, leftClick, rightClick, doubleClick, moveTo
from nucleo import InfoMsg, Atalhos_Teclado
import AutoBase
from time import sleep
from pyperclip import paste

NumFile = None
NameFIle = None
Status = "Original"

def Index():
    #mensagem alert
    InfoMsg.msgInicializacao("Automação iniciada. Por favor, não use o teclado "
                    "e nem mecha no mouse, para evitar erros!")

    #Abrindo a Aréa de Trabalho
    AutoBase.toAreaDeTrabalho()
    #Abrindo o Explorador de Arquivos
    AutoBase.openExplorerFiles()
    #Maximizando a tela
    sleep(1)
    moveTo(x=858, y=152)
    leftClick()

    #Acessando e informando o caminho do arquivo zip
    openPastaOriginal()

    entra = True
    indice = 1
    while entra:
        if indice <= 11:
            if abrirArquivo(indice):
                sleep(2)
                salvarArquivo(indice)
                sleep(2)
                closeFile()
                Atalhos_Teclado.altTab()
                createPlanilha(indice)
                sleep(2)
                indice = indice + 1
                entra = True
            else:
                InfoMsg.msgInicializacao("Impossivel Abrir arquivos .docx"
                                         " ou arquivos que não começam com um Númeral!!"
                                         "Pressione 'Enter' para Continuar.")
                press("esc")
                continue
        else:
            entra = False

    InfoMsg("Automação encerrada. Obrigado e volte Sempre!")

def openPastaOriginal():
    searchFiles("Downloads", "RPA-Artigo")
    sleep(2)
    moveTo(x=255, y=158)
    doubleClick()

def ordenarPorPdf():
    moveTo(x=569, y=103)
    leftClick()
    press("down")
    press("enter")

def ordenarPorNome():
    moveTo(x=262, y=99)
    leftClick()

def abrirArquivo(indice):
    sleep(2)
    conteudo = None
    if indice == 1:
        ordenarPorPdf()
        ordenarPorNome()
        press("down", 1)
        press("up")
        moveTo(x=524, y=133)
        rightClick()
        press("down", 5)
        press("enter")
        py.hotkey('ctrl', 'c') #Copia conteúdo selecionado
        sleep(2)
        conteudo = paste()
        py.typewrite(conteudo)  # Digita o texto da variável conteúdo
        file = str(conteudo)
        if validationArquivo(file):
            press("enter")
            sleep(1)
            press("enter")
            return True
        else:
            return False
    else:
        press("down", indice-1)
        press("apps")
        press("down", 5)
        press("enter")
        py.hotkey('ctrl', 'c') #Copia conteúdo selecionado
        conteudo = paste()
        py.typewrite(conteudo)  # Digita o texto da variável conteúdo
        sleep(2)
        file = str(conteudo)
        if validationArquivo(file):
            press("enter")
            sleep(1)
            press("enter")
            return True
        else:
            return False


def salvarArquivo(indice):
    sleep(5)
    moveTo(x=31, y=50)
    leftClick()
    sleep(3)
    moveTo(x=83, y=239)
    leftClick()
    sleep(4)
    py.write("Pagina " + NumFile + " - Modificado")
    sleep(3)
    press("enter")
    global Status
    Status = "documento alterado"

def validationArquivo(file):
    #Verifica se è numerica e .pdf
    if file[0].isnumeric():
        if file.__contains__(".pdf"):
            setInfoNameFile(file)
            return True
    else:
        return False

def setInfoNameFile(file):
    global NumFile
    global NameFIle
    if file[0:2].isnumeric():
        NumFile = file[0:2]
    else:
        NumFile = file[0]
    # Adiciona a variavel para posterior utilização
    NameFIle = file


# ----------------------------------------------------------------------------

def createPlanilha(indice):
    if indice == 1:
        AutoBase.openWPSOfiice()
        #Entrando na aréa de planilha
        moveTo(x=232, y=138)
        leftClick()
        sleep(3)
        #Criando uma planilha nova
        moveTo(x=391, y=314)
        leftClick()
        sleep(3)
        criarCabecalho()
        sleep(1)
    else:
        sleep(2)
        searchFiles("Documentos", "Relatorio de execucao")
        sleep(3)
        py.moveTo(x=320, y=116)
        py.doubleClick()
        sleep(3)

        # py.moveTo(x=255, y=319)
        # py.doubleClick()
        # time.sleep(7)

    addInfoPlanilha(indice)
    salvarPlanilha(indice)
    Atalhos_Teclado.altTab()
    if indice != 1:
        openPastaOriginal()

def criarCabecalho():
    #Seleciona o campo do cabecalho
    sleep(5)
    mesclarCelulas()
    #Voltando pro começo da planilha
    moveTo(x=93, y=205)
    leftClick()
    #Selecionando
    with py.hold('shift'):
        press('right')
        press("down", 10)

    #Clica em Inserir
    moveTo(x=323, y=46)
    leftClick()
    sleep(1)
    #Clica na setinha de lado
    moveTo(x=1016, y=100)
    leftClick()
    sleep(2)
    #Clica em cabecalho e rodapé
    moveTo(x=359, y=94)
    leftClick()
    sleep(2)
    #Clica em cabeçalho personalizado
    moveTo(x=662, y=277)
    leftClick()
    sleep(2)
    #Adiciona as colunas
    py.write("Nome do Documento")
    press("tab")
    py.write("Status")
    sleep(2)
    moveTo(x=695, y=525)
    leftClick()
    sleep(2)
    #confirma a criação
    moveTo(x=622, y=550)
    leftClick()
    sleep(2)

def mesclarCelulas():
    cont = 0
    while cont<10:
        with py.hold('shift'):
            press(['right', 'right', 'right', 'right'])
            sleep(1)
            #Indo até a opção mesclar células
            moveTo(x=619, y=112)
            leftClick()
            moveTo(x=655, y=217)
            leftClick()

        press("down", 1)
        cont = cont + 1

    sleep(1)
    moveTo(x=93, y=205)
    leftClick()
    press('right')
    cont = 0
    while cont<10:
        with py.hold('shift'):
            press(['right', 'right'])
            sleep(1)
            #Indo até a opção mesclar células
            moveTo(x=619, y=112)
            leftClick()
            moveTo(x=655, y=217)
            leftClick()

        press("down", 1)
        cont = cont + 1

def addInfoPlanilha(indice):
    if indice == 1:
        moveTo(x=93, y=205)
        leftClick()
    sleep(1)
    press("space")
    press("backspace")
    sleep(1)
    py.write(NameFIle)
    # py.press("enter")
    press("tab")
    py.write(Status)
    press("enter")
    sleep(2)

def salvarPlanilha(indice):
    sleep(1)
    moveTo(x=37, y=48)
    leftClick()
    sleep(1)
    if indice == 1:
        moveTo(x=116, y=228)
        leftClick()
        moveTo(x=396, y=322)
        doubleClick()
        press("tab")
        sleep(1)
        py.write("Relatorio de execucao")
        sleep(1)
        press("enter")
    else:
        sleep(1)
        moveTo(x=95, y=189)
        leftClick()

    closeFile()


# ----------------------------------------------------------------------------

def searchFiles(nameDirectory, nameFile):
    #Acessando o diretorio
    moveTo(x=615, y=68)
    leftClick()
    sleep(2)
    py.write(nameDirectory)
    press("enter")
    #Pesquisando e acessando o arquivo
    sleep(2)
    moveTo(x=724, y=64)
    leftClick()
    sleep(2)
    py.write(nameFile)
    press("enter")

def closeFile():
    moveTo(x=345, y=11)
    sleep(1)
    leftClick()
    sleep(2)
