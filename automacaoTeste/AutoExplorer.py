import pyautogui
from nucleo import InfoMsg, Atalhos_Teclado
import AutoBase
import time
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
    time.sleep(1.5)
    pyautogui.moveTo(x=858, y=152)
    pyautogui.leftClick()

    #Acessando e informando o caminho do arquivo zip
    openPastaOriginal()

    entra = True
    indice = 1
    while entra:
        if(abrirArquivo(indice)):
            time.sleep(5)
            salvarArquivo(indice)
            time.sleep(2)
            closeFile()
            Atalhos_Teclado.altTab()
            createPlanilha(indice)
            Atalhos_Teclado.altTab()
            time.sleep(2)
            indice = indice + 1
            entra = True
        else:
            InfoMsg.msgInicializacao("Impossivel Abrir arquivos .docx ou arquivos que não começam com um Númeral!!")
            continue

def openPastaOriginal():
    searchFiles("Downloads", "RPA-Artigo")
    time.sleep(5)
    pyautogui.moveTo(x=255, y=158)
    pyautogui.doubleClick()

def ordenarPorPdf():
    pyautogui.moveTo(x=569, y=103)
    pyautogui.leftClick()
    pyautogui.press("down")
    pyautogui.press("enter")

def abrirArquivo(indice):
    time.sleep(3)
    conteudo = None
    if (indice == 1):
        ordenarPorPdf()
        pyautogui.press("down", 1)
        pyautogui.press("up")
        pyautogui.moveTo(x=524, y=133)
        pyautogui.rightClick()
        pyautogui.press("down", 5)
        pyautogui.press("enter")
        pyautogui.hotkey('ctrl', 'c') #Copia conteúdo selecionado
        time.sleep(5)
        conteudo = paste()
        pyautogui.typewrite(conteudo)  # Digita o texto da variável conteúdo
        file = str(conteudo)
        if (validationArquivo(file)):
            pyautogui.press("enter")
            time.sleep(1)
            pyautogui.press("enter")
            return True
        else:
            return False
    else:
        pyautogui.press("down", 1)
        pyautogui.press("apps")
        pyautogui.press("down", 5)
        pyautogui.press("enter")
        pyautogui.hotkey('ctrl', 'c') #Copia conteúdo selecionado
        conteudo = paste()
        pyautogui.typewrite(conteudo)  # Digita o texto da variável conteúdo
        time.sleep(2)
        file = str(conteudo)
        if (validationArquivo(file)):
            pyautogui.press("enter")
            time.sleep(1)
            pyautogui.press("enter")
            return True
        else:
            return False


def salvarArquivo(indice):
    time.sleep(5)
    pyautogui.moveTo(x=31, y=50)
    pyautogui.leftClick()
    time.sleep(5)
    pyautogui.moveTo(x=83, y=239)
    pyautogui.leftClick()
    time.sleep(5)
    pyautogui.write("Pagina " + NumFile + " - Modificado")
    time.sleep(3)
    pyautogui.press("enter")
    global Status
    Status = "documento alterado"

def validationArquivo(file):
    #Verifica se è numerica e .pdf
    if (file[0].isnumeric()):
        if(file.__contains__(".pdf")):
            setInfoNameFile(file)
            return True
    else:
        return False

def setInfoNameFile(file):
    global NumFile
    global NameFIle
    if(file[0:2].isnumeric()):
        NumFile = file[0:2]
    else:
        NumFile = file[0]
    # Adiciona a variavel para posterior utilização
    NameFIle = file


# ----------------------------------------------------------------------------

def createPlanilha(indice):
    if (indice ==1):
        AutoBase.openWPSOfiice()
        #Entrando na aréa de planilha
        pyautogui.moveTo(x=232, y=138)
        pyautogui.leftClick()
        time.sleep(10)
        #Criando uma planilha nova
        pyautogui.moveTo(x=391, y=314)
        pyautogui.leftClick()
        time.sleep(7)
        criarCabecalho()
        time.sleep(2)
    else:
        time.sleep(2)
        searchFiles("Documentos", "Relatorio de execucao")
        time.sleep(5.5)
        pyautogui.moveTo(x=320, y=116)
        pyautogui.doubleClick()
        time.sleep(5)

        # pyautogui.moveTo(x=255, y=319)
        # pyautogui.doubleClick()
        # time.sleep(7)

    addInfoPlanilha(indice)
    salvarPlanilha(indice)
    if(indice != 1):
        openPastaOriginal()

def criarCabecalho():
    #Seleciona o campo do cabecalho
    mesclarCelulas()
    #Voltando pro começo da planilha
    pyautogui.moveTo(x=93, y=205)
    pyautogui.leftClick()
    #Selecionando
    with pyautogui.hold('shift'):
        pyautogui.press('right')
        pyautogui.press("down", 10)

    #Clica em Inserir
    pyautogui.moveTo(x=323, y=46)
    pyautogui.leftClick()
    time.sleep(1.5)
    #Clica na setinha de lado
    pyautogui.moveTo(x=1016, y=100)
    pyautogui.leftClick()
    time.sleep(2.5)
    #Clica em cabecalho e rodapé
    pyautogui.moveTo(x=359, y=94)
    pyautogui.leftClick()
    time.sleep(2.5)
    #Clica em cabeçalho personalizado
    pyautogui.moveTo(x=662, y=277)
    pyautogui.leftClick()
    time.sleep(2.5)
    #Adiciona as colunas
    pyautogui.write("Nome do Documento")
    pyautogui.press("tab")
    pyautogui.write("Status")
    time.sleep(2.5)
    pyautogui.moveTo(x=695, y=525)
    pyautogui.leftClick()
    time.sleep(2.5)
    #confirma a criação
    pyautogui.moveTo(x=622, y=550)
    pyautogui.leftClick()
    time.sleep(2.5)

def mesclarCelulas():
    cont = 0
    while cont<10:
        with pyautogui.hold('shift'):
            pyautogui.press(['right', 'right', 'right', 'right'])
            time.sleep(1)
            #Indo até a opção mesclar células
            pyautogui.moveTo(x=619, y=112)
            pyautogui.leftClick()
            pyautogui.moveTo(x=655, y=217)
            pyautogui.leftClick()

        pyautogui.press("down", 1)
        cont = cont + 1

    time.sleep(1)
    pyautogui.moveTo(x=93, y=205)
    pyautogui.leftClick()
    pyautogui.press('right')
    cont = 0
    while cont<10:
        with pyautogui.hold('shift'):
            pyautogui.press(['right', 'right'])
            time.sleep(1)
            #Indo até a opção mesclar células
            pyautogui.moveTo(x=619, y=112)
            pyautogui.leftClick()
            pyautogui.moveTo(x=655, y=217)
            pyautogui.leftClick()

        pyautogui.press("down", 1)
        cont = cont + 1

def addInfoPlanilha(indice):
    if (indice == 1):
        pyautogui.moveTo(x=93, y=205)
        pyautogui.leftClick()
    time.sleep(1.3)
    pyautogui.press("space")
    pyautogui.press("backspace")
    time.sleep(1.3)
    pyautogui.write(NameFIle)
    # pyautogui.press("enter")
    pyautogui.press("tab")
    pyautogui.write(Status)
    pyautogui.press("enter")
    time.sleep(2)

def salvarPlanilha(indice):
    time.sleep(1)
    pyautogui.moveTo(x=37, y=48)
    pyautogui.leftClick()
    time.sleep(1)
    if (indice == 1):
        pyautogui.moveTo(x=116, y=228)
        pyautogui.leftClick()
        pyautogui.moveTo(x=396, y=322)
        pyautogui.doubleClick()
        pyautogui.press("tab")
        time.sleep(1)
        pyautogui.write("Relatorio de execucao")
        time.sleep(1)
        pyautogui.press("enter")
    else:
        time.sleep(1)
        pyautogui.moveTo(x=95, y=189)
        pyautogui.leftClick()

    closeFile()

# ----------------------------------------------------------------------------

def searchFiles(nameDirectory, nameFile):
    #Acessando o diretorio
    pyautogui.moveTo(x=615, y=68)
    pyautogui.leftClick()
    time.sleep(2)
    pyautogui.write(nameDirectory)
    pyautogui.press("enter")
    #Pesquisando e acessando o arquivo
    time.sleep(2)
    pyautogui.moveTo(x=724, y=64)
    pyautogui.leftClick()
    time.sleep(5)
    pyautogui.write(nameFile)
    pyautogui.press("enter")

def closeFile():
    pyautogui.moveTo(x=345, y=11)
    time.sleep(1)
    pyautogui.leftClick()
    time.sleep(2)
