import pyautogui
import time
from pyperclip import paste

def openExplorerFile():

    # Mostrar mensagem de Alerta sobre o inicio da automação
    pyautogui.alert("Automação iniciada. Por favor, não use o teclado "
                    "e nem mecha no mouse, para evitar erros!")

    #Abrindo a Aréa de Trabalho
    pyautogui.hotkey("winleft", "d")

    #Abrindo o Explorador de Arquivos
    pyautogui.press("winleft")
    time.sleep(1.5)
    pyautogui.write("Explorador de Arquivos")
    time.sleep(5)
    pyautogui.press("enter")

    #Maximizando a tela
    time.sleep(2.5)
    pyautogui.moveTo(x=858, y=152)
    pyautogui.leftClick()

    #Acessando e informando o caminho do arquivo zip
    time.sleep(1.5)
    pyautogui.moveTo(x=615, y=68)
    pyautogui.leftClick()
    time.sleep(2)
    pyautogui.write("Downloads")
    pyautogui.press("enter")

    #Pesquisando por RPA
    time.sleep(2)
    pyautogui.moveTo(x=724, y=64)
    pyautogui.leftClick()
    time.sleep(1)
    pyautogui.write("RPA")
    pyautogui.press("enter")

    #Selecionando e abrindo o arquivo .zip
    time.sleep(5.5)
    pyautogui.moveTo(x=502, y=163)
    pyautogui.doubleClick()

    entra = True
    indice = 1
    while entra:
        if(abrirArquivo(indice)):
            time.sleep(5)
            salvarArquivo(indice)
            time.sleep(2)
            closeFile()
            backToExplorer()
            indice = indice + 1
            entra = True
        else:
            pyautogui.alert("Impossivel Abrir arquivos .docx ou arquivos que não começam com um Númeral!!")
            continue



def abrirArquivo(indice):
    if (indice == 1):
        conteudo = None
        pyautogui.press("down", 1)
        pyautogui.press("up")
        pyautogui.moveTo(x=524, y=133)
        pyautogui.rightClick()
        pyautogui.press("down", 5)
        pyautogui.press("enter")
        pyautogui.hotkey('ctrl', 'c') #Copia conteúdo selecionado
        conteudo = paste()
        pyautogui.typewrite(conteudo)  # Digita o texto da variável conteúdo
        time.sleep(5)
        file = str(conteudo)
        # setNameFile(file)
        if (validationArquivo(file)):
            time.sleep(3)
            pyautogui.press("enter")
            time.sleep(1.5)
            pyautogui.press("enter")
            return True
        else:
            return False

    else:
        conteudo = None
        pyautogui.press("down", 1)
        # pyautogui.press("up", 1)
        # pyautogui.press("enter")
        pyautogui.press("apps")
        pyautogui.press("down", 5)
        pyautogui.press("enter")
        pyautogui.hotkey('ctrl', 'c') #Copia conteúdo selecionado
        conteudo = paste()
        pyautogui.typewrite(conteudo)  # Digita o texto da variável conteúdo
        time.sleep(2)
        file = str(conteudo)
        # setNameFile(file)
        if (validationArquivo(file)):
            time.sleep(3)
            pyautogui.press("enter")
            time.sleep(1.5)
            pyautogui.press("enter")
            return True
        else:
            return False

def salvarArquivo(indice):
    time.sleep(5)
    pyautogui.moveTo(x=31, y=50)
    pyautogui.leftClick()
    time.sleep(2)
    pyautogui.moveTo(x=83, y=239)
    pyautogui.leftClick()
    time.sleep(2)
    pyautogui.write("Pagina " + str(indice)
                    + " - Modificado")

    pyautogui.press("enter")

def backToExplorer():
    pyautogui.hotkey("alt", "tab")
    time.sleep(3)

def closeFile():
    pyautogui.moveTo(x=345, y=11)
    time.sleep(1)
    pyautogui.leftClick()
    time.sleep(2)

def validationArquivo(file):
    if (file[0].isnumeric()):

        if(file.__contains__(".pdf")):
            return True
    else:
        return False
