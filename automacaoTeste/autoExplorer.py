import pyautogui
import time

from pyperclip import paste

CONT_DOWN_FILE = 1;

def openExplorerFile():


    # Mostrar mensagem de Alerta sobre o inicio da automação
    pyautogui.alert("Automação iniciada. Por favor, não use o teclado "
                    "e nem mecha no mouse, para evitar erros!")

    #Fechando possivéis janelas



    #Abrindo a Aréa de Trabalho
    pyautogui.hotkey("winleft", "d")

    #Abrindo o Explorador de Arquivos
    pyautogui.press("winleft")
    time.sleep(1)
    pyautogui.write("Explorador de Arquivos")
    time.sleep(2)
    pyautogui.press("enter")

    #Maximizando a tela
    time.sleep(2)
    pyautogui.moveTo(x=858, y=152)
    pyautogui.leftClick()

    #Acessando e informando o caminho do arquivo zip
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
    time.sleep(2)
    pyautogui.moveTo(x=502, y=163)
    pyautogui.doubleClick()

    while True:
        if (abrirArquivo(1)):
            CONT_DOWN_FILE + 1
            # salvarArquivo()
        else:
            pyautogui.alert("Impossivel Abrir arquivos .docx ou arquivos que não começam com um Númeral!!")



def abrirArquivo(indice):
    if (indice == 1):
        pyautogui.press("down", 1)
        pyautogui.press("up")
        pyautogui.moveTo(x=524, y=133)
        pyautogui.rightClick()
        pyautogui.press("down", 5)
        pyautogui.press("enter")
        pyautogui.hotkey('ctrl', 'c') #Copia conteúdo selecionado
        conteudo = paste()
        pyautogui.typewrite(conteudo)  # Digita o texto da variável conteúdo
        pyautogui.alert(conteudo)
        time.sleep(5)
        file = str(conteudo)
        if (validationArquivo(file)):
            pyautogui.press("enter")
            pyautogui.press("esc")
            pyautogui.doubleClick()
            return True
        else:
            return False
    else:
        pyautogui.press("down", CONT_DOWN_FILE)
        pyautogui.press("up")
        pyautogui.press("enter")
        pyautogui.press("apps")
        pyautogui.press("down", 5)
        pyautogui.press("enter")
        pyautogui.hotkey('ctrl', 'c') #Copia conteúdo selecionado
        conteudo = paste()
        pyautogui.typewrite(conteudo)  # Digita o texto da variável conteúdo
        pyautogui.alert(conteudo)
        time.sleep(5)
        file = str(conteudo)
        if (validationArquivo(file)):
            pyautogui.press("enter")
            pyautogui.press("esc")
            pyautogui.doubleClick()
            return True
        else:
            return False

# def salvarArquivo():
#     pyautogui.press("apps")

def validationArquivo(file):
    if (file[0].isnumeric()):
        if(file.__contains__(".pdf")):
            return True
    else:
        return False



