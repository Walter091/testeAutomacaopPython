import pyautogui as py
from pyautogui import press, leftClick, rightClick, doubleClick, moveTo
from nucleo import InfoMsg, Atalhos_Teclado
import AutoBase
from time import sleep
from pyperclip import paste

class TesteAutomacao():

    path = "C:\\Users\\hp\\Documents\\RPA-Artigo"
    pathWithName = "C:\\Users\\hp\\Documents\\RPA-Artigo\\Relatorio de execucao.xlsx"
    pathThis = "C:\\Users\\hp\\PycharmProjects\\TesteAutGPAguiaBr\\automacaoTeste\\Relatorio de execucao.xlsx"

    nomePlanilha = "Relatorio de execucao.xlsx"
    def __init__(self, numArquivo, nomeArquivo, status):
        self.__numArquivo = numArquivo
        self.__nomeArquivo = nomeArquivo
        self.__status = status

    def Index(self):
        #mensagem alert
        InfoMsg.msgInicializacao("Automação iniciada. Por favor, não use o teclado "
                        "e nem mecha no mouse, para evitar erros!")
        print('Automação iniciada')
        #Abrindo a Aréa de Trabalho
        AutoBase.toAreaDeTrabalho()
        #Abrindo o Explorador de Arquivos
        AutoBase.openExplorerFiles()
        #Maximizando a tela
        sleep(1)
        moveTo(x=858, y=152)
        leftClick()

        #Acessando e informando o caminho do arquivo zip
        self.openPastaOriginal()

        entra = True
        indice = 1
        while entra:
            if indice <= 10:
                if self.abrirArquivo(indice):
                    sleep(2)
                    self.salvarArquivo()
                    sleep(2)
                    self.closeFile()
                    Atalhos_Teclado.altTab()
                    if indice == 1:
                        AutoBase.criarPlanilha(self.nomePlanilha)

                    self.addInfoPlanilha(indice)
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

        AutoBase.moverPlanilha(self.pathThis, self.path)
        InfoMsg.msgInicializacao("Automação encerrada. Obrigado e volte Sempre!")

    def openPastaOriginal(self):
        self.searchFiles("Downloads", "RPA-Artigo")
        sleep(2)
        moveTo(x=255, y=158)
        doubleClick()

    def ordenarPorPdf(self):
        moveTo(x=569, y=103)
        leftClick()
        press("down")
        press("enter")

    def ordenarPorNome(self):
        moveTo(x=262, y=99)
        leftClick()

    def abrirArquivo(self, indice):
        sleep(2)
        conteudo = None
        if indice == 1:
            # self.ordenarPorPdf()
            # self.ordenarPorNome()
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
            if self.validationArquivo(file):
                press("enter")
                sleep(1)
                press("enter")
                return True
            else:
                return False
        else:
            press("down", 1)
            press("apps")
            press("down", 5)
            press("enter")
            py.hotkey('ctrl', 'c') #Copia conteúdo selecionado
            conteudo = paste()
            py.typewrite(conteudo)  # Digita o texto da variável conteúdo
            sleep(2)
            file = str(conteudo)
            if self.validationArquivo(file):
                press("enter")
                sleep(1)
                press("enter")
                return True
            else:
                return False

    def validationArquivo(self, file):
        #Verifica se è numerica e .pdf
        if file[0].isnumeric():
            if file.__contains__(".pdf"):
                self.setInfoNameFile(file)
                return True
        else:
            return False

    def setInfoNameFile(self, file):
        # global numFile
        # global nameFile
        if file[0:2].isnumeric():
            self.set_numArquivo(file[0:2])
        else:
            self.set_numArquivo(file[0])
        # Adiciona a variavel para posterior utilização
        self.set_nomeArquivo(file)

    def salvarArquivo(self):
        sleep(5)
        moveTo(x=31, y=50)
        leftClick()
        sleep(3)
        moveTo(x=83, y=239)
        leftClick()
        sleep(4)
        py.write("Pagina " + self.__numArquivo + " - Modificado")
        sleep(3)
        press("enter")
        global Status
        self.set_status("documento alterado")

        print("Arquivo Salvo")

    # def mesclarCelulasPlanilha(self, indice, planilha):
    #     # Mesclar a Célula do nome do documento
    #     AutoBase.mesclarCelulas(planilha, 'A1:D1', 1, 1, 1, 4)
    #     # Mesclar a Célula do status do documento
    #     AutoBase.mesclarCelulas(planilha, 'E1:G1', 1, 1, 5, 8)
    #     # Mesclar o resto das celulas
    #     i = 1
    #     linhaInicial = 2
    #     linhaFinal = 2
    #     while i < 10:
    #         AutoBase.mesclarCelulas(planilha, 'A' + str(linhaInicial) + ':D' + str(linhaFinal), linhaInicial, linhaFinal, 1, 4)
    #         AutoBase.mesclarCelulas(planilha, 'E' + str(linhaInicial) + ':G' + str(linhaFinal), linhaInicial, linhaFinal, 5, 8)
    #         linhaInicial = linhaInicial + 1
    #         linhaFinal = linhaFinal + 1
    #         i = i + 1

    def setDimensaoCelula(self, planilha, larguraColuna1, larguraColuna2):
        actv = planilha.active
        actv.column_dimensions['A'].width = larguraColuna1
        actv.column_dimensions['B'].width = larguraColuna2

    def addInfoPlanilha(self, indice):
        planilha = AutoBase.carregarPlanilha(self.nomePlanilha)
        documento = planilha['Sheet']
        if indice == 1:
            # self.mesclarCelulasPlanilha(indice, planilha)
            self.setDimensaoCelula(planilha, 52, 19)
            documento.append(["Nome do Documento", "Status"])


        documento.append([self.__nomeArquivo, self.__status])
        print("Add info Sucess")
        AutoBase.salvarPlanilha(planilha, self.nomePlanilha)

    # ---------------------------------------------------------------------

    def searchFiles(self, nameDirectory, nameFile):
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

    def closeFile(self):
        moveTo(x=345, y=11)
        sleep(1)
        leftClick()
        sleep(2)

    # ---------------------------------------------------------------------

    def get_numArquivo(self):
        return self.__numArquivo

    def set_numArquivo(self, numArquivo):
        self.__numArquivo = numArquivo

    def get_nomeArquivo(self):
        return self.__nomeArquivo

    def set_nomeArquivo(self, nomeArquivo):
        self.__nomeArquivo = nomeArquivo

    def get_status(self):
        return self.__status

    def set_status(self, status):
        self.__status = status
