
import xlsxwriter 
import fitz
import PySimpleGUI as sg


class TelaPython:
    def __init__(self):
        layout = [
            [sg.Text('Caminho do arquivo PDF:', size=(25,0)), sg.Input(size=(30,0),key='pdf')],
            [sg.Text('Exemplo: C:\caminhodoarquivo\\arquivo.pdf ', size=(50,0))],
            [sg.Text('Caminho do arquivo de saida XLSX:', size=(25,0)), sg.Input(size=(30,0),key='xlsx')],
            [sg.Text('Exemplo: C:\caminhodoarquivo\\arquivo.xlsx ', size=(50,0))],
            [sg.Button('Enviar dados')],
            #[sg.Output(size=(60,20))]
        ]
        self.janela = sg.Window("Dados dos Arquivos").layout(layout)
        self.button, self.values = self.janela.Read()

arquivos = TelaPython()        

with fitz.open(arquivos.values['pdf']) as pdf:
    texto = ""
    posNome = ""
    posPai = ""
    posMae = ""
    posNascimento = ""
    posRaca = ""
    posCargo = ""
    posAdmissao = ""
    nome = {}
    pai = {}
    mae = {}
    nascimento = {}
    raca = {}
    cargo = {}
    admissao = {}
    lista = {}
    Funcionario = []
    FiliacaoPai = []
    DataNascimento = []
    cont = 0
    row = 0
    column = 0

    for pagina in pdf:
        texto = pagina.getText()
        posNome = texto.find('Nome: ')
        posPai  = texto.find('Pai:')
        posMae  = texto.find('Mae:')
        posNascimento = texto.find('Nascimento:')
        posRaca = texto.find('Raça/Cor:')
        posCargo = texto.find('Cargo:')
        posAdmissao =  texto.find('Admissão:')

        nome[cont] = texto[posNome+6:posNome+45]
        pai[cont] = texto[posPai+5:posPai+45]
        mae[cont] = texto[posMae+3:posMae+12]
        nascimento[cont] = texto[posNascimento+12:posNascimento+22]
        raca[cont] = texto[posRaca+3:posRaca+12]
        cargo[cont] = texto[posCargo+3:posCargo+12]
        admissao[cont] = texto[posAdmissao+3:posAdmissao+12]
        cont+=1
    
    for i in nome:
        posQuebra = nome[i].find("\n")
        posLinhaErro = nome[i].find("ios")
        if posLinhaErro == -1:
            #print(str(row+1) + " - " + nome[i][0:posQuebra])
            Funcionario.append(nome[i][0:posQuebra])
            row += 1
    row = 0
    for i in pai:   
        posLinhaErro = pai[i].find("rios") 
        posLinhaVazia = pai[i].find("") 
        posQuebra = pai[i].find("\n")
    
        if posLinhaErro == -1: 
            nomePai = pai[i][0:posQuebra]
            if nomePai != " ":
                #print(str(row+1) + " - " + nomePai)
                FiliacaoPai.append(pai[i][0:posQuebra])
                row += 1
            elif nomePai == " ":
                # print(str(row+1) + " - NÃO RECONHECIDO") 
                FiliacaoPai.append("NÃO RECONHECIDO")
                row += 1
    
    for i in nascimento:
        posLinhaErro = nascimento[i].find("ta Inicial")
        if posLinhaErro == -1:
            DataNascimento.append(nascimento[i])

    NomesFuncs = {}
    for i in range(len(DataNascimento)):
        lista[Funcionario[i]]={'Nome':Funcionario[i] ,'Data_Nascimento':DataNascimento[i],'Filiacao':FiliacaoPai[i]}

    #print (("Nome: " + lista['ALINE CRISTINA DE CAMARGO']['Nome'] + "\n" + "Data de Nascimento: " + lista['ALINE CRISTINA DE CAMARGO']['Data_Nascimento'])+ "\n" + "Nome do Pai: " + (lista['ALINE CRISTINA DE CAMARGO']['Filiacao']))
    #print(lista)    
    
    workbook = xlsxwriter.Workbook(arquivos.values['xlsx']) 
    worksheet = workbook.add_worksheet()

    row = 0
    for i in range(len(DataNascimento)):
        worksheet.write(row + 1, column, Funcionario[i]) 
        worksheet.write(row + 1, column + 1,FiliacaoPai[i]) 
        worksheet.write(row + 1, column + 2,DataNascimento[i]) 
        row += 1    

    worksheet.write(0, 0, "Nome") 
    worksheet.write(0, 1, "Filiação Pai")   
    worksheet.write(0, 2, "Data Nascimento")   
    workbook.close() 

    