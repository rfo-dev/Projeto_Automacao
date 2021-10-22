
import xlsxwriter 
import fitz
    

with fitz.open(r"C:\Users\rafael.oliveira\Documents\Projeto_Automacao\FOLHA REF 01.2017 CAMPEIRO MATRIZ.pdf") as pdf:
    texto = ""
    posCompetencia = ""
    posCodEmpresa = ""
    posNomeEmpresa = ""
    posCNPJ = ""
    posCodFuncionario = ""
    posNomeFuncionario = ""
    posCodEvento = ""
    posEvento = ""
    posMultiplicador = ""
    posValorProvento = ""
    nome = {}
    codFuncionario= {}
    mae = {}
    nascimento = {}
    raca = {}
    cargo = {}
    admissao = {}
    lista = {}
    Funcionario = []
    codFunc = []
    DataNascimento = []
    cont = 0
    row = 0
    column = 0
    substr = "001.000.000"
    res = {}

    for pagina in pdf:
        texto = pagina.getText()
        res[cont] = [i for i in range(len(texto)) if texto.startswith(substr, i)]
        cont+=1
    valores= {}
    for i in res:
        if res[i]:
            valores[i+1] = res[i]
            
    print (valores[9][1])

    