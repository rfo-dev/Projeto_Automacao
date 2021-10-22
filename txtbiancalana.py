import os
import xlsxwriter 
import fitz
import linecache
import PySimpleGUI as sg

arquivoTXT = []
busca = "Nome:"
buscaCod = "Código:"
buscaRegime = "Regime:"
buscaSalario = "Salário:"
buscaAdmissao = "Admissão:"
buscaCargo = "Cargo:"
buscaEstCivil = "Est.Civi"
buscaNascimento = "Dt.Nasc."
buscaTitulo = "Título"
buscaPis = "PIS"
buscaCpf = "CPF:"
buscaRG = "Data Emissão:"
buscaCtps = "CTPS:"
buscaTipoPgto = "Tipo Pagto.:"
buscaNacionalidade = "Nacionalidade:"
buscaEscolaridade = "Grau Instrução:"
buscaSexo = "Nacionalidade:"
buscaEmpresa = "Empresa:"
aEmpresa = []
aNome = []
aCodigo = []
aRegime = []
aSalario = []
aAdmissao = []
aCargo = []
aEstCivil = []
aNascimento = []
aTitulo = []
aPis = []
aCpf = []
aRg = []
aCtps = []
aTipoPgto = []
aNacionalidade = []
aEscolaridade = []
aSexo = []


row = 0
column = 0

for file in os.listdir("./Biancalana/Ficha de Registro"):
    if file.endswith(".txt"):
        caminhoTXT = (os.path.join("./Biancalana/Ficha de Registro", file))
        arquivoTXT.append(caminhoTXT)
        
for i in range(len(arquivoTXT)):
    arquivo = open( arquivoTXT[i], 'r' )
    
    for linha in arquivo:
        valores = linha.split()
        juntos=  ' '.join(valores)
        #descobrir pocição das palavras
        posicao=juntos.find(busca)
        posicaoCod=juntos.find(buscaCod)
        posicaoRegime=juntos.find(buscaRegime)
        posicaoSalario = juntos.find(buscaSalario) 
        posicaoAdmissao = juntos.find(buscaAdmissao)  
        posicaoCargo = juntos.find(buscaCargo) 
        posicaoEstCivil = juntos.find(buscaEstCivil) 
        posicaoNascimento = juntos.find(buscaNascimento) 
        posicaoTitulo = juntos.find(buscaTitulo) 
        posicaoPis = juntos.find(buscaPis) 
        posicaoCpf = juntos.find(buscaCpf) 
        posicaoRg = juntos.find(buscaRG) 
        posicaoTipoPgto = juntos.find(buscaTipoPgto) 
        posicaoCtps = juntos.find(buscaCtps) 
        posicaoNacionalidade = juntos.find(buscaNacionalidade)
        posicaoEscolaridade = juntos.find(buscaEscolaridade) 
        posicaoSexo = juntos.find(buscaSexo) 
        posicaoEmpresa = juntos.find(buscaEmpresa)
        
        if (posicao != -1):
            nome = juntos[posicao+5:45]
            aNome.append(nome)
        
        if (posicaoEmpresa != -1):    
            empresa = juntos[posicaoEmpresa+8:posicaoEmpresa+60]
            aEmpresa.append(empresa) 
            
            
        if (posicaoCod != -1):
            codigo = juntos[posicaoCod+7:13]
            aCodigo.append(codigo)
                    
        if (posicaoRegime != -1):
            regime = juntos[posicaoRegime+7:15]
            aRegime.append(regime)        
            
        if (posicaoSalario != -1):
            salario = juntos[posicaoSalario+8:16]
            if (any(chr.isdigit() for chr in salario)):
                aSalario.append(salario)
            else:
                aSalario.append("VAZIO")
                
        if (posicaoAdmissao != -1):
            admissao = juntos[posicaoAdmissao+9:38]
            admissao = admissao.replace("Ca","")
            if len(admissao) != 0 :
                aAdmissao.append(admissao)
            
        if (posicaoCargo != -1):
            cargo = juntos[posicaoCargo+10:70]
            aCargo.append(cargo) 
        
        if (posicaoEstCivil != -1):
            estcivil= juntos[posicaoEstCivil+8:70]
            aEstCivil.append(estcivil) 
        
        if (posicaoNascimento != -1):
            nascimento = juntos[posicaoNascimento+8:18]
            nascimento = nascimento.replace(": Local: N", "VAZIO")
            aNascimento.append(nascimento)             
          
        if (posicaoTitulo != -1):
            titulo = juntos[posicaoTitulo+6:18]
            titulo = titulo.replace(":", "VAZIO")
            aTitulo.append(titulo)
        
        if (posicaoPis != -1):
            pis = juntos[posicaoPis+3:18]
            pis = pis.replace(": Data PIS: Ban", "VAZIO")
            aPis.append(pis)  
            
        if (posicaoCpf != -1):
            cpf = juntos[posicaoCpf+4:posicaoCpf+19]
            cpf = cpf.replace(".","")
            cpf = cpf.replace("/","")
            if (any(chr.isdigit() for chr in cpf)):
                aCpf.append(cpf) 
            else:
                aCpf.append("VAZIO")
                
        if (posicaoRg != -1):
            rg = juntos[1:15]
            rg = rg.replace(".","")
            rg = rg.replace("-","")
            rg = rg.replace("Data","")
            rg = rg.replace("G:", "")
            rg = rg.replace("Emissã","")
            rg = rg.replace("Dat","")
            rg = rg.replace("D","")
            rg = rg.replace("E","")
            rg = rg.replace("x", "0")
            rg = rg.replace("X", "0")
            rg = rg.replace("a", "0")
            if len(rg) > 2:
                aRg.append(rg)  
            else:
                aRg.append("VAZIO")    
                 
        if (posicaoTipoPgto != -1):
            tipopgto = juntos[posicaoTipoPgto+7:15]
            aTipoPgto.append(tipopgto)        
            
        if (posicaoCtps != -1):
            ctps = juntos[posicaoCtps+7:15]
            aCtps.append(ctps)  
            
        if (posicaoNacionalidade != -1):
            nacionalidade = juntos[posicaoNacionalidade+7:15]
            aNacionalidade.append(nacionalidade)         
            
        if (posicaoEscolaridade != -1):
            escolaridade = juntos[posicaoEscolaridade+7:15]
            aEscolaridade.append(escolaridade)       
            
        if (posicaoSexo != -1):
            sexo = juntos[posicaoSexo+7:15]
            aSexo.append(sexo)       

    arquivo.close()

workbook = xlsxwriter.Workbook("CadFunc.xlsx") 
worksheet = workbook.add_worksheet()
row = 0
for i in range(len(aNome)):
    worksheet.write(row + 1, column, aNome[i]) 
    worksheet.write(row + 1, column + 1,aCodigo[i]) 
    worksheet.write(row + 1, column + 2,aRegime[i]) 
    worksheet.write(row + 1, column + 3,aSalario[i]) 
    worksheet.write(row + 1, column + 4,aAdmissao[i]) 
    worksheet.write(row + 1, column + 5,aCargo[i])
    worksheet.write(row + 1, column + 6,aEstCivil[i])  
    worksheet.write(row + 1, column + 7,aNascimento[i]) 
    worksheet.write(row + 1, column + 8,aTitulo[i]) 
    worksheet.write(row + 1, column + 9,aPis[i]) 
    worksheet.write(row + 1, column + 10,aCpf[i]) 
    worksheet.write(row + 1, column + 11,aRg[i]) 
    row += 1    
worksheet.write(0, 0, "Nome") 
worksheet.write(0, 1, "Codigo")   
worksheet.write(0, 2, "Regime Contratação")   
worksheet.write(0, 3, "Salario")  
worksheet.write(0, 4, "Admissão")    
worksheet.write(0, 5, "Cargo")   
worksheet.write(0, 6, "Estado Civil")   
worksheet.write(0, 7, "Data de Nascimento")   
worksheet.write(0, 8, "Titulo")  
worksheet.write(0, 9, "PIS") 
worksheet.write(0, 10, "CPF")     
worksheet.write(0, 11, "RG")    
workbook.close()     

#print (len(aAdmissao))
#print (len(aNome))
#print (len(aCodigo))
#print (len(aRegime))
#print (len(aSalario))
#print (len(aCargo))
#print (len(aEstCivil))
#print (len(aNascimento))
#print (len(aTitulo))
#print (len(aPis))
#print(len(aCpf))
#print (len(aRg))
#print (aRg)
#print(aEmpresa)

#print (arquivo)

