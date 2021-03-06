from math import asin
import os
from typing import Tuple
import xlsxwriter 
import fitz
import linecache
import PySimpleGUI as sg
import time

pai = ""
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
buscaEndereco = "Endereço:"
buscaBairro = "Bairro"
buscaMunicipio = "Município:"
buscaLocalNascimento = "Dt.Nasc."
buscaUF = "UF:"
buscaCEP = "CEP:"
buscaSindicato = "Sindicato"
buscaPai = "Pai "
buscanPais = "Nome "
buscaMae = "Mãe "
aMae = []
anPais = []
aPai = []
aSindicato = []
aLocalNascimento = []
aMunicipio= []
aUF = []
aCEP = []
aBairro = []
aEndereco = []
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
aEmpresaLista = []
totalArquivos = []

row = 0
column = 0

t_ini = time.time()
for file in os.listdir("./Origem_dados/Biancalana/Ficha de Registro"):
    if file.endswith(".txt"):
        caminhoTXT = (os.path.join("./Origem_Dados/Biancalana/Ficha de Registro", file))
        arquivoTXT.append(caminhoTXT)
        
for i in range(len(arquivoTXT)):
    arquivo = open( arquivoTXT[i], 'r' )
    cont = 0
    for linha in arquivo:
        valores = linha.split()
        juntos=  ' '.join(valores)
        #descobrir posição das palavras
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
        posicaoEndereco = juntos.find(buscaEndereco)
        posicaoBairro= juntos.find(buscaBairro)
        posicaoMunicipio = juntos.find(buscaMunicipio)
        posicaoUF = juntos.find(buscaUF)
        posicaoCEP = juntos.find(buscaCEP)
        posicaoLocalNascimento = juntos.find(buscaLocalNascimento)
        posicaoSindicato = juntos.find(buscaSindicato)
        posicaoPai = juntos.find(buscaPai)
        posicaonPais = juntos.find(buscanPais)
        posicaoMae = juntos.find(buscaMae)
        
        if (posicao != -1):
            nome = juntos[posicao+5:45]
            aNome.append(nome)
            cont+=1
        
        if (posicaoEmpresa != -1):    
            empresa = juntos[posicaoEmpresa+8:posicaoEmpresa+60]
            
            #if empresa not in aEmpresa:
                #aEmpresa.append(empresa) 
            
        if (posicaoCod != -1):
            codigo = juntos[posicaoCod+7:13]
            aCodigo.append(codigo)
                    
        if (posicaoRegime != -1):
            regime = juntos[posicaoRegime+7:15]
            aRegime.append(regime)        
            
        if (posicaoSalario != -1):
            salario = juntos[posicaoSalario+8:16]
            if (any(chr.isdigit() for chr in salario)):
                salario = salario.replace("Ti", "")
                salario = salario.replace("Tip", "")
                salario = salario.replace("T", "")
                aSalario.append(salario)
            else:
                aSalario.append("VAZIO")
                
        if (posicaoAdmissao != -1):
            admissao = juntos[posicaoAdmissao+9:39]
            admissao = admissao.replace("Ca","")
            admissao = admissao.replace("r","")
            admissao.lstrip()
            admissao.rstrip()
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
            tipopgto = juntos[posicaoTipoPgto+13:posicaoTipoPgto+20]
            aTipoPgto.append(tipopgto)        
            
        if (posicaoCtps != -1):
            ctps = juntos[posicaoCtps+6:posicaoCtps+25]
            aCtps.append(ctps)  
  
        if (posicaoNacionalidade != -1):
            nacionalidade = juntos[posicaoNacionalidade+15:posicaoNacionalidade+26]
            nacionalidade = nacionalidade.replace("Sexo: MASCU","")
            nacionalidade = nacionalidade.replace("Sexo: FEMIN","")
            if nacionalidade == '':
                aNacionalidade.append("VAZIO")
            else:    
                aNacionalidade.append(nacionalidade)  
                  
        if (posicaoEscolaridade != -1):
            escolaridade = juntos[posicaoEscolaridade+15:posicaoEscolaridade+33]
            escolaridade = escolaridade.replace("Est.CiviOUTROS", "")
            escolaridade = escolaridade.replace("INCOMPLE","INCOMPLETO")
            escolaridade = escolaridade.replace("INCOMPLET","INCOMPLETO")
            if escolaridade == " ":
                aEscolaridade.append("VAZIO")
            else:    
                aEscolaridade.append(escolaridade)       
            
        if (posicaoSexo != -1):
            sexo = juntos[posicaoSexo+29:posicaoSexo+45]
            if len(sexo) < 3:
                sexo = juntos[posicaoSexo+20:posicaoSexo+55]
                sexo = sexo.replace("Sexo:", "")
                sexo = sexo.replace(" Sexo:", "")
                sexo = sexo.replace("S", "")
                sexo = sexo.replace(" Sexo: ", "")
                aSexo.append(sexo)
            else:    
                aSexo.append(sexo)
                
        if (posicaoEndereco != -1):
            endereco = juntos[posicaoEndereco+9:posicaoEndereco+70]
            aEndereco.append(endereco) 
        
        if (posicaoBairro != -1):
            bairro = juntos[posicaoBairro+6:30]
            aBairro.append(bairro)
             
        if (posicaoMunicipio != -1):
            municipio = juntos[posicaoMunicipio+10:posicaoMunicipio+25]
            municipio = municipio.replace(" UF: S"," ")
            municipio = municipio.replace("SP CE"," ")
            municipio = municipio.replace("UF: CEP: DDD"," ")
            municipio = municipio.replace("P CE","")
            municipio = municipio.replace(":" , "")
            municipio = municipio.replace("UF", "")
            municipio = municipio.replace("P ","")
            if municipio != "   ":
                aMunicipio.append(municipio)
            else:
                aMunicipio.append("VAZIO")     
            
        if (posicaoUF != -1):
            UF = juntos[posicaoUF+4:posicaoUF+6]
            aUF.append(UF)   
            
        if (posicaoCEP != -1):
            CEP = juntos[posicaoCEP+4:posicaoCEP+13]
            CEP = CEP.replace(" DDD: Fon", "VAZIO")
            aCEP.append(CEP)
            
        if (posicaoLocalNascimento != -1):
            localNascimento = juntos[posicaoLocalNascimento+30:posicaoLocalNascimento+43]
            localNascimento = localNascimento.replace("Nacional", "")
            localNascimento = localNascimento.replace("Naciona", "")
            localNascimento = localNascimento.replace("Nacio", "")
            localNascimento = localNascimento.replace("Naci", "")
            localNascimento = localNascimento.replace("Nac", "")
            localNascimento = localNascimento.replace("Na", "")
            localNascimento = localNascimento.replace("onalidade: BR", "VAZIO")
            localNascimento = localNascimento.replace("onalidade: OU", "VAZIO")
            localNascimento = localNascimento.replace("oalidade: Se", "VAZIO")
            localNascimento = localNascimento.replace(": Sexo: MASCU", "VAZIO")
            localNascimento = localNascimento.replace("n", "")
            aLocalNascimento.append(localNascimento)     
            
        if (posicaoSindicato != -1):
            sindicato = juntos[posicaoSindicato+14:posicaoSindicato+100]
            aSindicato.append(sindicato)   
        
        if (posicaonPais) != -1 :
            aPai.append("VAZIO")
            aMae.append("VAZIO")
         
        if (posicaoPai) != -1 :
            pai = linha[0:40]
            indice = len(aPai)-1  
            if (aPai[indice]) == "VAZIO" :
                pai = pai.strip()
                aPai.pop()
                aPai.append(pai)
                
        if (posicaoMae) != -1 :
            mae = linha[0:45]
            indice2 = len(aMae)-1  
            if (aMae[indice2]) == "VAZIO" :
                mae = mae.strip()
                aMae.pop()
                aMae.append(mae)        
                         
    x = 0        
    while x < cont:
        aEmpresa.append(empresa)  
        x+=1
    
    arquivo.close()

workbook = xlsxwriter.Workbook("CadFunc.xlsx") 
worksheet = workbook.add_worksheet()
row = 0
for i in range(len(aNome)):
    worksheet.write(row + 1, column, aEmpresa[i]) 
    worksheet.write(row + 1, column + 1, aNome[i]) 
    worksheet.write(row + 1, column + 2,aCodigo[i]) 
    worksheet.write(row + 1, column + 3,aRegime[i]) 
    worksheet.write(row + 1, column + 4,aSalario[i]) 
    worksheet.write(row + 1, column + 5,aTipoPgto[i]) 
    worksheet.write(row + 1, column + 6,aAdmissao[i]) 
    worksheet.write(row + 1, column + 7,aCargo[i])
    worksheet.write(row + 1, column + 8,aEstCivil[i])  
    worksheet.write(row + 1, column + 9,aNascimento[i]) 
    worksheet.write(row + 1, column + 10,aTitulo[i]) 
    worksheet.write(row + 1, column + 11,aPis[i]) 
    worksheet.write(row + 1, column + 12,aCpf[i]) 
    worksheet.write(row + 1, column + 13,aRg[i]) 
    worksheet.write(row + 1, column + 14,aCtps[i]) 
    worksheet.write(row + 1, column + 15,aNacionalidade[i]) 
    worksheet.write(row + 1, column + 16,aEscolaridade[i]) 
    worksheet.write(row + 1, column + 17,aSexo[i]) 
    worksheet.write(row + 1, column + 18,aEndereco[i]) 
    worksheet.write(row + 1, column + 19,aBairro[i]) 
    worksheet.write(row + 1, column + 20,aMunicipio[i]) 
    worksheet.write(row + 1, column + 21,aUF[i]) 
    worksheet.write(row + 1, column + 22,aCEP[i]) 
    worksheet.write(row + 1, column + 23,aLocalNascimento[i]) 
    worksheet.write(row + 1, column + 24,aSindicato[i]) 
    worksheet.write(row + 1, column + 25,aPai[i]) 
    worksheet.write(row + 1, column + 26,aMae[i]) 
    
    row += 1
worksheet.write(0, 0, "Empresa")        
worksheet.write(0, 1, "Nome") 
worksheet.write(0, 2, "Codigo")   
worksheet.write(0, 3, "Regime Contratação")   
worksheet.write(0, 4, "Salario")  
worksheet.write(0, 5, "Tipo Pagamento")  
worksheet.write(0, 6, "Admissão")    
worksheet.write(0, 7, "Cargo")   
worksheet.write(0, 8, "Estado Civil")   
worksheet.write(0, 9, "Data de Nascimento")   
worksheet.write(0, 10, "Titulo")  
worksheet.write(0, 11, "PIS") 
worksheet.write(0, 12, "CPF")     
worksheet.write(0, 13, "RG")    
worksheet.write(0, 14, "CTPS")   
worksheet.write(0, 15, "Nacionalidade")  
worksheet.write(0, 16, "Escolaridade")  
worksheet.write(0, 17, "Sexo")  
worksheet.write(0, 18, "Endereço") 
worksheet.write(0, 19, "Bairro")  
worksheet.write(0, 20, "Municipio")
worksheet.write(0, 21, "UF")    
worksheet.write(0, 22, "CEP")  
worksheet.write(0, 23, "Local de Nascimento") 
worksheet.write(0, 24, "Sindicato") 
worksheet.write(0, 25, "Filiação - Pai") 
worksheet.write(0, 26, "Filiação - Mãe") 
workbook.close() 

print("Executado em: " , time.time()-t_ini)    

#print (len(aAdmissao))
#print (aAdmissao)
#print (len(aNome))
#print (len(aCodigo))
#print (len(aRegime))
#print (len(aSalario))
#print(aSalario)
#print(aCargo)
#print (len(aCargo))
#print (len(aEstCivil))
#print (len(aNascimento))
#print (len(aTitulo))
#print (len(aPis))
#print(len(aCpf))
#print (len(aRg))
#print (aRg)
#print(aEmpresa)
#print(len(aEmpresa))
#print (arquivo)
#print(len(listEmpresas))
#print (aTipoPgto)
#print(len(aTipoPgto))
#print(aCtps)
#print(aNacionalidade)
#print(len(aNacionalidade))
#print(aEscolaridade)
#print(len(aEscolaridade))
#print(aSexo)
#print (len(aSexo))
#print(aEndereco)
#print (len(aEndereco))
#print (aBairro)
#print(len(aBairro))
#print(aMunicipio)
#print (len(aMunicipio))
#print (aCEP)
#print(len(aCEP))
#print(aUF)
#print(len(aUF))
#print (aLocalNascimento)
#print(len(aLocalNascimento))
#print(aSindicato)
#print (len(aSindicato))
#print (arquivo) 
#print(aPai)
#print(len(aPai))
#print(aMae)
#print(len(aMae))