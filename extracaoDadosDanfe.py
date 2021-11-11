
# importa biblioteca
import xlsxwriter 
import fitz

# faz a leitura
with fitz.open("DANFES.pdf") as pdf:
    texto = ""
    pos = ""
    posTotal = ""
    NF = {}
    valor = {}
    cont = 0
    row = 0
    column = 0

    for pagina in pdf:
        texto = pagina.getText()
        pos = texto.find('N.')
        posTotal = texto.find('VALOR TOTAL DA NOTA')
        NF[cont] = texto[pos+3:pos+12]
        valor[cont] = texto[posTotal+28:posTotal+40]
        cont+=1
       
    workbook = xlsxwriter.Workbook('teste.xlsx') 
    worksheet = workbook.add_worksheet() 
  
    for item in NF : 
        worksheet.write(row + 1, column, NF[item]) 
        worksheet.write(row + 1, column + 1, "R$" + valor[item]) 
        
        row += 1

    worksheet.write(0, 0, "Número NF") 
    worksheet.write(0, 1, "Valor NF")   
    workbook.close() 



    