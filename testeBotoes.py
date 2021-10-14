import PySimpleGUI as sg

class TelaPython:
    def __init__(self):
        layout = [
            [sg.Text('Caminho do arquivo PDF'), sg.Input()],
            [sg.Text('Caminho do arquivo de saida XLSX'), sg.Input()],
            [sg.Button('Enviar dados')]
        ]
        janela = sg.Window("Dados dos Arquivos").layout(layout)
        self.button, self.values = janela.Read()

    def Iniciar(self):
        print(self.values)  
        valores = self.values  
    
    def imprimir(self,palavra):
        self.word = palavra
        print(self.word)    
    
tela = TelaPython()
#tela.Iniciar()
print(tela.values[0])
print(tela.values[1])

#tela.imprimir("não sei")

