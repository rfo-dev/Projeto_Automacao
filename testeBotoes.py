import PySimpleGUI as sg

class TelaPython:
    def __init__(self):
        layout = [
            [sg.Text('Nome'), sg.Input()],
            [sg.Text('Idade'), sg.Input()],
            [sg.Button('Enviar dados')]
        ]
        janela = sg.Window("Dados do usuário").layout(layout)
        self.button, self.values = janela.Read()

    def Iniciar(self):
        print(self.values)  
        valores = self.values  
    
    def imprimir(self,palavra):
        self.word = palavra
        print(self.word)    
    
tela = TelaPython()
tela.Iniciar()
print(tela.values)

#tela.imprimir("não sei")