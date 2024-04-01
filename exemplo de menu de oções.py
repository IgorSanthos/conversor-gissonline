import tkinter as tk

def selecionar_opcao(valor):
    print("Opção selecionada:", valor)

# Lista de opções
opcoes = ["Opção 1", "Opção 2", "Opção 3", "Opção 4"]

# Janela principal
janela = tk.Tk()
janela.title("Exemplo de OptionMenu")

# Variável de controle para armazenar a opção selecionada
opcao_selecionada = tk.StringVar(janela)
opcao_selecionada.set(opcoes[0])  # Define o valor padrão como a primeira opção

# Cria o OptionMenu e associa à variável de controle
menu_opcoes = tk.OptionMenu(janela, opcao_selecionada, *opcoes, command=selecionar_opcao)
menu_opcoes.pack()

# Exibe a janela
janela.mainloop()
