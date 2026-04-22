import tkinter as tk
import csv
import json
import os
from openpyxl import Workbook 

                   
produtos = []  

#Funçao de Adicionar Produtos
class Produto:
    def __init__(self, nome, preco, categoria, estoque, fornecedor):
        self.nome = nome
        self.preco = preco
        self.categoria = categoria
        self.estoque = estoque
        self.fornecedor = fornecedor
        
    
   

def salvar_dados():
    with open("produtos.json", "w") as arquivo:
        json.dump([produto.__dict__ for produto in produtos], arquivo, indent=4)


def carregar_dados():
    if os.path.exists("produtos.json"):
        with open("produtos.json", "r") as arquivo:
            dados = json.load(arquivo)

            for item in dados:
                produto = Produto(
                    item["nome"],
                    item["preco"],
                    item["categoria"],
                    item["estoque"],
                    item.get("fornecedor", "Desconhecido")
                )

                produtos.append(produto)
                lista_produtos.insert(
    tk.END,
    f"{produto.nome} - {produto.estoque} un - {produto.fornecedor}"
)
                

def adicionar_produto():
    try:
        nome = entrada_nome.get()
        preco = float(entrada_preço.get())
        categoria = entrada_categ.get()
        estoque = int(entrada_estoq.get())
        fornecedor = entrada_fornecedor.get()

        produto = Produto(nome, preco, categoria, estoque, fornecedor)
        produtos.append(produto)

        lista_produtos.insert(
    tk.END,
    f"{produto.nome} - {produto.estoque} un - {produto.fornecedor}"
)

        print("Produto Adicionado:", nome)
        salvar_dados()


    except ValueError:
        print("Erro: Preço ou estoque inválido")                




def remover_produto(): 
    selecionado = lista_produtos.curselection()    #### curselection() retorna uma tupla com os índices dos itens selecionados
    if not selecionado:
        print("Nenhum produto selecionado ")
        return
    salvar_dados()
    index = selecionado[0] # Pega o primeiro índice selecionado
    produtos.pop(index) # Remove o produto da lista
    lista_produtos.delete(index)    # Remove o item da Listbox
    salvar_dados()
     
              
              
              
    #Gerar Planilha

def gerar_plan():
    plan = Workbook()
    aba = plan.active 

    aba.append(["Nome", "Preço", "Categoria", "Estoque", "Fornecedor"])

    for produto in produtos:
        aba.append([
            produto.nome,
            produto.preco,
            produto.categoria,
            produto.estoque,
            produto.fornecedor
            
            
        ])
        
    

 
    plan.save("SistemaM2.xlsx")
    print("Planilha Criada") 




    #Janela
jan = tk.Tk()
jan.title("SistemaM2")
jan.geometry("700x700")




    #Campos de Entrada
tk.Label(jan, text="Nome").pack()
entrada_nome = tk.Entry(jan)
entrada_nome.pack()

tk.Label(jan, text="Preço").pack()
entrada_preço = tk.Entry(jan)
entrada_preço.pack()

tk.Label(jan, text="Categoria").pack()
entrada_categ = tk.Entry(jan)
entrada_categ.pack()

tk.Label(jan, text="Estoque").pack()
entrada_estoq = tk.Entry(jan)
entrada_estoq.pack()
tk.Label(jan, text="Fornecedor").pack()
entrada_fornecedor = tk.Entry(jan)
entrada_fornecedor.pack()           




    #Botões Janela
tk.Button(jan, text="Adicionar", command = adicionar_produto).pack(pady=40)

tk.Button(jan, text="Remover Produto", command=remover_produto).pack(pady=10)

tk.Button(jan, text="Salvar Excel", command=gerar_plan).pack(pady=10)

lista_produtos = tk.Listbox(jan, width=40   , height=20)
lista_produtos.pack()



carregar_dados()   
jan.mainloop()

