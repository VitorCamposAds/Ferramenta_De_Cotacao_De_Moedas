import tkinter as tk

def enviar():
    print(var_aviao.get())

janela = tk.Tk()

var_aviao = tk.StringVar(value="Nenhuma Opção Marcada")

botao_classe_economica = tk.Radiobutton(text="Classe Econômica", variable=var_aviao, value="Classe Econômica", command=enviar)
botao_classe_executiva = tk.Radiobutton(text="Classe Executiva", variable=var_aviao, value="Classe Executiva", command=enviar)
botao_primeira_classe = tk.Radiobutton(text="Primeira Classe", variable=var_aviao, value="Primeira Classe", command=enviar)
botao_classe_economica.grid(row=0, column=0)
botao_classe_executiva.grid(row=0, column=1)
botao_primeira_classe.grid(row=0, column=2)




#botao_enviar = tk.Button(text="enviar", command=enviar)
#botao_enviar.grid(row=1, column=0)
 
janela.mainloop()