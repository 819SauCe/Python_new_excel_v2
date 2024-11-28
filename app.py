#Cria o executavel
#interface grafica
#gerencia a ordem dos scripts
#organiza os passos

import tkinter as tk
from tkinter import messagebox
from tkinter import PhotoImage
from threading import Thread
import subprocess
import os

def executar_scripts_em_ordem(path):
    """
    Função para executar os scripts em uma sequência predefinida.
    """
    # Lista dos scripts a serem executados em ordem
    scripts = ["StockandLot_1.py", "IntegrateStockData_2.py", "Verification_1.py", "estilizar_planilha.py"]

    try:
        for script in scripts:
            # Construir o caminho completo do script
            script_path = os.path.join(path, script)

            # Verificar se o arquivo existe
            if not os.path.exists(script_path):
                messagebox.showerror("Erro", f"Arquivo '{script}' não encontrado em:\n{path}")
                return False

            # Executar o script
            subprocess.run(["python", script_path], check=True)
        
        return True
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Erro", f"Erro ao executar um script: {e}")
        return False
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro inesperado: {e}")
        return False

def run_main():
    path = entry_path.get().strip()
    
    if not path or not os.path.exists(path):
        messagebox.showerror("Erro", "Por favor, insira um caminho válido.")
        return

    try:
        def task():
            sucesso = executar_scripts_em_ordem(path)
            if sucesso:
                messagebox.showinfo("Sucesso", "Todos os scripts foram executados com sucesso!")
            else:
                messagebox.showerror("Erro", "Ocorreu um erro ao executar os scripts. Verifique os logs para mais detalhes.")
        
        Thread(target=task).start()
        
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro inesperado ao tentar executar os scripts:\n{e}")

# Criação da janela principal
root = tk.Tk()
root.title("Executor de Scripts")
root.geometry("400x200")
root.configure(bg="#34495E")

# Carrega o ícone do aplicativo (usando .png)
try:
    icon = PhotoImage(file="3979302.png")  # Certifique-se de ter um arquivo '3979302.png' no mesmo diretório
    root.iconphoto(True, icon)
except Exception as e:
    print("Erro ao carregar o ícone:", e)

# Estilos modernizados
style_font = ("Helvetica Neue", 11)
color_bg = "#34495E"
color_fg = "#ECF0F1"
color_button = "#1ABC9C"
color_button_hover = "#16A085"

# Label para instrução
label_instruction = tk.Label(root, text="Informe o Caminho dos Scripts:", bg=color_bg, fg=color_fg, font=style_font)
label_instruction.pack(pady=15)

# Campo de entrada para o caminho
entry_path = tk.Entry(root, width=50, font=style_font, bg="#2C3E50", fg=color_fg, insertbackground=color_fg, relief="flat")
entry_path.pack(pady=5)
entry_path.insert(0, os.getcwd())  # Caminho padrão: diretório atual

# Função para efeito hover no botão
def on_enter(e):
    button_start['background'] = color_button_hover

def on_leave(e):
    button_start['background'] = color_button

# Botão para executar os scripts com estilo flat e hover
button_start = tk.Button(root, text="Executar", command=run_main, bg=color_button, fg="#FFFFFF", font=style_font, relief="flat", cursor="hand2")
button_start.pack(pady=30)
button_start.bind("<Enter>", on_enter)
button_start.bind("<Leave>", on_leave)

# Inicia o loop da interface
root.mainloop()
