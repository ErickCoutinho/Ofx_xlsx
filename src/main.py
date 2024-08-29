import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
from processador_ofx import processar_arquivo_ofx, remove_closing_tags
import ofxparse
import io


def processar_ofx():
    banco_selecionado = banco_var.get()
    caminho_arquivo = filedialog.askopenfilename(title="Selecione o arquivo OFX", filetypes=[("OFX files", "*.ofx")])

    if not caminho_arquivo:
        messagebox.showwarning("Aviso", "Nenhum arquivo foi selecionado.")
        return

    try:
        with open(caminho_arquivo, 'rb') as file:
            content = file.read()

            if banco_selecionado == "Bradesco":
                # Leitura específica para o arquivo Bradesco
                decoded_content = content.decode('utf-8', errors='ignore')
                modified_content = remove_closing_tags(decoded_content)
                cleaned_file = io.StringIO(modified_content)
                ofx = ofxparse.OfxParser.parse(cleaned_file)
            elif banco_selecionado == "Inter":
                # Leitura específica para o arquivo Inter
                ofx = ofxparse.OfxParser.parse(file)

        # Processa o arquivo OFX e gera o Excel
        processar_arquivo_ofx(ofx)

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao processar arquivo: {e}")
        return

    messagebox.showinfo("Sucesso", "Arquivo OFX processado e Excel criado com sucesso.")


# Interface gráfica em Tkinter
root = tk.Tk()
root.title("Processador de Arquivos OFX")

# Variável para armazenar a escolha do usuário
banco_var = tk.StringVar(value="Inter")

# Label para instruções
label_instrucoes = tk.Label(root, text="Selecione o banco correspondente ao arquivo OFX:")
label_instrucoes.pack(pady=10)

# Opções de banco (Bradesco ou Inter)
radio_bradesco = ttk.Radiobutton(root, text="Bradesco", variable=banco_var, value="Bradesco")
radio_inter = ttk.Radiobutton(root, text="Inter", variable=banco_var, value="Inter")

radio_bradesco.pack(anchor=tk.W)
radio_inter.pack(anchor=tk.W)

# Botão para processar o arquivo OFX
botao_processar = ttk.Button(root, text="Selecionar Arquivo e Processar", command=processar_ofx)
botao_processar.pack(pady=20)

root.mainloop()
