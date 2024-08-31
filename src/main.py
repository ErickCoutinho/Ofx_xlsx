import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
from processador_ofx import processar_arquivo_ofx
from extract_PagBank import extract_ofx_data
import ofxparse
import io


def processar_ofx():
    banco_selecionado = banco_var.get()
    caminho_arquivo = filedialog.askopenfilename(title="Selecione o arquivo OFX", filetypes=[("OFX files", "*.ofx")])

    if not caminho_arquivo:
        messagebox.showwarning("Aviso", "Nenhum arquivo foi selecionado.")
        return

    try:
        if banco_selecionado == "Bradesco" or banco_selecionado == "Inter":
            # Leitura específica para Bradesco e Inter
            with open(caminho_arquivo, 'r', encoding='utf-8', errors="ignore") as file:
                decoded_content = file.read()
                ofx = ofxparse.OfxParser.parse(io.StringIO(decoded_content))
            processar_arquivo_ofx(ofx)

        elif banco_selecionado == "PagBank":
            # Processamento específico para PagBank
            transactions_df, agency, account_number, account_type, initial_balance, available_balance, start_date, end_date = extract_ofx_data(
                caminho_arquivo)

            # Criando um objeto similar ao esperado por processar_arquivo_ofx
            class AccountStatement:
                def __init__(self, transactions, start_date, end_date):
                    self.start_date = start_date
                    self.end_date = end_date
                    self.transactions = transactions

            class Account:
                def __init__(self, branch_id, account_id, account_type, statement):
                    self.branch_id = branch_id
                    self.account_id = account_id
                    self.account_type = account_type
                    self.statement = statement

            # Convertendo transações para o formato esperado
            transactions = []
            for _, row in transactions_df.iterrows():
                transaction = ofxparse.Transaction()
                transaction.date = row['Data']
                transaction.amount = row['Valor']
                transaction.memo = row['Descrição']
                transaction.id = row['ID da Transação']
                transactions.append(transaction)

            # Criando o objeto de conta e extrato
            statement = AccountStatement(transactions, start_date, end_date)
            ofx_data = Account(agency, account_number, account_type, statement)

            # Chama o método para processar e gerar o Excel
            processar_arquivo_ofx(ofx_data)

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

# Opções de banco (Bradesco, Inter, PagBank)
radio_bradesco = ttk.Radiobutton(root, text="Bradesco", variable=banco_var, value="Bradesco")
radio_inter = ttk.Radiobutton(root, text="Inter", variable=banco_var, value="Inter")
radio_pagbank = ttk.Radiobutton(root, text="PagBank", variable=banco_var, value="PagBank")

radio_bradesco.pack(anchor=tk.W)
radio_inter.pack(anchor=tk.W)
radio_pagbank.pack(anchor=tk.W)

# Botão para processar o arquivo OFX
botao_processar = ttk.Button(root, text="Selecionar Arquivo e Processar", command=processar_ofx)
botao_processar.pack(pady=20)

root.mainloop()
