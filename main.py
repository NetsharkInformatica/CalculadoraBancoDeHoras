
import tkinter as tk
from tkinter import messagebox
import sqlite3
from datetime import datetime
from openpyxl import Workbook
from tkinter import messagebox, filedialog


# Função para criar o banco de dados e a tabela
def criar_banco_de_dados():
    conn = sqlite3.connect('horas_extras.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS horas_extras
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, 
                  data TEXT, 
                  entrada TEXT, 
                  saida_turno_normal TEXT, 
                  saida_horas_extras TEXT, 
                  horas_excedentes REAL, 
                  descricao TEXT)''')
    conn.commit()
    conn.close()

# Função para calcular as horas excedentes
def calcular_horas_excedentes(entrada, saida_turno_normal, saida_horas_extras):
    formato = "%H:%M"
    entrada_time = datetime.strptime(entrada, formato)
    saida_turno_normal_time = datetime.strptime(saida_turno_normal, formato)
    saida_horas_extras_time = datetime.strptime(saida_horas_extras, formato)

    # Horas excedentes são a diferença entre a saída das horas extras e a saída do turno normal
    horas_excedentes = (saida_horas_extras_time - saida_turno_normal_time).total_seconds() / 3600
    return horas_excedentes

# Função para inserir horas extras
def inserir_horas():
    data = entry_data.get()
    entrada = entry_entrada.get()
    saida_turno_normal = entry_saida_turno_normal.get()
    saida_horas_extras = entry_saida_horas_extras.get()
    descricao = entry_descricao.get()

    if data and entrada and saida_turno_normal and saida_horas_extras and descricao:
        try:
            horas_excedentes = calcular_horas_excedentes(entrada, saida_turno_normal, saida_horas_extras)
            conn = sqlite3.connect('horas_extras.db')
            c = conn.cursor()
            c.execute("INSERT INTO horas_extras (data, entrada, saida_turno_normal, saida_horas_extras, horas_excedentes, descricao) VALUES (?, ?, ?, ?, ?, ?)",
                      (data, entrada, saida_turno_normal, saida_horas_extras, horas_excedentes, descricao))
            conn.commit()
            conn.close()
            messagebox.showinfo("Sucesso", "Horas extras inseridas com sucesso!")
            limpar_campos()
            listar_horas()
        except ValueError:
            messagebox.showwarning("Erro", "Formato de horário inválido! Use HH:MM.")
    else:
        messagebox.showwarning("Erro", "Preencha todos os campos!")

# Função para listar horas extras
def listar_horas():
    conn = sqlite3.connect('horas_extras.db')
    c = conn.cursor()
    c.execute("SELECT * FROM horas_extras")
    rows = c.fetchall()
    conn.close()

    listbox_horas.delete(0, tk.END)
    for row in rows:
        listbox_horas.insert(tk.END, f"ID: {row[0]}, Data: {row[1]}, Entrada: {row[2]}, Saída Turno Normal: {row[3]}, Saída Horas Extras: {row[4]}, Horas Excedentes: {row[5]:.2f}, Descrição: {row[6]}")

# Função para limpar os campos de entrada
def limpar_campos():
    entry_data.delete(0, tk.END)
    entry_entrada.delete(0, tk.END)
    entry_saida_turno_normal.delete(0, tk.END)
    entry_saida_horas_extras.delete(0, tk.END)
    entry_descricao.delete(0, tk.END)

# Função para editar horas extras
def editar_horas():
    selected = listbox_horas.curselection()
    if selected:
        id = listbox_horas.get(selected[0]).split(",")[0].split(":")[1].strip()
        data = entry_data.get()
        entrada = entry_entrada.get()
        saida_turno_normal = entry_saida_turno_normal.get()
        saida_horas_extras = entry_saida_horas_extras.get()
        descricao = entry_descricao.get()

        if data and entrada and saida_turno_normal and saida_horas_extras and descricao:
            try:
                horas_excedentes = calcular_horas_excedentes(entrada, saida_turno_normal, saida_horas_extras)
                conn = sqlite3.connect('horas_extras.db')
                c = conn.cursor()
                c.execute("UPDATE horas_extras SET data=?, entrada=?, saida_turno_normal=?, saida_horas_extras=?, horas_excedentes=?, descricao=? WHERE id=?",
                          (data, entrada, saida_turno_normal, saida_horas_extras, horas_excedentes, descricao, id))
                conn.commit()
                conn.close()
                messagebox.showinfo("Sucesso", "Horas extras atualizadas com sucesso!")
                limpar_campos()
                listar_horas()
            except ValueError:
                messagebox.showwarning("Erro", "Formato de horário inválido! Use HH:MM.")
        else:
            messagebox.showwarning("Erro", "Preencha todos os campos!")
    else:
        messagebox.showwarning("Erro", "Selecione um item para editar!")

# Função para deletar horas extras
def deletar_horas():
    selected = listbox_horas.curselection()
    if selected:
        id = listbox_horas.get(selected[0]).split(",")[0].split(":")[1].strip()
        conn = sqlite3.connect('horas_extras.db')
        c = conn.cursor()
        c.execute("DELETE FROM horas_extras WHERE id=?", (id,))
        conn.commit()
        conn.close()
        messagebox.showinfo("Sucesso", "Horas extras deletadas com sucesso!")
        listar_horas()
    else:
        messagebox.showwarning("Erro", "Selecione um item para deletar!")

# Função para exportar para Excel
def exportar_excel():
    pasta_destino = filedialog.askdirectory()
    if not pasta_destino:
        return
    
    conn = sqlite3.connect('horas_extras.db')
    c = conn.cursor()
    c.execute("SELECT * FROM horas_extras")
    rows = c.fetchall()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.append(["ID", "Data", "Entrada", "Saída Turno Normal", "Saída Horas Extras", "Horas Excedentes", "Descrição"])
    for row in rows:
        ws.append(row)

    caminho_arquivo = f"{pasta_destino}/horas_extras.xlsx"
    wb.save(caminho_arquivo)
    messagebox.showinfo("Sucesso", f"Dados exportados para {caminho_arquivo}")


""" def exportar_excel():
    conn = sqlite3.connect('horas_extras.db')
    c = conn.cursor()
    c.execute("SELECT * FROM horas_extras")
    rows = c.fetchall()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.append(["ID", "Data", "Entrada", "Saída Turno Normal", "Saída Horas Extras", "Horas Excedentes", "Descrição"])
    for row in rows:
        ws.append(row)

    wb.save("horas_extras.xlsx")
    messagebox.showinfo("Sucesso", "Dados exportados para horas_extras.xlsx") """

# Interface gráfica
root = tk.Tk()
root.title("Calculadora de Horas Extras")
root.resizable(False, False)  # Impede que a janela seja maximizada
root.iconbitmap("calculadora.ico")  # Define o ícone da janela

# Campos de entrada
tk.Label(root, text="Data (DD/MM/AAAA):").grid(row=0, column=0)
entry_data = tk.Entry(root)
entry_data.grid(row=0, column=1)

tk.Label(root, text="Horário de Entrada (HH:MM):").grid(row=1, column=0)
entry_entrada = tk.Entry(root)
entry_entrada.grid(row=1, column=1)

tk.Label(root, text="Horário de Saída do Turno Normal (HH:MM):").grid(row=2, column=0)
entry_saida_turno_normal = tk.Entry(root)
entry_saida_turno_normal.grid(row=2, column=1)

tk.Label(root, text="Horário de Saída das Horas Extras (HH:MM):").grid(row=3, column=0)
entry_saida_horas_extras = tk.Entry(root)
entry_saida_horas_extras.grid(row=3, column=1)

tk.Label(root, text="Descrição:").grid(row=4, column=0)
entry_descricao = tk.Entry(root)
entry_descricao.grid(row=4, column=1)

# Botões
tk.Button(root, text="Inserir", command=inserir_horas).grid(row=5, column=0)
tk.Button(root, text="Editar", command=editar_horas).grid(row=5, column=1)
tk.Button(root, text="Deletar", command=deletar_horas).grid(row=5, column=2)
tk.Button(root, text="Exportar para Excel", command=exportar_excel).grid(row=6, column=0, columnspan=3)

# Lista de horas extras
listbox_horas = tk.Listbox(root, width=100)
listbox_horas.grid(row=7, column=0, columnspan=3)

# Inicializar banco de dados e listar horas
criar_banco_de_dados()
listar_horas()

# Iniciar aplicação
root.mainloop()


