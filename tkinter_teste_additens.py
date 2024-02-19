#!/usr/bin/env python
# coding: utf-8

# In[2]:


#teste com tkinter pra add items:

import openpyxl
import tkinter as tk
from tkinter import Entry, Label, Button, messagebox

def adicionar_item(planilha, item):
    """
    Adiciona um novo item à planilha.
    """
    planilha.append(item)

def obter_valores():
    """
    Obtém os valores dos campos de entrada e chama a função para adicionar o novo item.
    """
    novo_item = []
    for i, header in enumerate(headers):
        valor = entry_fields[i].get()
        novo_item.append(valor)

    adicionar_item(sheet, novo_item)
    messagebox.showinfo("Sucesso", "Item adicionado com sucesso!")

headers = ['ID', 'NOME','VALOR', 'MARCA', 'MODELO', 'ANO', 'FORNECEDOR', 'Nº FORNECEDOR', 'QUANTIDADE']
    

try:
    workbook = openpyxl.load_workbook('estoque.xlsx')
    sheet = workbook.active
except FileNotFoundError:
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(headers)

root = tk.Tk()
root.title("Adicionar Item ao Estoque")
root.geometry("600x400")


entry_fields = []
for i, header in enumerate(headers):
    tk.Label(root, text=header).grid(row=i, column=0, padx=5, pady=5)
    entry = tk.Entry(root)
    entry.grid(row=i, column=1, padx=5, pady=5)
    entry_fields.append(entry)


adicionar_button = tk.Button(root, text="Adicionar Item", command=obter_valores)
adicionar_button.grid(row=len(headers), column=0, columnspan=2, pady=10)


root.mainloop()


workbook.save('estoque.xlsx')

