#!/usr/bin/env python
# coding: utf-8

# In[4]:


#teste para ler as duas guias ao mesmo tempo com filtro numa delas:

import pandas as pd

estoque = pd.read_excel("estoque.xlsx")

item = 'ITEM 1'


filtro = estoque['NOME'] == item
itens_filtrados = estoque[filtro]


display(itens_filtrados)


usuarios = pd.read_excel("estoque.xlsx", sheet_name="USUARIOS")


display(usuarios)

