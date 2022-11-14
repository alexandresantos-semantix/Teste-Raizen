#!/usr/bin/env python
# coding: utf-8

# In[530]:


# Import das bibliotecas nencessárias para o desenvolvimento

import glob
import sys
import win32com.client as win32
import pandas as pd
import numpy as np
from pathlib import Path
import re
import sys

win32c = win32.constants


# In[531]:


# Lendo e processando os conjuntos de dados
# crindo o excel object
excel_vendas_fuel = win32.gencache.EnsureDispatch('Excel.Application')
file = "C:/Users/W10/Documents/vendas/vendas-combustiveis-m3"
wb_data = excel_vendas_fuel.Workbooks.Open(file)

for i in range(49,50): 
        pvtTable       =wb_data.Sheets("plan1").Range("B"+str(i)).PivotTable
print(pvtTable)


# In[532]:


# limpando os filtros de dados da planilha dinamica
page_range_item = []
for i in pvtTable.PageRange:
    page_range_item.append(str(i))
    
print(page_range_item)
pvtTable.PivotFields("UN. DA FEDERAÇÃO").ClearAllFilters()
pvtTable.PivotFields("PRODUTO").ClearAllFilters()


# In[533]:


# filtrando de dados da planilha dinamica (UF/PRODUTO)
pvtTable.PivotFields("UN. DA FEDERAÇÃO").CurrentPage = "ACRE"
pvtTable.PivotFields("PRODUTO").CurrentPage = "ETANOL HIDRATADO (m3)"
page_range_item = []
for i in pvtTable.PageRange:
    page_range_item.append(str(i))
    
print(page_range_item)


# In[534]:


# Select the path of the folder with all the files
#Selecionado caminho da pasta com todos os arquivos caso haja mais de um
files = glob.glob("C:/Users/W10/Documents/vendas*.xls")


# In[535]:


# Redireciona os dados para o arquivo txt
orig_stdout = sys.stdout
bk = io.open("vendas_fuel.txt", mode="w", encoding="utf-8")
sys.stdout = bk


# In[536]:


#realiza a busca de todos os arquivos na pasta
for file in files:
    print(file.split('\\')[2])
    wb_data = excel_vendas_fuel.Workbooks.Open(file)
    


# In[543]:


nwSheet = Worksheets.Add 
nwSheet.Activate 
pvtTable = Worksheets("plan1").Range("C49").PivotTable 
rw = 0 
for pvtField in pvtTable.PivotFields 
    rw = rw + 1 
    nwSheet.Cells(rw, 1).Value = pvtField.Name 
next pvtField


# In[503]:


# busca e define a estrutura dos dados(a serem exportados para o arquivo vendas_fuel.txt)
for i in range(54,66): 
        mes       =wb_data.Worksheets("plan1").Range("B"+str(i))
        valor_2000=wb_data.Worksheets("plan1").Range("C"+str(i))
        valor_2001=wb_data.Worksheets("plan1").Range("D"+str(i))
        valor_2002=wb_data.Worksheets("plan1").Range("E"+str(i))
        valor_2003=wb_data.Worksheets("plan1").Range("F"+str(i))
        valor_2004=wb_data.Worksheets("plan1").Range("G"+str(i))
        valor_2005=wb_data.Worksheets("plan1").Range("H"+str(i))
        valor_2006=wb_data.Worksheets("plan1").Range("I"+str(i))
        valor_2007=wb_data.Worksheets("plan1").Range("J"+str(i))
        valor_2008=wb_data.Worksheets("plan1").Range("K"+str(i))
        valor_2009=wb_data.Worksheets("plan1").Range("L"+str(i))
        valor_2010=wb_data.Worksheets("plan1").Range("M"+str(i))
        valor_2011=wb_data.Worksheets("plan1").Range("N"+str(i))
        valor_2012=wb_data.Worksheets("plan1").Range("O"+str(i))
        valor_2013=wb_data.Worksheets("plan1").Range("P"+str(i))
        valor_2014=wb_data.Worksheets("plan1").Range("Q"+str(i))
        valor_2015=wb_data.Worksheets("plan1").Range("R"+str(i))
        valor_2016=wb_data.Worksheets("plan1").Range("S"+str(i))
        valor_2017=wb_data.Worksheets("plan1").Range("T"+str(i))
        valor_2018=wb_data.Worksheets("plan1").Range("U"+str(i))
        valor_2019=wb_data.Worksheets("plan1").Range("V"+str(i))
        valor_2020=wb_data.Worksheets("plan1").Range("W"+str(i))
        venda=wb_data.Worksheets("plan1").Range("B44")
        uf=wb_data.Worksheets("plan1").Range("C49")
        prod=wb_data.Worksheets("plan1").Range("C50")
        print("Mes;Valor;Ano;Venda;Produto;UF")
        print(mes,';',valor_2000,';2000;',venda,';',prod,';',uf)
        print(mes,';',valor_2001,';2001;',venda,';',prod,';',uf)
        print(mes,';',valor_2002,';2002;',venda,';',prod,';',uf)
        print(mes,';',valor_2003,';2003;',venda,';',prod,';',uf)
        print(mes,';',valor_2004,';2004;',venda,';',prod,';',uf)
        print(mes,';',valor_2005,';2005;',venda,';',prod,';',uf)
        print(mes,';',valor_2006,';2006;',venda,';',prod,';',uf)
        print(mes,';',valor_2007,';2007;',venda,';',prod,';',uf)
        print(mes,';',valor_2008,';2008;',venda,';',prod,';',uf)
        print(mes,';',valor_2009,';2009;',venda,';',prod,';',uf)
        print(mes,';',valor_2010,';2010;',venda,';',prod,';',uf)
        print(mes,';',valor_2011,';2011;',venda,';',prod,';',uf)
        print(mes,';',valor_2012,';2012;',venda,';',prod,';',uf)
        print(mes,';',valor_2013,';2013;',venda,';',prod,';',uf)
        print(mes,';',valor_2014,';2014;',venda,';',prod,';',uf)
        print(mes,';',valor_2015,';2015;',venda,';',prod,';',uf)
        print(mes,';',valor_2016,';2016;',venda,';',prod,';',uf)
        print(mes,';',valor_2017,';2017;',venda,';',prod,';',uf)
        print(mes,';',valor_2018,';2018;',venda,';',prod,';',uf)
        print(mes,';',valor_2019,';2019;',venda,';',prod,';',uf)
        print(mes,';',valor_2020,';2020;',venda,';',prod,';',uf)
        print()


# In[504]:


# fecha o arquivo fonte sem salvar
wb_data.Close(True)


# In[510]:


#importa o arquivo vendas_fuel.txt que foi gerado com os dados da planilha dinamica.
filename = 'vendas_fuel.txt'
array_venda_fuel = np.loadtxt(filename, delimiter=';', skiprows=0, dtype=str)
print(array_venda_fuel)


# In[511]:


#transforma o array em lista
list_venda_fuel = array_venda_fuel.tolist()

print(list_venda_fuel)


# In[512]:


#transforma a lista em DataFrame
df_vendas_fuel = pd.DataFrame(list_venda_fuel, columns=['Mes', 'Valor', 'Ano', 'Venda', 'Produto', 'UF'])


# In[513]:


display(df_vendas_fuel)


# In[ ]:




