#!/usr/bin/env python
# coding: utf-8

# In[246]:


#Import das bibliotecas nencessárias para o desenvolvimento
import glob 
import sys 
import win32com.client as win32 
import pandas as pd 
import numpy as np 
from pathlib import Path 
import re 
import sys 
import datetime
import io


# In[247]:


#Lendo e processando os conjuntos de dados
#criando o excel object
excel_vendas_fuel = win32.gencache.EnsureDispatch('Excel.Application') 
file = "C:/Users/W10/Documents/vendas/vendas-combustiveis-m3"
wb_data = excel_vendas_fuel.Workbooks.Open(file)


# ##### 'Vendas, pelas distribuidoras¹, dos derivados combustíveis de petróleo por Unidade da Federação e produto - 2000-2020 (m3)'

# In[248]:


#celulas: ('B49','B50' segmentaçoa de dados 'UF','Produto')
for i in range(49,50): 
    pvtTable =wb_data.Sheets("plan1").Range("B"+str(i)).PivotTable 
    print(pvtTable)


# In[249]:


# limpando os filtros de dados da planilha dinamica
page_range_item = []
for i in pvtTable.PageRange:
    page_range_item.append(str(i))
    

pvtTable.PivotFields("UN. DA FEDERAÇÃO").ClearAllFilters()
pvtTable.PivotFields("PRODUTO").ClearAllFilters()
print(page_range_item)


# In[250]:


# encontrar todos os itens UF
#uf_items = []
#for item in pvtTable.PivotFields("UN. DA FEDERAÇÃO").PivotItems():
#    uf = str(item)
#    uf_items.append(uf)
#   uf_todos = ['ACRE','ALAGOAS','AMAPÁ','AMAZONAS','BAHIA','CEARÁ','DISTRITO FEDERAL','ESPÍRITO SANTO','GOIÁS',
#                   'MARANHÃO','MATO GROSSO','MATO GROSSO DO SUL','MINAS GERAIS','PARÁ','PARAÍBA','PARANÁ','PERNAMBUCO',
#                  'PIAUÍ','RIO DE JANEIRO','RIO GRANDE DO NORTE','RIO GRANDE DO SUL','RONDÔNIA','RORAIMA','SANTA CATARINA',
#                   'SÃO PAULO','SERGIPE','TOCANTINS']
#uf_vazio = [x for x in uf_items if x not in uf_todos]


# In[251]:


#filtrando de dados da planilha dinamica (UF/PRODUTO)
pvtTable.PivotFields("UN. DA FEDERAÇÃO").CurrentPage = "MINAS GERAIS"
pvtTable.PivotFields("PRODUTO").CurrentPage = "ETANOL HIDRATADO (m3)"
page_range_item = []
for i in pvtTable.PageRange:
    page_range_item.append(str(i))
    
print(page_range_item)


# In[252]:


#Selecionado caminho da pasta com todos os arquivos caso haja mais de um
files = glob.glob("C:/Users/W10/Documents/vendas*.xls")


# In[253]:


# Redireciona os dados para o arquivo txt
orig_stdout = sys.stdout
bk = io.open("vendas_petroleo_uf.txt", mode="w", encoding="utf-8")
sys.stdout = bk


# In[254]:


#realiza a busca de todos os arquivos na pasta
for file in files:
    print(file.split('\\')[2])
    wb_data = vendas_petroleo_uf.Workbooks.Open(file)
    


# In[255]:


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
        var_acumul=wb_data.Worksheets("plan1").Range("x"+str(i))
        venda=wb_data.Worksheets("plan1").Range("B44")
        uf=wb_data.Worksheets("plan1").Range("C49")
        prod=wb_data.Worksheets("plan1").Range("C50")
        geo_ref=wb_data.Worksheets("plan1").Range("B49")
        classe=wb_data.Worksheets("plan1").Range("B50")
        print("Mes;Valor;Ano;Venda;Produto;UF;Georef;Classificacao;Ingestion_Date")
        print(mes,';',valor_2000,';2000;',venda,';',prod,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2001,';2001;',venda,';',prod,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2002,';2002;',venda,';',prod,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2003,';2003;',venda,';',prod,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2004,';2004;',venda,';',prod,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2005,';2005;',venda,';',prod,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2006,';2006;',venda,';',prod,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2007,';2007;',venda,';',prod,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2008,';2008;',venda,';',prod,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2009,';2009;',venda,';',prod,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2010,';2010;',venda,';',prod,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2011,';2011;',venda,';',prod,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2012,';2012;',venda,';',prod,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2013,';2013;',venda,';',prod,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2014,';2014;',venda,';',prod,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2015,';2015;',venda,';',prod,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2016,';2016;',venda,';',prod,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2017,';2017;',venda,';',prod,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2018,';2018;',venda,';',prod,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2019,';2019;',venda,';',prod,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2020,';2020;',venda,';',prod,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',var_acumul,';Var_Acuml;',venda,';',prod,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print()


# In[ ]:


# fecha o arquivo fonte sem salvar
wb_data.Close(True)


# In[ ]:


#importa o arquivo vendas_fuel.txt que foi gerado com os dados da planilha dinamica.
filename = 'vendas_petroleo_uf.txt'
array_vendas_petroleo_uf = np.loadtxt(filename, delimiter=';', skiprows=0, dtype=str)


# In[ ]:


#transforma o array em lista
list_vendas_petroleo_uf = array_vendas_petroleo_uf.tolist()


# In[ ]:


#transforma a lista em DataFrame
df_vendas_petroleo_uf = pd.DataFrame(list_vendas_petroleo_uf, columns=['Mes', 'Valor', 'Ano', 'Venda','Produto','UF','Georef','Classificacao','Ingestion_Date'])
print(df_vendas_petroleo_uf)


# #### Vendas, pelas distribuidoras¹, dos derivados combustíveis de petróleo por Grande Região e produto - 2000-2020 (m3)

# In[141]:


#Lendo e processando os conjuntos de dados
#criando o excel object
excel_vendas_fuel = win32.gencache.EnsureDispatch('Excel.Application') 
file = "C:/Users/W10/Documents/vendas/vendas-combustiveis-m3"
wb_data = excel_vendas_fuel.Workbooks.Open(file)


# In[142]:


#celulas: ('B89','B90' segmentaçoa de dados 'Regiao','Produto')
for i in range(89,90): 
    pvtTable =wb_data.Sheets("plan1").Range("B"+str(i)).PivotTable 
    print(pvtTable)


# In[143]:


# limpando os filtros de dados da planilha dinamica
page_range_item = []
for i in pvtTable.PageRange:
    page_range_item.append(str(i))
    
pvtTable.PivotFields("GRANDE REGIÃO").ClearAllFilters()
pvtTable.PivotFields("PRODUTO").ClearAllFilters()
print(page_range_item)


# In[144]:


#filtrando de dados da planilha dinamica (REGIÃO/PRODUTO)
pvtTable.PivotFields("GRANDE REGIÃO").CurrentPage = "REGIÃO CENTRO-OESTE"
pvtTable.PivotFields("PRODUTO").CurrentPage = "ETANOL HIDRATADO (m3)"
page_range_item = []
for i in pvtTable.PageRange:
    page_range_item.append(str(i))
    
print(page_range_item)


# In[145]:


#Selecionado caminho da pasta com arquivos
files = glob.glob("C:/Users/W10/Documents/vendas*.xls")


# In[146]:


# Redireciona os dados para o arquivo txt
orig_stdout = sys.stdout
bk = io.open("vendas_petroleo_regiao.txt", mode="w", encoding="utf-8")
sys.stdout = bk


# In[147]:


#realiza a busca de todos os arquivos na pasta
for file in files:
    print(file.split('\\')[2])
    wb_data = vendas_petroleo_regiao.Workbooks.Open(file)
    


# In[148]:


# busca e define a estrutura dos dados(a serem exportados para o arquivo vendas_fuel.txt)
for i in range(94,106): 
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
        var_acumul=wb_data.Worksheets("plan1").Range("x"+str(i))
        venda=wb_data.Worksheets("plan1").Range("B84")
        regiao=wb_data.Worksheets("plan1").Range("C89")
        prod=wb_data.Worksheets("plan1").Range("C90")
        geo_ref=wb_data.Worksheets("plan1").Range("B89")
        classe=wb_data.Worksheets("plan1").Range("B90")
        print("Mes;Valor;Ano;Venda;Produto;Regiao;Georef;Classificacao;Ingestion_Date")
        print(mes,';',valor_2000,';2000;',venda,';',prod,';',regiao,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2001,';2001;',venda,';',prod,';',regiao,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2002,';2002;',venda,';',prod,';',regiao,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2003,';2003;',venda,';',prod,';',regiao,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2004,';2004;',venda,';',prod,';',regiao,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2005,';2005;',venda,';',prod,';',regiao,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2006,';2006;',venda,';',prod,';',regiao,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2007,';2007;',venda,';',prod,';',regiao,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2008,';2008;',venda,';',prod,';',regiao,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2009,';2009;',venda,';',prod,';',regiao,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2010,';2010;',venda,';',prod,';',regiao,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2011,';2011;',venda,';',prod,';',regiao,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2012,';2012;',venda,';',prod,';',regiao,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2013,';2013;',venda,';',prod,';',regiao,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2014,';2014;',venda,';',prod,';',regiao,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2015,';2015;',venda,';',prod,';',regiao,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2016,';2016;',venda,';',prod,';',regiao,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2017,';2017;',venda,';',prod,';',regiao,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2018,';2018;',venda,';',prod,';',regiao,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2019,';2019;',venda,';',prod,';',regiao,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2020,';2020;',venda,';',prod,';',regiao,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',var_acumul,';Var_Acuml;',venda,';',prod,';',regiao,';',geo_ref,';',classe,';',datetime.date.today())
        print()


# In[149]:


# fecha o arquivo fonte sem salvar
wb_data.Close(True)


# In[150]:


#importa o arquivo vendas_fuel.txt que foi gerado com os dados da planilha dinamica.
filename = 'vendas_petroleo_regiao.txt'
array_vendas_petroleo_regiao = np.loadtxt(filename, delimiter=';', skiprows=0, dtype=str)


# In[151]:


#transforma o array em lista
list_vendas_petroleo_regiao = array_vendas_petroleo_regiao.tolist()


# In[152]:


#transforma a lista em DataFrame
df_vendas_petroleo_regiao = pd.DataFrame(list_vendas_petroleo_regiao, columns=['Mes', 'Valor', 'Ano', 'Venda','Produto','Regiao','Georef','Classificacao','Ingestion_Date'])


# #### Vendas, pelas distribuidoras¹, de óleo diesel por tipo e Unidade da Federação - 2013-2020 (m3)

# In[153]:


#Lendo e processando os conjuntos de dados
#criando o excel object
excel_vendas_fuel = win32.gencache.EnsureDispatch('Excel.Application') 
file = "C:/Users/W10/Documents/vendas/vendas-combustiveis-m3"
wb_data = excel_vendas_fuel.Workbooks.Open(file)


# In[154]:


#celulas: ('B129','B130' segmentaçoa de dados 'UF','Produto')
for i in range(129,130): 
    pvtTable =wb_data.Sheets("plan1").Range("B"+str(i)).PivotTable 
    print(pvtTable)


# In[155]:


# limpando os filtros de dados da planilha dinamica
page_range_item = []
for i in pvtTable.PageRange:
    page_range_item.append(str(i))


pvtTable.PivotFields("UN. DA FEDERAÇÃO").ClearAllFilters()
pvtTable.PivotFields("PRODUTO").ClearAllFilters()
print(page_range_item)


# In[156]:


#filtrando de dados da planilha dinamica (REGIÃO/PRODUTO)
pvtTable.PivotFields("UN. DA FEDERAÇÃO").CurrentPage = "ACRE"
pvtTable.PivotFields("PRODUTO").CurrentPage = "ÓLEO DIESEL (OUTROS ) (m3)"
page_range_item = []
for i in pvtTable.PageRange:
    page_range_item.append(str(i))
    
print(page_range_item)


# In[157]:


#Selecionado caminho da pasta com arquivos
files = glob.glob("C:/Users/W10/Documents/vendas*.xls")


# In[158]:


# Redireciona os dados para o arquivo txt
orig_stdout = sys.stdout
bk = io.open("vendas_oleo_disel_tipo_uf.txt", mode="w", encoding="utf-8")
sys.stdout = bk


# In[159]:


#realiza a busca de todos os arquivos na pasta
for file in files:
    print(file.split('\\')[2])
    wb_data = vendas_oleo_disel_tipo_uf.Workbooks.Open(file)
    


# In[160]:


# busca e define a estrutura dos dados(a serem exportados para o arquivo vendas_fuel.txt)
for i in range(134,146): 
        mes       =wb_data.Worksheets("plan1").Range("B"+str(i))
        valor_2013=wb_data.Worksheets("plan1").Range("C"+str(i))
        valor_2014=wb_data.Worksheets("plan1").Range("D"+str(i))
        valor_2015=wb_data.Worksheets("plan1").Range("E"+str(i))
        valor_2016=wb_data.Worksheets("plan1").Range("F"+str(i))
        valor_2017=wb_data.Worksheets("plan1").Range("G"+str(i))
        valor_2018=wb_data.Worksheets("plan1").Range("H"+str(i))
        valor_2019=wb_data.Worksheets("plan1").Range("I"+str(i))
        valor_2020=wb_data.Worksheets("plan1").Range("J"+str(i))
        var_acumul=wb_data.Worksheets("plan1").Range("K"+str(i))
        venda=wb_data.Worksheets("plan1").Range("B124")
        uf=wb_data.Worksheets("plan1").Range("C129")
        prod=wb_data.Worksheets("plan1").Range("C130")
        geo_ref=wb_data.Worksheets("plan1").Range("B129")
        classe=wb_data.Worksheets("plan1").Range("B130")
        print("Mes;Valor;Ano;Venda;Produto;UF;Georef;Classificacao;Ingestion_Date")
        print(mes,';',valor_2013,';2013;',venda,';',prod,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2014,';2014;',venda,';',prod,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2015,';2015;',venda,';',prod,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2016,';2016;',venda,';',prod,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2017,';2017;',venda,';',prod,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2018,';2018;',venda,';',prod,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2019,';2019;',venda,';',prod,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2020,';2020;',venda,';',prod,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',var_acumul,';Var_Acuml;',venda,';',prod,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print()


# In[161]:


# fecha o arquivo fonte sem salvar
wb_data.Close(True)


# In[162]:


#importa o arquivo vendas_fuel.txt que foi gerado com os dados da planilha dinamica.
filename = 'vendas_oleo_disel_tipo_uf.txt'
array_vendas_oleo_disel_tipo_uf = np.loadtxt(filename, delimiter=';', skiprows=0, dtype=str)


# In[163]:


#transforma o array em lista
list_vendas_oleo_disel_tipo_uf = array_vendas_oleo_disel_tipo_uf.tolist()


# In[164]:


#transforma a lista em DataFrame
df_vendas_oleo_disel_tipo_uf = pd.DataFrame(list_vendas_oleo_disel_tipo_uf, columns=['Mes', 'Valor', 'Ano', 'Venda','Produto','UF','Georef','Classificacao','Ingestion_Date'])


# #### Vendas, pelas distribuidoras¹, de óleo diesel por tipo e Grande Região - 2013-2020 (m3)

# In[165]:


#Lendo e processando os conjuntos de dados
#criando o excel object
excel_vendas_fuel = win32.gencache.EnsureDispatch('Excel.Application') 
file = "C:/Users/W10/Documents/vendas/vendas-combustiveis-m3"
wb_data = excel_vendas_fuel.Workbooks.Open(file)


# In[166]:


#celulas: ('B167','B168' segmentaçoa de dados 'Regiao','Produto')
for i in range(167,168): 
    pvtTable =wb_data.Sheets("plan1").Range("B"+str(i)).PivotTable 
    print(pvtTable)


# In[167]:


# limpando os filtros de dados da planilha dinamica
page_range_item = []
for i in pvtTable.PageRange:
    page_range_item.append(str(i))
    

pvtTable.PivotFields("REGIÃO").ClearAllFilters()
pvtTable.PivotFields("PRODUTO").ClearAllFilters()
print(page_range_item)


# In[168]:


#filtrando de dados da planilha dinamica (REGIÃO/PRODUTO)
pvtTable.PivotFields("REGIÃO").CurrentPage = "REGIÃO CENTRO-OESTE"
pvtTable.PivotFields("PRODUTO").CurrentPage = "ÓLEO DIESEL (OUTROS ) (m3)"
page_range_item = []
for i in pvtTable.PageRange:
    page_range_item.append(str(i))
    
print(page_range_item)


# In[169]:


#Selecionado caminho da pasta com arquivos
files = glob.glob("C:/Users/W10/Documents/vendas*.xls")


# In[170]:


# Redireciona os dados para o arquivo txt
orig_stdout = sys.stdout
bk = io.open("vendas_oleo_disel_tipo_regiao.txt", mode="w", encoding="utf-8")
sys.stdout = bk


# In[171]:


#realiza a busca de todos os arquivos na pasta
for file in files:
    print(file.split('\\')[2])
    wb_data = vendas_oleo_disel_tipo_regiao.Workbooks.Open(file)


# In[172]:


# busca e define a estrutura dos dados(a serem exportados para o arquivo vendas_fuel.txt)
for i in range(172,184): 
        mes       =wb_data.Worksheets("plan1").Range("B"+str(i))
        valor_2013=wb_data.Worksheets("plan1").Range("C"+str(i))
        valor_2014=wb_data.Worksheets("plan1").Range("D"+str(i))
        valor_2015=wb_data.Worksheets("plan1").Range("E"+str(i))
        valor_2016=wb_data.Worksheets("plan1").Range("F"+str(i))
        valor_2017=wb_data.Worksheets("plan1").Range("G"+str(i))
        valor_2018=wb_data.Worksheets("plan1").Range("H"+str(i))
        valor_2019=wb_data.Worksheets("plan1").Range("I"+str(i))
        valor_2020=wb_data.Worksheets("plan1").Range("J"+str(i))
        var_acumul=wb_data.Worksheets("plan1").Range("K"+str(i))
        venda=wb_data.Worksheets("plan1").Range("B162")
        regiao=wb_data.Worksheets("plan1").Range("C167")
        prod=wb_data.Worksheets("plan1").Range("C18")
        geo_ref=wb_data.Worksheets("plan1").Range("B167")
        classe=wb_data.Worksheets("plan1").Range("B168")
        print("Mes;Valor;Ano;Venda;Produto;Regiao;Georef;Classificacao;Ingestion_Date")
        print(mes,';',valor_2013,';2013;',venda,';',prod,';',regiao,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2014,';2014;',venda,';',prod,';',regiao,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2015,';2015;',venda,';',prod,';',regiao,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2016,';2016;',venda,';',prod,';',regiao,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2017,';2017;',venda,';',prod,';',regiao,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2018,';2018;',venda,';',prod,';',regiao,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2019,';2019;',venda,';',prod,';',regiao,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2020,';2020;',venda,';',prod,';',regiao,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',var_acumul,';Var_Acuml;',venda,';',prod,';',regiao,';',geo_ref,';',classe,';',datetime.date.today())
        print()


# In[173]:


# fecha o arquivo fonte sem salvar
wb_data.Close(True)


# In[174]:


#importa o arquivo vendas_fuel.txt que foi gerado com os dados da planilha dinamica.
filename = 'vendas_oleo_disel_tipo_regiao.txt'
array_vendas_oleo_disel_tipo_regiao = np.loadtxt(filename, delimiter=';', skiprows=0, dtype=str)


# In[175]:


#transforma o array em lista
list_vendas_oleo_disel_tipo_regiao = array_vendas_oleo_disel_tipo_regiao.tolist()
#transforma a lista em DataFrame
df_vendas_oleo_disel_tipo_regiao = pd.DataFrame(list_vendas_oleo_disel_tipo_regiao, columns=['Mes', 'Valor', 'Ano', 'Venda','Produto','Regiao','Georef','Classificacao','Ingestion_Date'])


# #### Vendas, pelas distribuidoras¹, de GLP por Unidade da Federação e Vasilhame - 2010-2020 (m3)

# In[176]:


#Lendo e processando os conjuntos de dados
#criando o excel object
excel_vendas_fuel = win32.gencache.EnsureDispatch('Excel.Application') 
file = "C:/Users/W10/Documents/vendas/vendas-combustiveis-m3"
wb_data = excel_vendas_fuel.Workbooks.Open(file)


# In[177]:


#celulas: ('B206','B207' segmentaçoa de dados 'UF','Vasilhame')
for i in range(206,207): 
    pvtTable =wb_data.Sheets("plan1").Range("B"+str(i)).PivotTable 
    print(pvtTable)


# In[178]:


# limpando os filtros de dados da planilha dinamica
page_range_item = []
for i in pvtTable.PageRange:
    page_range_item.append(str(i))
    

pvtTable.PivotFields("UN. DA FEDERAÇÃO").ClearAllFilters()
pvtTable.PivotFields("VASILHAME").ClearAllFilters()
print(page_range_item)


# In[179]:


#filtrando de dados da planilha dinamica (REGIÃO/PRODUTO)
pvtTable.PivotFields("UN. DA FEDERAÇÃO").CurrentPage = "ACRE"
pvtTable.PivotFields("VASILHAME").CurrentPage = "GLP - Até P13 (m3)"
page_range_item = []
for i in pvtTable.PageRange:
    page_range_item.append(str(i))
    
print(page_range_item)


# In[180]:


#Selecionado caminho da pasta com arquivos
files = glob.glob("C:/Users/W10/Documents/vendas*.xls")


# In[181]:


# Redireciona os dados para o arquivo txt
orig_stdout = sys.stdout
bk = io.open("vendas_GLP_UF_vasilhame.txt", mode="w", encoding="utf-8")
sys.stdout = bk


# In[182]:


#realiza a busca de todos os arquivos na pasta
for file in files:
    print(file.split('\\')[2])
    wb_data = vendas_GLP_UF_vasilhame.Workbooks.Open(file)


# In[183]:


# busca e define a estrutura dos dados(a serem exportados para o arquivo vendas_fuel.txt)
for i in range(211,223): 
        mes       =wb_data.Worksheets("plan1").Range("B"+str(i))
        valor_2010=wb_data.Worksheets("plan1").Range("C"+str(i))
        valor_2011=wb_data.Worksheets("plan1").Range("D"+str(i))
        valor_2012=wb_data.Worksheets("plan1").Range("E"+str(i))
        valor_2013=wb_data.Worksheets("plan1").Range("F"+str(i))
        valor_2014=wb_data.Worksheets("plan1").Range("G"+str(i))
        valor_2015=wb_data.Worksheets("plan1").Range("H"+str(i))
        valor_2016=wb_data.Worksheets("plan1").Range("I"+str(i))
        valor_2017=wb_data.Worksheets("plan1").Range("J"+str(i))
        valor_2018=wb_data.Worksheets("plan1").Range("K"+str(i))
        valor_2019=wb_data.Worksheets("plan1").Range("L"+str(i))
        valor_2020=wb_data.Worksheets("plan1").Range("M"+str(i))
        var_acumul=wb_data.Worksheets("plan1").Range("N"+str(i))
        venda=wb_data.Worksheets("plan1").Range("B201")
        uf=wb_data.Worksheets("plan1").Range("C206")
        vas=wb_data.Worksheets("plan1").Range("C207")
        geo_ref=wb_data.Worksheets("plan1").Range("B206")
        classe=wb_data.Worksheets("plan1").Range("B207")
        print("Mes;Valor;Ano;Venda;Vasilhame;UF;Georef;Classificacao;Ingestion_Date")
        print(mes,';',valor_2010,';2010;',venda,';',vas,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2011,';2011;',venda,';',vas,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2012,';2012;',venda,';',vas,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2013,';2013;',venda,';',vas,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2014,';2014;',venda,';',vas,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2015,';2015;',venda,';',vas,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2016,';2016;',venda,';',vas,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2017,';2017;',venda,';',vas,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2018,';2018;',venda,';',vas,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2019,';2019;',venda,';',vas,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2020,';2020;',venda,';',vas,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',var_acumul,';Var_Acuml;',venda,';',vas,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print()


# In[184]:


#importa o arquivo vendas_fuel.txt que foi gerado com os dados da planilha dinamica.
filename = 'vendas_GLP_UF_vasilhame.txt'
array_vendas_GLP_uf_vasilhame = np.loadtxt(filename, delimiter=';', skiprows=0, dtype=str)


# In[187]:


#transforma o array em lista
list_vendas_GLP_vasilhame_uf = array_vendas_GLP_uf_vasilhame.tolist()
#transforma a lista em DataFrame
df_vendas_GLP_vasilhame_uf = pd.DataFrame(list_vendas_GLP_vasilhame_uf, columns=['Mes', 'Valor', 'Ano', 'Venda','Vasilhame','UF','Georef','Classificacao','Ingestion_Date'])


# #### Vendas, pelas distribuidoras¹, de GLP por Região e Vasilhame - 2010-2020 (m3)

# In[188]:


#Lendo e processando os conjuntos de dados
#criando o excel object
excel_vendas_fuel = win32.gencache.EnsureDispatch('Excel.Application') 
file = "C:/Users/W10/Documents/vendas/vendas-combustiveis-m3"
wb_data = excel_vendas_fuel.Workbooks.Open(file)


# In[189]:


#celulas: ('B242','B243' segmentaçoa de dados 'Regiao','Vasilhame')
for i in range(242,243): 
    pvtTable =wb_data.Sheets("plan1").Range("B"+str(i)).PivotTable 
    print(pvtTable)


# In[190]:


# limpando os filtros de dados da planilha dinamica
page_range_item = []
for i in pvtTable.PageRange:
    page_range_item.append(str(i))
    

pvtTable.PivotFields("REGIÃO").ClearAllFilters()
pvtTable.PivotFields("VASILHAME").ClearAllFilters()
print(page_range_item)


# In[191]:


#filtrando de dados da planilha dinamica (REGIÃO/VASILHAME)
pvtTable.PivotFields("REGIÃO").CurrentPage = "REGIÃO CENTRO-OESTE"
pvtTable.PivotFields("VASILHAME").CurrentPage = "GLP - Até P13 (m3)"
page_range_item = []
for i in pvtTable.PageRange:
    page_range_item.append(str(i))
    
print(page_range_item)


# In[192]:


#Selecionado caminho da pasta com arquivos
files = glob.glob("C:/Users/W10/Documents/vendas*.xls")


# In[193]:


# Redireciona os dados para o arquivo txt
orig_stdout = sys.stdout
bk = io.open("vendas_GLP_Regiao_vasilhame.txt", mode="w", encoding="utf-8")
sys.stdout = bk


# In[194]:


#realiza a busca de todos os arquivos na pasta
for file in files:
    print(file.split('\\')[2])
    wb_data = vendas_GLP_Regiao_vasilhame.Workbooks.Open(file)


# In[195]:


# busca e define a estrutura dos dados(a serem exportados para o arquivo vendas_fuel.txt)
for i in range(247,259): 
        mes       =wb_data.Worksheets("plan1").Range("B"+str(i))
        valor_2010=wb_data.Worksheets("plan1").Range("C"+str(i))
        valor_2011=wb_data.Worksheets("plan1").Range("D"+str(i))
        valor_2012=wb_data.Worksheets("plan1").Range("E"+str(i))
        valor_2013=wb_data.Worksheets("plan1").Range("F"+str(i))
        valor_2014=wb_data.Worksheets("plan1").Range("G"+str(i))
        valor_2015=wb_data.Worksheets("plan1").Range("H"+str(i))
        valor_2016=wb_data.Worksheets("plan1").Range("I"+str(i))
        valor_2017=wb_data.Worksheets("plan1").Range("J"+str(i))
        valor_2018=wb_data.Worksheets("plan1").Range("K"+str(i))
        valor_2019=wb_data.Worksheets("plan1").Range("L"+str(i))
        valor_2020=wb_data.Worksheets("plan1").Range("M"+str(i))
        var_acumul=wb_data.Worksheets("plan1").Range("N"+str(i))
        venda=wb_data.Worksheets("plan1").Range("B237")
        reg=wb_data.Worksheets("plan1").Range("C242")
        vas=wb_data.Worksheets("plan1").Range("C243")
        geo_ref=wb_data.Worksheets("plan1").Range("B242")
        classe=wb_data.Worksheets("plan1").Range("B243")
        print("Mes;Valor;Ano;Venda;Vasilhame;Regiao;Georef;Classificacao;Ingestion_Date")
        print(mes,';',valor_2010,';2010;',venda,';',vas,';',reg,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2011,';2011;',venda,';',vas,';',reg,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2012,';2012;',venda,';',vas,';',reg,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2013,';2013;',venda,';',vas,';',reg,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2014,';2014;',venda,';',vas,';',reg,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2015,';2015;',venda,';',vas,';',reg,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2016,';2016;',venda,';',vas,';',reg,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2017,';2017;',venda,';',vas,';',reg,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2018,';2018;',venda,';',vas,';',reg,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2019,';2019;',venda,';',vas,';',reg,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2020,';2020;',venda,';',vas,';',reg,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',var_acumul,';Var_Acuml;',venda,';',vas,';',reg,';',geo_ref,';',classe,';',datetime.date.today())
        print()


# In[196]:


#importa o arquivo vendas_fuel.txt que foi gerado com os dados da planilha dinamica.
filename = 'vendas_GLP_Regiao_vasilhame.txt'
array_vendas_GLP_Regiao_vasilhame = np.loadtxt(filename, delimiter=';', skiprows=0, dtype=str)


# In[198]:


#transforma o array em lista
list_vendas_GLP_vasilhame_regiao = array_vendas_GLP_Regiao_vasilhame.tolist()
#transforma a lista em DataFrame
df_vendas_GLP_vasilhame_regiao = pd.DataFrame(list_vendas_GLP_vasilhame_regiao, columns=['Mes', 'Valor', 'Ano', 'Venda','Vasilhame','Regiao','Georef','Classificacao','Ingestion_Date'])


# #### Vendas, pelas distribuidoras¹, de etanol hidratado por segmento e Unidade da Federação - 2012-2020 (m3)

# In[199]:


#Lendo e processando os conjuntos de dados
#criando o excel object
excel_vendas_fuel = win32.gencache.EnsureDispatch('Excel.Application') 
file = "C:/Users/W10/Documents/vendas/vendas-combustiveis-m3"
wb_data = excel_vendas_fuel.Workbooks.Open(file)


# In[200]:


#celulas: ('B278','B279' segmentaçoa de dados 'UF','Segmento')
for i in range(278,279): 
    pvtTable =wb_data.Sheets("plan1").Range("B"+str(i)).PivotTable 
    print(pvtTable)


# In[201]:


# limpando os filtros de dados da planilha dinamica
page_range_item = []
for i in pvtTable.PageRange:
    page_range_item.append(str(i))
    

pvtTable.PivotFields("UN. DA FEDERAÇÃO").ClearAllFilters()
pvtTable.PivotFields("SEGMENTO").ClearAllFilters()
print(page_range_item)


# In[202]:


#filtrando de dados da planilha dinamica (UF/SEGMENTO)
pvtTable.PivotFields("UN. DA FEDERAÇÃO").CurrentPage = "ACRE"
pvtTable.PivotFields("SEGMENTO").CurrentPage = "CONSUMIDOR FINAL"
page_range_item = []
for i in pvtTable.PageRange:
    page_range_item.append(str(i))
    
print(page_range_item)


# In[203]:


#Selecionado caminho da pasta com arquivos
files = glob.glob("C:/Users/W10/Documents/vendas*.xls")


# In[204]:


# Redireciona os dados para o arquivo txt
orig_stdout = sys.stdout
bk = io.open("vendas_etanol_hidratado_segmento_uf.txt", mode="w", encoding="utf-8")
sys.stdout = bk


# In[205]:


#realiza a busca de todos os arquivos na pasta
for file in files:
    print(file.split('\\')[2])
    wb_data = vendas_etanol_hidratado_segmento_uf.Workbooks.Open(file)


# In[206]:


# busca e define a estrutura dos dados(a serem exportados para o arquivo vendas_fuel.txt)
for i in range(247,259): 
        mes       =wb_data.Worksheets("plan1").Range("B"+str(i))
        valor_2012=wb_data.Worksheets("plan1").Range("C"+str(i))
        valor_2013=wb_data.Worksheets("plan1").Range("D"+str(i))
        valor_2014=wb_data.Worksheets("plan1").Range("E"+str(i))
        valor_2015=wb_data.Worksheets("plan1").Range("F"+str(i))
        valor_2016=wb_data.Worksheets("plan1").Range("G"+str(i))
        valor_2017=wb_data.Worksheets("plan1").Range("H"+str(i))
        valor_2018=wb_data.Worksheets("plan1").Range("I"+str(i))
        valor_2019=wb_data.Worksheets("plan1").Range("J"+str(i))
        valor_2020=wb_data.Worksheets("plan1").Range("K"+str(i))
        var_acumul=wb_data.Worksheets("plan1").Range("L"+str(i))
        venda=wb_data.Worksheets("plan1").Range("B273")
        uf=wb_data.Worksheets("plan1").Range("C278")
        seg=wb_data.Worksheets("plan1").Range("C279")
        geo_ref=wb_data.Worksheets("plan1").Range("B278")
        classe=wb_data.Worksheets("plan1").Range("B279")
        print("Mes;Valor;Ano;Venda;Segmento;UF;Georef;Classificacao;Ingestion_Date")
        print(mes,';',valor_2012,';2012;',venda,';',seg,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2013,';2013;',venda,';',seg,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2014,';2014;',venda,';',seg,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2015,';2015;',venda,';',seg,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2016,';2016;',venda,';',seg,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2017,';2017;',venda,';',seg,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2018,';2018;',venda,';',seg,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2019,';2019;',venda,';',seg,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2020,';2020;',venda,';',seg,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',var_acumul,';Var_Acuml;',venda,';',seg,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print()


# In[207]:


#importa o arquivo vendas_fuel.txt que foi gerado com os dados da planilha dinamica.
filename = 'vendas_etanol_hidratado_segmento_uf.txt'
array_vendas_etanol_hidratado_segmento_uf = np.loadtxt(filename, delimiter=';', skiprows=0, dtype=str)


# In[208]:


#transforma o array em lista
list_vendas_etanol_hidratado_segmento_uf = array_vendas_etanol_hidratado_segmento_uf.tolist()
#transforma a lista em DataFrame
df_vendas_etanol_hidratado_segmento_uf = pd.DataFrame(list_vendas_etanol_hidratado_segmento_uf, columns=['Mes', 'Valor', 'Ano', 'Venda','Segmento','UF','Georef','Classificacao','Ingestion_Date'])


# #### Vendas, pelas distribuidoras¹, de gasolina C por segmento e Unidade da Federação - 2012-2020 (m3)

# In[209]:


#Lendo e processando os conjuntos de dados
#criando o excel object
excel_vendas_fuel = win32.gencache.EnsureDispatch('Excel.Application') 
file = "C:/Users/W10/Documents/vendas/vendas-combustiveis-m3"
wb_data = excel_vendas_fuel.Workbooks.Open(file)


# In[210]:


#celulas: ('B313','B314' segmentaçoa de dados 'UF','Segmento')
for i in range(313,314): 
    pvtTable =wb_data.Sheets("plan1").Range("B"+str(i)).PivotTable 
    print(pvtTable)


# In[211]:


# limpando os filtros de dados da planilha dinamica
page_range_item = []
for i in pvtTable.PageRange:
    page_range_item.append(str(i))
    

pvtTable.PivotFields("UN. DA FEDERAÇÃO").ClearAllFilters()
pvtTable.PivotFields("SEGMENTO").ClearAllFilters()
print(page_range_item)


# In[212]:


#filtrando de dados da planilha dinamica (UF/SEGMENTO)
pvtTable.PivotFields("UN. DA FEDERAÇÃO").CurrentPage = "ACRE"
pvtTable.PivotFields("SEGMENTO").CurrentPage = "CONSUMIDOR FINAL"
page_range_item = []
for i in pvtTable.PageRange:
    page_range_item.append(str(i))
    
print(page_range_item)


# In[213]:


#Selecionado caminho da pasta com arquivos
files = glob.glob("C:/Users/W10/Documents/vendas*.xls")


# In[214]:


# Redireciona os dados para o arquivo txt
orig_stdout = sys.stdout
bk = io.open("vendas_gasolinaC_segmento_uf.txt", mode="w", encoding="utf-8")
sys.stdout = bk


# In[215]:


#realiza a busca de todos os arquivos na pasta
for file in files:
    print(file.split('\\')[2])
    wb_data = vendas_gasolinaC_segmento_uf.Workbooks.Open(file)


# In[216]:


# busca e define a estrutura dos dados(a serem exportados para o arquivo vendas_fuel.txt)
for i in range(318,330): 
        mes       =wb_data.Worksheets("plan1").Range("B"+str(i))
        valor_2012=wb_data.Worksheets("plan1").Range("C"+str(i))
        valor_2013=wb_data.Worksheets("plan1").Range("D"+str(i))
        valor_2014=wb_data.Worksheets("plan1").Range("E"+str(i))
        valor_2015=wb_data.Worksheets("plan1").Range("F"+str(i))
        valor_2016=wb_data.Worksheets("plan1").Range("G"+str(i))
        valor_2017=wb_data.Worksheets("plan1").Range("H"+str(i))
        valor_2018=wb_data.Worksheets("plan1").Range("I"+str(i))
        valor_2019=wb_data.Worksheets("plan1").Range("J"+str(i))
        valor_2020=wb_data.Worksheets("plan1").Range("K"+str(i))
        var_acumul=wb_data.Worksheets("plan1").Range("L"+str(i))
        venda=wb_data.Worksheets("plan1").Range("B308")
        uf=wb_data.Worksheets("plan1").Range("C2313")
        seg=wb_data.Worksheets("plan1").Range("C314")
        geo_ref=wb_data.Worksheets("plan1").Range("B313")
        classe=wb_data.Worksheets("plan1").Range("B314")
        print("Mes;Valor;Ano;Venda;Segmento;UF;Georef;Classificacao;Ingestion_Date")
        print(mes,';',valor_2012,';2012;',venda,';',seg,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2013,';2013;',venda,';',seg,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2014,';2014;',venda,';',seg,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2015,';2015;',venda,';',seg,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2016,';2016;',venda,';',seg,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2017,';2017;',venda,';',seg,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2018,';2018;',venda,';',seg,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2019,';2019;',venda,';',seg,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2020,';2020;',venda,';',seg,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',var_acumul,';Var_Acuml;',venda,';',seg,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print()


# In[217]:


#importa o arquivo vendas_fuel.txt que foi gerado com os dados da planilha dinamica.
filename = 'vendas_gasolinaC_segmento_uf.txt'
array_vendas_gasolinaC_segmento_uf = np.loadtxt(filename, delimiter=';', skiprows=0, dtype=str)


# In[218]:


#transforma o array em lista
list_vendas_gasolinaC_segmento_uf = array_vendas_gasolinaC_segmento_uf.tolist()
#transforma a lista em DataFrame
df_vendas_gasolinaC_segmento_uf = pd.DataFrame(list_vendas_gasolinaC_segmento_uf, columns=['Mes', 'Valor', 'Ano', 'Venda','Segmento','UF','Georef','Classificacao','Ingestion_Date'])


# #### Vendas, pelas distribuidoras¹, dos derivados combustíveis de petróleo por Grande Região e produto - 2000-2020 (m3)

# In[219]:


#Lendo e processando os conjuntos de dados
#criando o excel object
excel_vendas_fuel = win32.gencache.EnsureDispatch('Excel.Application') 
file = "C:/Users/W10/Documents/vendas/vendas-combustiveis-m3"
wb_data = excel_vendas_fuel.Workbooks.Open(file)


# In[221]:


#celulas: ('B3148,'B349' segmentaçoa de dados 'UF','Segmento')
for i in range(348,349): 
    pvtTable =wb_data.Sheets("plan1").Range("B"+str(i)).PivotTable 
    print(pvtTable)


# In[222]:


# limpando os filtros de dados da planilha dinamica
page_range_item = []
for i in pvtTable.PageRange:
    page_range_item.append(str(i))
    

pvtTable.PivotFields("UN. DA FEDERAÇÃO").ClearAllFilters()
pvtTable.PivotFields("SEGMENTO").ClearAllFilters()
print(page_range_item)


# In[223]:


#filtrando de dados da planilha dinamica (UF/SEGMENTO)
pvtTable.PivotFields("UN. DA FEDERAÇÃO").CurrentPage = "ACRE"
pvtTable.PivotFields("SEGMENTO").CurrentPage = "CONSUMIDOR FINAL"
page_range_item = []
for i in pvtTable.PageRange:
    page_range_item.append(str(i))
    
print(page_range_item)


# In[224]:


#Selecionado caminho da pasta com arquivos
files = glob.glob("C:/Users/W10/Documents/vendas*.xls")


# In[225]:


# Redireciona os dados para o arquivo txt
orig_stdout = sys.stdout
bk = io.open("vendas_derivados_combustíveis_petróleo_Regiao_produto.txt", mode="w", encoding="utf-8")
sys.stdout = bk


# In[226]:


#realiza a busca de todos os arquivos na pasta
for file in files:
    print(file.split('\\')[2])
    wb_data = vendas_derivados_combustíveis_petróleo_Regiao_produto.Workbooks.Open(file)


# In[227]:


# busca e define a estrutura dos dados(a serem exportados para o arquivo vendas_fuel.txt)
for i in range(353,365): 
        mes       =wb_data.Worksheets("plan1").Range("B"+str(i))
        valor_2012=wb_data.Worksheets("plan1").Range("C"+str(i))
        valor_2013=wb_data.Worksheets("plan1").Range("D"+str(i))
        valor_2014=wb_data.Worksheets("plan1").Range("E"+str(i))
        valor_2015=wb_data.Worksheets("plan1").Range("F"+str(i))
        valor_2016=wb_data.Worksheets("plan1").Range("G"+str(i))
        valor_2017=wb_data.Worksheets("plan1").Range("H"+str(i))
        valor_2018=wb_data.Worksheets("plan1").Range("I"+str(i))
        valor_2019=wb_data.Worksheets("plan1").Range("J"+str(i))
        valor_2020=wb_data.Worksheets("plan1").Range("K"+str(i))
        var_acumul=wb_data.Worksheets("plan1").Range("L"+str(i))
        venda=wb_data.Worksheets("plan1").Range("B343")
        uf=wb_data.Worksheets("plan1").Range("C348")
        seg=wb_data.Worksheets("plan1").Range("C349")
        geo_ref=wb_data.Worksheets("plan1").Range("B348")
        classe=wb_data.Worksheets("plan1").Range("B349")
        print("Mes;Valor;Ano;Venda;Segmento;UF;Georef;Classificacao;Ingestion_Date")
        print(mes,';',valor_2012,';2012;',venda,';',seg,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2013,';2013;',venda,';',seg,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2014,';2014;',venda,';',seg,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2015,';2015;',venda,';',seg,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2016,';2016;',venda,';',seg,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2017,';2017;',venda,';',seg,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2018,';2018;',venda,';',seg,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2019,';2019;',venda,';',seg,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',valor_2020,';2020;',venda,';',seg,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print(mes,';',var_acumul,';Var_Acuml;',venda,';',seg,';',uf,';',geo_ref,';',classe,';',datetime.date.today())
        print()


# In[228]:


#importa o arquivo vendas_fuel.txt que foi gerado com os dados da planilha dinamica.
filename = 'vendas_derivados_combustíveis_petróleo_Regiao_produto.txt'
array_vendas_derivados_combustíveis_petróleo_produto_regiao = np.loadtxt(filename, delimiter=';', skiprows=0, dtype=str)


# In[229]:


#transforma o array em lista
list_vendas_derivados_combustíveis_petróleo_produto_regiao = array_vendas_derivados_combustíveis_petróleo_produto_regiao.tolist()
#transforma a lista em DataFrame
df_vendas_derivados_combustíveis_petróleo_produto_regiao = pd.DataFrame(list_vendas_derivados_combustíveis_petróleo_produto_regiao, columns=['Mes', 'Valor', 'Ano', 'Venda','Segmento','UF','Georef','Classificacao','Ingestion_Date'])
print(df_vendas_derivados_combustíveis_petróleo_produto_regiao)


# In[238]:


#union all dos df de vendas por UF 
df_vendas_fuel_raizen_uf= pd.concat([df_vendas_petroleo_uf, 
                         df_vendas_oleo_disel_tipo_uf,
                         df_vendas_GLP_vasilhame_uf,
                         df_vendas_etanol_hidratado_segmento_uf,
                         df_vendas_gasolinaC_segmento_uf])
df_vendas_fuel_raizen_uf


# In[239]:


#union all dos df de vendas por REGIAO
df_vendas_fuel_raizen_regiao= pd.concat([df_vendas_petroleo_regiao,
                          df_vendas_oleo_disel_tipo_regiao,
                          df_vendas_derivados_combustíveis_petróleo_produto_regiao,
                          df_vendas_GLP_vasilhame_regiao])
df_vendas_fuel_raizen_regiao


# In[ ]:




