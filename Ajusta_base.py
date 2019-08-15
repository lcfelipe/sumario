
# coding: utf-8

# In[ ]:


import pandas as pd
import os

from datetime import date
import datetime

now = datetime.datetime.now()
anomes = '201907' #str(now.year) + '{0:0>2}'.format(str(now.month))

path = os.getcwd() + '/dataraw/base_carteira.xlsx'

ano = {'FY10':'2017','FY11':'2018','FY12':'2019',
       '2019':'2019','2018':'2018','2019':'2019'}

mesn = {1:  'Janeiro',
        2:  'Fevereiro',
        3:  'Março',
        4:  'Abril',
        5:  'Maio',
        6:  'Junho',
        7:  'Julho',
        8:  'Agosto',
        9:  'Setembro',
        10:  'Outubro',
        11:  'Novembro',
        12:  'Dezembro'}

mesc = {'Janeiro':1,
        'Fevereiro':2,
        'Março':3,
        'Abril':4,
        'Maio':5,
        'Junho':6,
        'Julho':7,
        'Agosto':8,
        'Setembro':9,
        'Outubro':10,
        'Novembro':11,
        'Dezembro':12}

indicador = {'Sobra Líquida': 'Sobra Líquida'
             ,'PAT_LIQUIDO':'Patrimônio Líquido'
             ,'CA InadGlobal 90': 'Over90'
             ,'Cart Crdto Total': 'Carteira Total'
             ,'Saldo Provisão': 'Provisão'
             ,'Depósitos': 'Depósitos'
             ,'Ativos Adm.': 'Ativos Adm.'
             ,'Códigos de Produto':'Carteira Cred.'

}


# In[ ]:


def lista(anomes):
    coop = pd.read_excel('/Users/felipe_campos/Desktop/git/'+anomes+'_inventario.xlsx',
                         sheet_name='coop',
                         converters={'Credis':str,'Nº':str})
    coop['Credis'] = coop['Credis'].apply(lambda x: '{0:0>6}'.format(x))
    coop['Nº'] = coop['Nº'].apply(lambda x: '{0:0>4}'.format(x))
    #coop.set_index('Credis')

    lista_c = coop.set_index('Credis').to_dict()
    lista_n = coop.set_index('Nº').to_dict()

    return lista_c,lista_n

def lista_indi():
    indicador = pd.read_excel(os.getcwd() + '/dataraw/smart_base.xlsx',
                         sheet_name='INDI')

    lista_i = indicador.set_index('indi').to_dict()

    return lista_i

def ajusta_nps(arq):

    print('Importando arquivo... ')
    NPS = pd.read_excel(arq,sheet_name='NPS - Consulta - Relatório',converters={'Unnamed: 0':str,'Unnamed: 1':float},skiprows=1)
    ieic = pd.read_excel(arq,sheet_name='ie - ic - rsp',converters={'Unnamed: 0':str,'Unnamed: 1':float},skiprows=1)
    metas = pd.read_excel(arq,sheet_name='Metas - Relatório',converters={'Unnamed: 0':str,'Unnamed: 1':float},skiprows=1)
    
    print('Ajustando colunas...')
    NPS['ag'] = NPS['Unnamed: 0'].map(lista1['Nº'])
    NPS['mes'] = NPS['Unnamed: 2'].map(mesc)
    NPS['indicador'] = NPS['Unnamed: 3'].map(lista3['alias'])
    NPS['data'] = pd.to_datetime(NPS['Unnamed: 1']*10000+NPS['mes']*100+1,format='%Y%m%d')

    ieic['ag'] = ieic['Unnamed: 0'].map(lista1['Nº'])
    ieic['mes'] = ieic['Unnamed: 2'].map(mesc)
    ieic['indicador'] = ieic['Unnamed: 3'].map(lista3['alias'])
    ieic['data'] = pd.to_datetime(ieic['Unnamed: 1']*10000+ieic['mes']*100+1,format='%Y%m%d')

    metas['ag'] = metas['Unnamed: 0'].map(lista1['Nº'])
    metas['mes'] = metas['Unnamed: 2'].map(mesc)
    metas['indicador'] = metas['Unnamed: 3'].map(lista3['alias'])
    metas['data'] = pd.to_datetime(metas['Unnamed: 1']*10000+metas['mes']*100+1,format='%Y%m%d')

    metas = metas[(metas['ag'].isna() == False) & (metas['RS_REAL_ATU']>0)][['data','ag','indicador','RS_META_ATU','RS_REAL_ATU']].reset_index(drop=True)
    metas.rename(columns={'RS_META_ATU':'Planejado','RS_REAL_ATU':'Realizado'},inplace=True)    

    NPS = NPS[(NPS['ag'].isna() == False) & (NPS['Realizado']>0)][['data','ag','indicador','Planejado','Realizado']].reset_index(drop=True)
    ieic = ieic[(ieic['ag'].isna() == False) & (ieic['Realizado']>0)][['data','ag','indicador','Planejado','Realizado']].reset_index(drop=True)
    print('Feito...')
    
    return NPS.append(ieic.append(metas))


# In[ ]:


lista1,lista2 = lista(anomes)
lista3 = lista_indi()


# In[ ]:


NPS = pd.read_excel(caminho,sheet_name='NPS - Consulta - Relatório',converters={'Unnamed: 0':str,'Unnamed: 1':float},skiprows=1) 


# In[ ]:


caminho =os.getcwd() + '/dataraw/smart_base.xlsx'


# In[ ]:


ind_aj = ajusta_nps(caminho)


# In[ ]:


ind_aj[(ind_aj['ag']=='0101') & ((ind_aj['data']=='2019-02-01')|(ind_aj['data']=='2018-02-01'))]


# In[ ]:


NPS.head()


# # Crédito

# In[ ]:


cred = pd.read_excel(path,sheet_name='credito-consulta',skiprows=1)
cred = cred[cred['Unnamed: 0'].map(lista1['Nº']).isna()==False]
cred['Unnamed: 0'] = cred['Unnamed: 0'].map(lista1['Nº'])

cred['Saldo Inadimplência'] = cred[['Saldo Inadimplência','Valor Ano']].sum(axis=1).where(cred['Saldo Inadimplência'] == 0
                                                                                          , cred['Saldo Inadimplência'])

cred = cred[['Unnamed: 0', 'Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3','Saldo Inadimplência']]
cred.rename(columns={'Unnamed: 0':'credis','Unnamed: 1':'ano','Unnamed: 2':'mes','Unnamed: 3':'indicador'
                     ,'Saldo Inadimplência':'valor'},inplace=True)

pl = pd.read_excel(path,sheet_name='pl - Relatório',skiprows=1,
                         converters={'Unnamed: 0':str})
pl = pl[pl['Unnamed: 0'].map(lista1['Nº']).isna()==False]
pl['Unnamed: 0'] = pl['Unnamed: 0'].map(lista1['Nº'])
pl.rename(columns={'Unnamed: 0':'credis','Unnamed: 1':'ano','Unnamed: 2':'mes','Unnamed: 3':'indicador'
                   ,'Saldo Inadimplência':'valor','RS_REAL_ATU':'valor'},inplace=True)


meta = pd.read_excel(path,sheet_name='Planilha5 - Relatório',skiprows=1,
                         converters={'Unnamed: 0':str})
meta = meta[meta['Unnamed: 0'].map(lista1['Nº']).isna()==False]
meta['Unnamed: 0'] = meta['Unnamed: 0'].map(lista1['Nº'])
meta.rename(columns={'Unnamed: 0':'credis','Unnamed: 1':'ano','Unnamed: 2':'mes','Unnamed: 3':'indicador'
                     ,'Saldo Inadimplência':'valor','Cenario':'valor'},inplace=True)

meta_aj = meta[(meta['indicador'] == 'Fundos') 
               | (meta['indicador'] == 'Previdência Total')
               | (meta['indicador'] == 'Recursos Direcionados')
               | (meta['indicador'] == 'Depósitos Poupança')]

ativo = pd.read_excel(path,sheet_name='planning-planning - - Relatório',skiprows=2
                      ,converters={'Unnamed: 0':str})

ativo = ativo[ativo['Unnamed: 0'].map(lista1['Nº']).isna()==False]
ativo['Unnamed: 0'] = ativo['Unnamed: 0'].map(lista1['Nº'])
ativo['Unnamed: 1'] = ativo['Unnamed: 1'].map(ano)

ativo.rename(columns={'Unnamed: 0':'credis','Unnamed: 1':'ano','Unnamed: 2':'mes','Unnamed: 3':'indicador'
                      ,'Saldo Inadimplência':'valor','Realizado':'valor'}
             ,inplace=True)

ativo_adm = ativo.append(meta_aj)
mespassado = now + datetime.timedelta(days=-1)
ativo_adm = ativo_adm.groupby('credis').sum().reset_index()

ativo_adm['indicador'] = 'Ativos Adm.'
ativo_adm['ano'] = mespassado.year
ativo_adm['mes'] = mespassado.month -1
ativo_adm['mes'] = ativo_adm['mes'].map(mesn)

ativo_adm = ativo_adm[['credis','ano','mes','indicador','valor']]


# In[ ]:


cred_cart = pd.read_excel(path,sheet_name='cred',skiprows=1,converters={'Unnamed: 0':str})
cred_cart = cred_cart[cred_cart['Unnamed: 0'].map(lista1['Nº']).isna()==False]
cred_cart['Unnamed: 0'] = cred_cart['Unnamed: 0'].map(lista1['Nº'])
cred_cart.rename(columns={'Unnamed: 0':'credis','Unnamed: 1':'ano','Unnamed: 2':'mes'
                     ,'Unnamed: 3':'indicador','Saldo Atual':'valor'},inplace=True)


# In[ ]:


tabela = ativo_adm.append(meta[meta['indicador']=='Depósitos'].append(pl.append(cred.append(cred_cart))))

tabela['indicador'] = tabela['indicador'].map(indicador)

tabela = tabela.pivot(index='credis',columns='indicador',values='valor')

tabela['Sobra Líquida / PL'] = tabela['Sobra Líquida'] / tabela['Patrimônio Líquido']
tabela['Provisão/Carteira'] = tabela['Provisão'] / tabela['Carteira Total']
tabela['Over/Carteira'] = tabela['Over90'] / tabela['Carteira Total']

tabela['Sobra Líquida / PL'] = tabela['Sobra Líquida / PL'].apply(lambda x: '{0:.2%}'.format(x))
tabela['Provisão/Carteira'] = tabela['Provisão/Carteira'].apply(lambda x: '{0:.2%}'.format(x))
tabela['Over/Carteira'] = tabela['Over/Carteira'].apply(lambda x: '{0:.2%}'.format(x))

tabela=tabela.reset_index()
tabela['Nome'] = tabela.reset_index()['credis'].map(lista2['Nome Fantasia'])

tabela = tabela[['Nome','Patrimônio Líquido','Sobra Líquida','Sobra Líquida / PL','Carteira Cred.','Provisão'
                 ,'Provisão/Carteira','Over90','Over/Carteira','Depósitos','Ativos Adm.']]

tabela.to_excel(os.path.join(os.getcwd(),'dataset','estudo_carteira_'+anomes+'.xlsx'))

