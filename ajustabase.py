import pandas as pd
import os

from datetime import date
import datetime

now = datetime.datetime.now()
anomes = str(now.year) + '{0:0>2}'.format(str(now.month))

path = os.getcwd() + '/base_carteira.xlsx'

ano = {'FY10':'2017','FY11':'2018','FY12':'2019',
       '2019':'2019','2018':'2018','2019':'2019'}
	   
indicador = {'PAT_LIQUIDO':'Patrimônio Líquido',
			 'Sobra Líquida':'Sobra Líquida',
			 'Saldo Provisão':'Provisão',
			 'CA InadGlobal 90':'Over90',
			 'Cart Crdto Total': 'Carteira Total',
			 'Depósitos':'Depósitos',
			 'ativos_adm':'ativos_adm'}
	   
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

lista1,lista2 = lista(anomes)

cred = pd.read_excel(path,sheet_name='credito-consulta',skiprows=1)
cred = cred[cred['Unnamed: 0'].map(lista1['Nº']).isna()==False]
cred['Unnamed: 0'] = cred['Unnamed: 0'].map(lista1['Nº'])

cred['Saldo Inadimplência'] = cred[['Saldo Inadimplência','Valor Ano']].sum(axis=1).where(cred['Saldo Inadimplência'] == 0, cred['Saldo Inadimplência'])

cred = cred[['Unnamed: 0', 'Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3','Saldo Inadimplência']]
cred.rename(columns={'Unnamed: 0':'credis','Unnamed: 1':'ano','Unnamed: 2':'mes','Unnamed: 3':'indicador','Saldo Inadimplência':'valor'},inplace=True)

pl = pd.read_excel(path,sheet_name='pl - Relatório',skiprows=1,
                         converters={'Unnamed: 0':str})
pl = pl[pl['Unnamed: 0'].map(lista1['Nº']).isna()==False]
pl['Unnamed: 0'] = pl['Unnamed: 0'].map(lista1['Nº'])
pl.rename(columns={'Unnamed: 0':'credis','Unnamed: 1':'ano','Unnamed: 2':'mes','Unnamed: 3':'indicador','Saldo Inadimplência':'valor','RS_REAL_ATU':'valor'},inplace=True)


meta = pd.read_excel(path,sheet_name='Planilha5 - Relatório',skiprows=1,
                         converters={'Unnamed: 0':str})
meta = meta[meta['Unnamed: 0'].map(lista1['Nº']).isna()==False]
meta['Unnamed: 0'] = meta['Unnamed: 0'].map(lista1['Nº'])
meta.rename(columns={'Unnamed: 0':'credis','Unnamed: 1':'ano','Unnamed: 2':'mes','Unnamed: 3':'indicador','Saldo Inadimplência':'valor','Cenario':'valor'},inplace=True)

meta_aj = meta[(meta['indicador'] == 'Fundos') | (meta['indicador'] == 'Previdência Total')
        | (meta['indicador'] == 'Recursos Direcionados') | (meta['indicador'] == 'Depósitos Poupança')]

ativo = pd.read_excel(path,sheet_name='planning-planning - - Relatório',skiprows=2,
                         converters={'Unnamed: 0':str})
						 
ativo = ativo[ativo['Unnamed: 0'].map(lista1['Nº']).isna()==False]
ativo['Unnamed: 0'] = ativo['Unnamed: 0'].map(lista1['Nº'])
ativo['Unnamed: 1'] = ativo['Unnamed: 1'].map(ano)
ativo.rename(columns={'Unnamed: 0':'credis','Unnamed: 1':'ano','Unnamed: 2':'mes','Unnamed: 3':'indicador','Saldo Inadimplência':'valor','Realizado':'valor'},inplace=True)

ativo_adm = ativo.append(meta_aj)
mespassado = now + datetime.timedelta(days=-1)
ativo_adm = ativo_adm.groupby('credis').sum().reset_index()

ativo_adm['indicador'] = 'ativos_adm'
ativo_adm['ano'] = mespassado.year
ativo_adm['mes'] = mespassado.month -1
ativo_adm['mes'] = ativo_adm['mes'].map(mesn)

ativo_adm = ativo_adm[['credis','ano','mes','indicador','valor']]

print(ativo_adm.head())

tabela = ativo_adm.append(meta[meta['indicador']=='Depósitos'].append(pl.append(cred)))
tabela['indicador'] = tabela['indicador'].map(indicador)

print('\n')

tabela = tabela.pivot(index='credis',columns='indicador',values='valor')
print(tabela.head())

tabela['Sobra Líquida / PL'] = tabela['Sobra Líquida'] / tabela['Patrimônio Líquido']
tabela['Provisão/Carteira'] = tabela['Provisão'] / tabela['Carteira Total']
tabela['Over/Carteira'] = tabela['Over90'] / tabela['Carteira Total']

tabela['Sobra Líquida / PL'] = tabela['Sobra Líquida / PL'].apply(lambda x: '{0:.2%}'.format(x))
tabela['Provisão/Carteira'] = tabela['Provisão/Carteira'].apply(lambda x: '{0:.2%}'.format(x))
tabela['Over/Carteira'] = tabela['Over/Carteira'].apply(lambda x: '{0:.2%}'.format(x))

tabela=tabela.reset_index()
tabela['Nome']=tabela['credis'].map(lista2['Nome Fantasia'])

tabela = tabela[['Nome','Patrimônio Líquido','Sobra Líquida','Sobra Líquida / PL','Carteira Total','Provisão','Provisão/Carteira','Over90','Over/Carteira','Depósitos','ativos_adm']]
print(tabela.head())

print('Exportando bases para o arquivo')
writer = pd.ExcelWriter('base_carteira_credito_' + anomes + '.xlsx', engine='xlsxwriter')

cred.to_excel(writer, sheet_name='cred')
pl.to_excel(writer, sheet_name='pl')
meta.to_excel(writer, sheet_name='meta')
ativo.to_excel(writer, sheet_name='ativo')
ativo_adm.to_excel(writer, sheet_name='ativo_adm')
tabela.to_excel(writer, sheet_name='tabela')

writer.save()
print('Feito')