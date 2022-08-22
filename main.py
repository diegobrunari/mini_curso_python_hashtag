import pandas as pd

import win32com.client as win32

#importar base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

#visualizar base de dados
pd.set_option('display.max_columns', None)
print(tabela_vendas)

#faturamento por loja
faturamento = tabela_vendas [['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

#qntd de produtos vendidos por loja
quantidade = tabela_vendas [['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

#ticket médio por produto em cada loja (fat/qtd)
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print (ticket_medio)

#enviar e-mail com o relatorio
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'wcosta.ale@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o relatório de vendas por cada loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio dos produtos por loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>

<p>Att.,</p> 
<p>Diego.</p>
'''

mail.Send()

print('Email enviado')