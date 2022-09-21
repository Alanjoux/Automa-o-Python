import pandas as pd
import win32com.client as win32

# Importa a base de dados
tabela = pd.read_excel('Vendas.xlsx') # Código para importa tabela

# Visualizar a base de dados
pd.set_option('display.max_columns', None) # Código para exibir número maxímo de linhas
print(tabela)

# Faturamento por loja
faturamento = tabela[['ID Loja','Valor Final']].groupby('ID Loja').sum() # Código para agrupar(.groupby) e somar(.sum) colunas de uma tabela.
print(faturamento)

# Tipo de produtos vendidos por loja
qtd_produto = tabela[['Produto', 'Quantidade']].groupby('Produto').sum()
print(qtd_produto)

# Quantidade de produtos vendidos por loja
quantidade = tabela[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

# Ticket médio por produto em cada loja
tick_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame() # (.to_frame) transforma dados em tabela.
tick_medio = tick_medio.rename(columns={0: 'Ticket Médio'})
print(tick_medio)

# Enviar um e-mail com relatótio 
outlook = win32.Dispatch('outlook.application') # Linha de código para envio de email automático.
mail = outlook.CreateItem(0)
mail.to = 'alanjoux30@gmail.com'
mail.subject = 'Relatório de vendas por loja'
mail.HTMLBody = f'''
<p>Prezados,</P>

<p>Segue o Relatório de Vesdas por cada Loja.</P>

<p>Faturamento</P>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade de Produtos Vendida</P>
{qtd_produto.to_html(formatters={'Quantidade': '{:,.0f}'.format})}

<p>Quantidade Vendida</P>
{quantidade.to_html(formatters={'Quantidade': '{:,.0f}'.format})}

<p>Ticket Médio dos Produtos em cada Loja</P>
{tick_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou a disposição.</P>

<p>Att.,</P>
<p>Alan.</P>
'''

mail.Send()

print('Email Enviado')