# ***************************************************************************
# Lógica do problema proposto:
# Importar a base dados
# Visualizar/analisar base de dados
# Faturamento por loja
# Quantidade de produto vendidos por loja
# Ticket medio por produto em cada loja media (faturamento/qtd_produto)
# Enviar um email com o relatorio
#********************************************************************#

# Importar a base dados
# pip install pandas  - instalar biblioteca de bd
# pip install openpyxl - intalar bib. para ler arquivo xlsx
from pickle import NONE
import pandas as pd
tabelaVendas = pd.read_excel('Vendas.xlsx')

# Visualizar/analisar base de dados
pd.set_option('display.max_columns', None) # define o maximo de coluna(todas) do bd
print(tabelaVendas)

print (' - ' * 50)
# Faturamento por loja
#Filtra as colunas desejadas, agrupa por loja e soma o valor final de cada loja
faturamento = tabelaVendas[['ID Loja','Valor Final']].groupby('ID Loja').sum()
print(faturamento)

print (' - ' * 50)
# Quantidade de produto vendidos por loja
qtdProdVendido = tabelaVendas[['ID Loja','Quantidade']].groupby('ID Loja').sum()
print(qtdProdVendido)


print (' - ' * 50)
# Ticket medio por produto em cada loja media (faturamento/qtdProduto)
ticketMedio = (faturamento['Valor Final'] / qtdProdVendido['Quantidade']).to_frame()
print(ticketMedio)

# Enviar um email com o relatorio
import smtplib
import email.message

# colocar f anes das aspas para formatar
corpo_email = f"""
<p>Prezado,</p>
<p>Segue o Relatório de Vendas por cada Loja.</p>
<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}
<p>Quantidade Vendida:</p>  
{qtdProdVendido.to_html()}
<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticketMedio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}
<p>Qualquer dúvida estou à disposição.</p>
<p>Att.,</p>
<p>Ivan</p>
"""

msg = email.message.Message()
msg['Subject'] = 'Relatório de Vendas por Loja'
msg['From'] = 'largato18@gmail.com'
msg['To'] = 'ivanleoni18@hotmail.com'
password = 'XXXXXXX'
msg.add_header('Content-Type','text/html')
msg.set_payload(corpo_email )

s = smtplib.SMTP('smtp.gmail.com: 587')
s.starttls()
# Login Credentials for sending the mail
s.login(msg['From'], password)
s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))

print (' - ' * 50)
print('Email enviado')