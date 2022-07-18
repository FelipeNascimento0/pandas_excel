import pandas as pd
import win32com.client as win32

sales_table = pd.read_excel('Vendas.xlsx')

invoicing = sales_table[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()

amount = sales_table[['ID Loja', 'Produto', 'Quantidade']].groupby('ID Loja').sum()

ticket_medium = (invoicing['Valor Final'] / amount['Quantidade']).to_frame()
ticket_medium = ticket_medium.rename(columns={0: 'Ticket Médio'})


pd.set_option('display.max_columns', None)


outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'recipient'
mail.Subject = 'Relatorio de Vendas'
mail.HTMLBody = f'''

<p>Hello sir</p>

<p>here is the sales report assigned to me</p>

<p>Invoicing:</p>
{invoicing.to_html(formatters={'Valor Final':'R${:,.2f}'.format})}

<p>Amount:</p>
{amount.to_html()}

<p>Ticket Medium:</p>
{ticket_medium.to_html(formatters={'Ticket Médio':'R${:,.2f}'.format})}


<p>If you have any doubts<br>
I will be at your disposal.</p>

'''

mail.Send()