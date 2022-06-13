# 1- importar bilbliotecas
import pandas as pd
import win32com.client as win32
# 2 - Trazer a base de dados para o Python e ver o que tem nela
tabela = pd.read_excel('Vendas.xlsx')
print(tabela)

# 4 - Pegar um panorama geral sobre sua base de dados

faturamento_total = tabela["Valor Final"].sum()

# 5 - Pegar o faturamento por loja
faturamento_por_loja = tabela[["ID Loja", "Valor Final"]].grupby("ID Loja").sum()

# 6 - Entrar no detalhe para enteder
faturamento_por_produto = tabela[["ID Loja", "Produto", "Valor_Final"]].grupby("[ID Loja","Produto]").sum()

# 7 - Criar a integração com o Outlook
outlook = win32.Dispatch('outlook.application')

# 8 - Criar um email
email = outlook.CreateItem(0)

# 9 - Configurar as informações do seu email
email.To = "exemplo.@gmail.com"
email.Subject = "resultados da analise de dados"
email.HTMLBody = f"""
<p>Olá seugue os resultados da analise do faturamento menssal</p>
<p>O faturamento total foi de R${faturamento_total}
<p>O faturamento por loja foi de R${faturamento_por_loja}
<p>O faturamento por produto foi de R${faturamento_por_produto}
"""
