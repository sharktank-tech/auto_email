# 1- Trazer a base de dados para o Python e ver o que tem nela
import pandas as pd

tabela = pd.read_excel('Vendas.xlsx')
print(tabela)

# 2 - Pegar um panorama geral sobre sua base de dados

faturamento_total = tabela["Valor_Final"].sum()
print(faturamento_total)


# 3 - Pegar o faturamento por loja
faturamento_por_loja = tabela[["ID Loja", "Valor Final"]].grupby("ID Loja").sum()
print(faturamento_por_loja)

# 4 - Entrar no detalhe para enteder
faturamento_por_produto = tabela[["ID Loja", "Produto", "Valor Final"]].grupby("[ID Loja","Produto]").sum()
print(faturamento_por_loja)
