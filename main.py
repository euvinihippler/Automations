import pandas as pd

# Ler as tabelas
df_vendas = pd.read_excel("vendas_empresa_grande.xlsx")
df_estoque_lojas = pd.read_excel("estoque_lojas_empresa_grande.xlsx")

# Merge das duas tabelas com base nas colunas "Loja", "Marca" e "Produto" (adiciona a coluna de estoque)
df_Lojas = pd.merge(df_vendas, df_estoque_lojas, on=["Loja", "Marca", "Produto"])

# Média diária de vendas por loja
df_Lojas['Vendas Diárias'] = df_Lojas['Vendas_30_dias'] / 30  # Dado que o valor das planilhas é mensal

# Cobertura Atual de Estoque
df_Lojas['Cobertura Atual'] = df_Lojas['Estoque_Atual'] / df_Lojas['Vendas Diárias'] # Resultado em dias de estoque disponível seguindo a média de vendas por dia


