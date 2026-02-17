import pandas as pd
import numpy as np

# ============= FUNÇÕES AUXILIARES =============

def carregar_dados():
    """Carrega as três planilhas necessárias."""
    df_vendas = pd.read_excel("vendas_empresa_grande.xlsx")
    df_estoque_lojas = pd.read_excel("estoque_lojas_empresa_grande.xlsx")
    df_estoque_cd = pd.read_excel("estoque_cd_empresa_grande.xlsx")
    return df_vendas, df_estoque_lojas, df_estoque_cd


def preparar_lojas(df_vendas, df_estoque_lojas):
    """Merge das tabelas e cálculos iniciais de vendas."""
    df_lojas = pd.merge(
        df_vendas,
        df_estoque_lojas,
        on=["Loja", "Marca", "Produto"]
    )
    
    # Média diária de vendas (conversão de mensal para diária)
    df_lojas['Vendas Diárias'] = (df_lojas['Vendas_30_dias'] / 30).round(2)
    
    # Cobertura atual em dias de estoque
    df_lojas['Cobertura Atual'] = (
        df_lojas['Estoque_Atual'] / df_lojas['Vendas Diárias']
    ).round(2)
    
    return df_lojas


def calcular_necessidade_estoque(df_lojas):
    """Calcula a necessidade de estoque para 15 dias de cobertura."""
    estoque_ideal = df_lojas['Vendas Diárias'] * 15
    necessidade = estoque_ideal - df_lojas['Estoque_Atual']
    
    df_lojas['Necessidade de Estoque'] = np.where(
        necessidade < 0,
        0,
        necessidade.round(2)
    )
    
    return df_lojas


def alocar_estoque_cd(df_reposicao, df_estoque_cd):
    """Aloca estoque do CD para as lojas por ordem de prioridade."""
    # Criar dicionário do estoque do CD
    available = df_estoque_cd.set_index(['Marca', 'Produto'])['Estoque_CD'].to_dict()
    
    # Iterar e alocar
    for index, row in df_reposicao.iterrows():
        key = (row['Marca'], row['Produto'])
        necessidade = row['Necessidade de Estoque']
        estoque_cd = available.get(key, 0)
        
        # Definir quantidade entregue
        if estoque_cd <= 0:
            entregue = 0
        elif estoque_cd >= necessidade:
            entregue = necessidade
        else:
            entregue = estoque_cd
        
        # Registrar entrega e atualizar estoque do CD
        df_reposicao.at[index, 'Quantidade Entregue'] = entregue
        available[key] = estoque_cd - entregue
    
    return df_reposicao, available


def exibir_resultados(df_reposicao, available):
    """Exibe os resultados da alocação."""
    print("\n" + "="*80)
    print("RESULTADO DA ALOCAÇÃO DE ESTOQUE")
    print("="*80)
    
    print("\n10 Primeiras Lojas com Maior Prioridade:")
    print(df_reposicao[['Loja', 'Marca', 'Produto', 'Necessidade de Estoque', 
                        'Quantidade Entregue', 'Cobertura Atual']].head(10))
    
    print("\n\nEstoque Remanescente do CD:")
    df_estoque_final = pd.DataFrame(
        [(m, p, q) for (m, p), q in available.items()],
        columns=['Marca', 'Produto', 'Estoque_CD']
    )
    print(df_estoque_final[df_estoque_final['Estoque_CD'] > 0])
    
    print("\n" + "="*80)


# ============= FLUXO PRINCIPAL =============

if __name__ == "__main__":
    # Carregar dados
    df_vendas, df_estoque_lojas, df_estoque_cd = carregar_dados()
    
    # Preparar lojas
    df_lojas = preparar_lojas(df_vendas, df_estoque_lojas)
    
    # Calcular necessidade
    df_lojas = calcular_necessidade_estoque(df_lojas)
    
    # Ordenar por prioridade (menor cobertura = maior prioridade)
    df_reposicao = df_lojas.sort_values(by='Cobertura Atual', ascending=True)
    
    # Alocar estoque do CD
    df_reposicao, available = alocar_estoque_cd(df_reposicao, df_estoque_cd)
    
def gerar_relatorio(df_reposicao):
    """Gera relatório com informações de reposição e salva em Excel."""
    # Calcular pendência (necessidade - entregue)
    df_reposicao['Pendência'] = (
        df_reposicao['Necessidade de Estoque'] - 
        df_reposicao['Quantidade Entregue']
    ).clip(lower=0).round(2)
    
    # Selecionar colunas para o relatório
    relatorio = df_reposicao[[
        'Marca',
        'Produto',
        'Loja',
        'Estoque_Atual',
        'Necessidade de Estoque',
        'Cobertura Atual',
        'Quantidade Entregue',
        'Pendência'
    ]].copy()
    
    # Renomear colunas para melhor visualização
    relatorio.columns = [
        'Marca',
        'Produto',
        'Loja',
        'Estoque Loja',
        'Necessidade',
        'Prioridade (Cobertura)',
        'Quantidade Entregue',
        'Pendência'
    ]
    
    # Salvar em Excel
    relatorio.to_excel('relatorio.xlsx', index=False, sheet_name='Reposição')
    
    print("\n✓ Relatório salvo como 'relatorio.xlsx'")
    print(f"Total de linhas: {len(relatorio)}")
    print(f"\nPrimeiras 10 linhas:")
    print(relatorio.head(10))
    
    return relatorio


# ============= FLUXO PRINCIPAL =============

if __name__ == "__main__":
    # Carregar dados
    df_vendas, df_estoque_lojas, df_estoque_cd = carregar_dados()
    
    # Preparar lojas
    df_lojas = preparar_lojas(df_vendas, df_estoque_lojas)
    
    # Calcular necessidade
    df_lojas = calcular_necessidade_estoque(df_lojas)
    
    # Ordenar por prioridade (menor cobertura = maior prioridade)
    df_reposicao = df_lojas.sort_values(by='Cobertura Atual', ascending=True)
    
    # Alocar estoque do CD
    df_reposicao, available = alocar_estoque_cd(df_reposicao, df_estoque_cd)
    
    # Gerar relatório
    relatorio = gerar_relatorio(df_reposicao)
    
    # Exibir resultados
    exibir_resultados(df_reposicao, available)

fim
