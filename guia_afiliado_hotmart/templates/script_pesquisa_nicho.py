"""
Script de Pesquisa de Nicho para Afiliados Hotmart
===================================================
Este script ajuda a validar nichos e subnichos usando dados do Google Trends
e análise de volume de busca.

Requisitos:
    pip install pytrends pandas matplotlib

Como usar:
    1. Edite a lista 'nichos' abaixo com os termos que deseja pesquisar
    2. Execute: python script_pesquisa_nicho.py
    3. O script gerará um relatório com gráficos comparativos
"""

from pytrends.request import TrendReq
import pandas as pd
import matplotlib.pyplot as plt
import time
import os

# ============================================================
# CONFIGURAÇÃO — Edite aqui os nichos que deseja pesquisar
# ============================================================

# Adicione até 5 termos por vez (limitação do Google Trends)
nichos = [
    "educação financeira",
    "inteligência artificial curso",
    "saúde mental",
    "trabalho remoto",
    "artesanato para vender"
]

# Período de análise (ex: 'today 12-m' para último ano, 'today 3-m' para 3 meses)
periodo = "today 12-m"

# País (BR = Brasil)
pais = "BR"

# ============================================================
# EXECUÇÃO DO SCRIPT
# ============================================================

def pesquisar_tendencias(termos, periodo, pais):
    """Pesquisa tendências no Google Trends para os termos informados."""
    print(f"\n{'='*60}")
    print(f"  PESQUISA DE NICHOS — Google Trends")
    print(f"{'='*60}")
    print(f"  Termos: {', '.join(termos)}")
    print(f"  Período: {periodo}")
    print(f"  País: {pais}")
    print(f"{'='*60}\n")

    try:
        pytrends = TrendReq(hl='pt-BR', tz=180)
        pytrends.build_payload(termos, cat=0, timeframe=periodo, geo=pais)

        # Interesse ao longo do tempo
        interesse = pytrends.interest_over_time()

        if interesse.empty:
            print("Nenhum dado encontrado. Verifique os termos e tente novamente.")
            return None

        # Remover coluna 'isPartial'
        if 'isPartial' in interesse.columns:
            interesse = interesse.drop(columns=['isPartial'])

        return interesse

    except Exception as e:
        print(f"Erro ao acessar Google Trends: {e}")
        print("Dica: Tente novamente em alguns minutos ou reduza o número de termos.")
        return None


def gerar_relatorio(dados, termos):
    """Gera relatório visual e textual dos resultados."""

    # 1. Gráfico de tendências
    plt.figure(figsize=(14, 7))
    for termo in termos:
        if termo in dados.columns:
            plt.plot(dados.index, dados[termo], label=termo, linewidth=2)

    plt.title("Comparativo de Interesse ao Longo do Tempo (Google Trends)", fontsize=14, fontweight='bold')
    plt.xlabel("Data")
    plt.ylabel("Interesse Relativo (0-100)")
    plt.legend(loc='upper left', fontsize=10)
    plt.grid(True, alpha=0.3)
    plt.tight_layout()

    arquivo_grafico = "relatorio_nichos_tendencias.png"
    plt.savefig(arquivo_grafico, dpi=150)
    print(f"Gráfico salvo em: {os.path.abspath(arquivo_grafico)}")
    plt.close()

    # 2. Tabela resumo
    print(f"\n{'='*60}")
    print(f"  RESUMO DOS NICHOS")
    print(f"{'='*60}\n")

    resumo = pd.DataFrame({
        'Nicho': termos,
        'Média de Interesse': [dados[t].mean() if t in dados.columns else 0 for t in termos],
        'Máximo': [dados[t].max() if t in dados.columns else 0 for t in termos],
        'Mínimo': [dados[t].min() if t in dados.columns else 0 for t in termos],
        'Tendência (Último vs Primeiro Mês)': [
            round(dados[t].tail(4).mean() - dados[t].head(4).mean(), 1)
            if t in dados.columns else 0
            for t in termos
        ]
    })

    resumo = resumo.sort_values('Média de Interesse', ascending=False)
    print(resumo.to_string(index=False))

    # 3. Recomendação
    print(f"\n{'='*60}")
    print(f"  RECOMENDAÇÃO")
    print(f"{'='*60}\n")

    melhor = resumo.iloc[0]
    crescendo = resumo[resumo['Tendência (Último vs Primeiro Mês)'] > 0]

    print(f"  Nicho com MAIOR interesse médio: {melhor['Nicho']} ({melhor['Média de Interesse']:.1f}/100)")

    if not crescendo.empty:
        print(f"\n  Nichos em CRESCIMENTO (tendência positiva):")
        for _, row in crescendo.iterrows():
            print(f"    - {row['Nicho']} (+{row['Tendência (Último vs Primeiro Mês)']:.1f} pontos)")
    else:
        print("  Nenhum nicho apresentou crescimento significativo no período.")

    print(f"\n{'='*60}")
    print(f"  DICA: Combine o nicho com maior interesse com uma tendência")
    print(f"  de crescimento para maximizar seu potencial de vendas.")
    print(f"{'='*60}\n")

    # Salvar resumo em CSV
    arquivo_csv = "relatorio_nichos_resumo.csv"
    resumo.to_csv(arquivo_csv, index=False, encoding='utf-8-sig')
    print(f"Resumo salvo em: {os.path.abspath(arquivo_csv)}")


if __name__ == "__main__":
    dados = pesquisar_tendencias(nichos, periodo, pais)
    if dados is not None:
        gerar_relatorio(dados, nichos)
    print("\nPesquisa concluída!")
