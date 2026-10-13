from pathlib import Path
import pandas as pd

base_dir = Path(__file__).resolve().parent.parent / "data"
base_dir.mkdir(parents=True, exist_ok=True)

# 1. Ingestão: novos dados para o SQL Server
df_ingestao = pd.DataFrame({
    "id_venda": [101, 102, 103, 104, 105],
    "data_venda": ["2026-10-01", "2026-10-02", "2026-10-03", "2026-10-04", "2026-10-05"],
    "filial": ["Sul", "Sudeste", "Nordeste", "Sul", "Sudeste"],
    "produto": ["Notebook Pro", "Monitor 27", "Teclado Mecânico", "Cadeira Ergonômica", "Mouse Sem Fio"],
    "quantidade": [5, 12, 25, 8, 40],
    "valor_unitario": [4500.0, 1400.0, 350.0, 1200.0, 150.0],
})
df_ingestao["valor_total"] = df_ingestao["quantidade"] * df_ingestao["valor_unitario"]
df_ingestao.to_excel(base_dir / "01_ingestao_dados.xlsx", sheet_name="Carga_Novas_Vendas", index=False)

# 2. Extração: dados simulados extraídos do banco
df_extracao = pd.DataFrame({
    "mes_ano": ["2026-08", "2026-09", "2026-10"],
    "filial": ["Sul", "Sudeste", "Nordeste"],
    "faturamento_total": [185000.0, 210000.0, 245000.0],
    "ticket_medio": [1250.0, 1310.0, 1420.0],
    "qtd_pedidos": [148, 160, 172]
})
df_extracao.to_excel(base_dir / "02_extracao_dw.xlsx", sheet_name="Dados_Extraidos_SQL", index=False)

# 3. Relatório Executivo: dados e KPIs com fórmulas
with pd.ExcelWriter(base_dir / "03_relatorio_executivo.xlsx", engine="openpyxl") as writer:
    df_resumo = pd.DataFrame({
        "Indicador": ["Faturamento Realizado", "Meta do Mês", "Desvio (R$)", "Atingimento (%)"],
        "Valor": ["=SOMA(Dados_Extraidos_SQL!C2:C4)", 600000.0, "=B2-B3", "=B2/B3"]
    })
    df_extracao.to_excel(writer, sheet_name="Dados_Extraidos_SQL", index=False)
    df_resumo.to_excel(writer, sheet_name="Painel_KPI", index=False)

print("Planilhas de teste criadas em data/ com sucesso!")