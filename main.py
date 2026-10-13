import os
import sys
import time
import logging
from pathlib import Path
import pandas as pd
import pyodbc
from dotenv import load_dotenv
import win32com.client as win32

# 1. Configurações de Ambiente e Logs
load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S"
)
logger = logging.getLogger("PipelineAutomacao")

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"


# 2. Conexão com o SQL Server
def get_db_connection() -> pyodbc.Connection:
    server = os.getenv("DB_SERVER", r"localhost\BEIRAMAR")
    database = os.getenv("DB_DATABASE", "db_automacao")
    trusted = os.getenv("DB_TRUSTED_CONNECTION", "yes")

    if trusted.lower() in ("yes", "true", "1"):
        conn_str = (
            f"DRIVER={{ODBC Driver 17 for SQL Server}};"
            f"SERVER={server};"
            f"DATABASE={database};"
            f"Trusted_Connection=yes;"
            f"TrustServerCertificate=yes;"
        )
    else:
        user = os.getenv("DB_USER")
        pwd = os.getenv("DB_PASSWORD")
        conn_str = (
            f"DRIVER={{ODBC Driver 17 for SQL Server}};"
            f"SERVER={server};"
            f"DATABASE={database};"
            f"UID={user};"
            f"PWD={pwd};"
            f"TrustServerCertificate=yes;"
        )
    return pyodbc.connect(conn_str)


# 3. Etapa 1: Carga da Planilha 01 no SQL Server
def etapa_1_carga_sql():
    logger.info("=== [ETAPA 1/3] Carga de Dados no SQL Server ===")
    arquivo = DATA_DIR / "01_ingestao_dados.xlsx"
    
    if not arquivo.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {arquivo}")

    logger.info(f"Lendo registros de {arquivo.name}...")
    df = pd.read_excel(arquivo, sheet_name="Carga_Novas_Vendas")

    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("TRUNCATE TABLE dbo.stg_vendas;")
        
        insert_sql = """
            INSERT INTO dbo.stg_vendas 
            (id_venda, data_venda, filial, produto, quantidade, valor_unitario, valor_total)
            VALUES (?, ?, ?, ?, ?, ?, ?);
        """
        for _, row in df.iterrows():
            cursor.execute(
                insert_sql,
                int(row["id_venda"]),
                str(row["data_venda"]),
                str(row["filial"]),
                str(row["produto"]),
                int(row["quantidade"]),
                float(row["valor_unitario"]),
                float(row["valor_total"])
            )

        # Atualiza histórico evitando duplicatas
        merge_sql = """
            INSERT INTO dbo.fato_vendas_historico 
            (id_venda, data_venda, filial, produto, quantidade, valor_unitario, valor_total)
            SELECT s.id_venda, s.data_venda, s.filial, s.produto, s.quantidade, s.valor_unitario, s.valor_total
            FROM dbo.stg_vendas s
            WHERE NOT EXISTS (
                SELECT 1 FROM dbo.fato_vendas_historico f 
                WHERE f.id_venda = s.id_venda AND f.data_venda = s.data_venda
            );
        """
        cursor.execute(merge_sql)
        conn.commit()
        logger.info(f"[OK] {len(df)} registros processados e consolidados no banco.")
    finally:
        cursor.close()
        conn.close()


# 4. Etapas 2 e 3: Atualização Headless do Excel
def atualizar_planilha_headless(arquivo_path: Path, descricao_etapa: str):
    logger.info(f"=== {descricao_etapa} ===")
    if not arquivo_path.exists():
        raise FileNotFoundError(f"Arquivo não localizado: {arquivo_path}")

    inicio = time.time()
    excel_app = None
    workbook = None

    try:
        # Abre o Excel em background (sem tela gráfica)
        excel_app = win32.DispatchEx("Excel.Application")
        excel_app.Visible = False
        excel_app.DisplayAlerts = False
        excel_app.ScreenUpdating = False

        workbook = excel_app.Workbooks.Open(str(arquivo_path.resolve()))

        # Atualiza conexões de dados se existirem
        for conn in workbook.Connections:
            if conn.Type in (1, 2):
                try:
                    conn.OLEDBConnection.BackgroundQuery = False
                except Exception:
                    pass
        workbook.RefreshAll()
        excel_app.CalculateUntilAsyncQueriesDone()

        # Atualiza tabelas dinâmicas
        for sheet in workbook.Sheets:
            for pt in sheet.PivotTables():
                pt.Update()

        # Recalcula fórmulas
        excel_app.Calculate()

        workbook.Save()
        duracao = time.time() - inicio
        logger.info(f"[OK] {arquivo_path.name} atualizado e recalculado em {duracao:.2f}s.")
    finally:
        if workbook:
            workbook.Close(SaveChanges=False)
        if excel_app:
            excel_app.Quit()
        del workbook
        del excel_app


# 5. Fluxo Principal
def main():
    inicio_geral = time.time()
    logger.info("INICIANDO PIPELINE DE AUTOMAÇÃO DE RELATÓRIOS")

    try:
        # 1. Carrega dados no SQL Server
        etapa_1_carga_sql()

        # 2. Atualiza Planilha 02 (Extração do Banco)
        atualizar_planilha_headless(
            DATA_DIR / "02_extracao_dw.xlsx",
            "[ETAPA 2/3] Atualizando Planilha 02 (Extração do Banco)"
        )

        # 3. Atualiza Planilha 03 (Relatório Executivo e KPIs)
        atualizar_planilha_headless(
            DATA_DIR / "03_relatorio_executivo.xlsx",
            "[ETAPA 3/3] Atualizando Planilha 03 (Relatório Executivo e Fórmulas)"
        )

        tempo_total = time.time() - inicio_geral
        logger.info(f"PIPELINE CONCLUÍDO COM SUCESSO! Tempo total: {tempo_total:.2f}s")

    except Exception as e:
        logger.critical(f"Falha na execução do pipeline: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()