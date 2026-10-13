-- 1. Criação do banco de dados (caso não exista)
IF NOT EXISTS (SELECT * FROM sys.databases WHERE name = 'db_automacao')
BEGIN
    CREATE DATABASE db_automacao;
END
GO

USE db_automacao;
GO

-- 2. Tabela de Staging (recebe a carga bruta da Planilha 01)
IF OBJECT_ID('dbo.stg_vendas', 'U') IS NOT NULL
    DROP TABLE dbo.stg_vendas;
GO

CREATE TABLE dbo.stg_vendas (
    id_venda        INT NOT NULL,
    data_venda      DATE NOT NULL,
    filial          VARCHAR(50) NOT NULL,
    produto         VARCHAR(100) NOT NULL,
    quantidade      INT NOT NULL,
    valor_unitario  DECIMAL(10, 2) NOT NULL,
    valor_total     DECIMAL(12, 2) NOT NULL,
    data_carga      DATETIME DEFAULT GETDATE()
);
GO

-- 3. Tabela Histórica / DW (acumula os dados para as Planilhas 02 e 03)
IF OBJECT_ID('dbo.fato_vendas_historico', 'U') IS NOT NULL
    DROP TABLE dbo.fato_vendas_historico;
GO

CREATE TABLE dbo.fato_vendas_historico (
    sk_venda        INT IDENTITY(1,1) PRIMARY KEY,
    id_venda        INT NOT NULL,
    data_venda      DATE NOT NULL,
    filial          VARCHAR(50) NOT NULL,
    produto         VARCHAR(100) NOT NULL,
    quantidade      INT NOT NULL,
    valor_unitario  DECIMAL(10, 2) NOT NULL,
    valor_total     DECIMAL(12, 2) NOT NULL,
    data_insercao   DATETIME DEFAULT GETDATE()
);
GO

-- 4. Inserção de uma base histórica inicial de exemplo (meses anteriores)
INSERT INTO dbo.fato_vendas_historico (id_venda, data_venda, filial, produto, quantidade, valor_unitario, valor_total)
VALUES
    (1, '2026-08-10', 'Sul',      'Notebook Pro',       10, 4500.00,  45000.00),
    (2, '2026-08-15', 'Sudeste',  'Monitor 27',         20, 1400.00,  28000.00),
    (3, '2026-08-20', 'Nordeste', 'Teclado Mecânico',   50,  350.00,  17500.00),
    (4, '2026-09-05', 'Sul',      'Cadeira Ergonômica', 15, 1200.00,  18000.00),
    (5, '2026-09-12', 'Sudeste',  'Notebook Pro',       25, 4500.00, 112500.00),
    (6, '2026-09-25', 'Nordeste', 'Mouse Sem Fio',     100,  150.00,  15000.00);
GO

-- 5. View Analítica consumida pelas planilhas de BI e relatórios
CREATE OR ALTER VIEW dbo.vw_resumo_mensal AS
SELECT 
    FORMAT(data_venda, 'yyyy-MM') AS mes_ano,
    filial,
    SUM(valor_total)              AS faturamento_total,
    AVG(valor_total)              AS ticket_medio,
    COUNT(id_venda)               AS qtd_pedidos
FROM dbo.fato_vendas_historico
GROUP BY 
    FORMAT(data_venda, 'yyyy-MM'), 
    filial;
GO