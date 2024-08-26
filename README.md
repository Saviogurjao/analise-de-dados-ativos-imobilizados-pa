# Projeto de Análise e Automação para Benefícios Fiscais de Ativos Imobilizados no Estado do Pará

Este projeto automatiza a análise e geração de relatórios sobre benefícios fiscais concedidos para ativos imobilizados (máquinas e equipamentos) no Estado do Pará, utilizando Excel, VBA, Power BI e SQL. Também inclui a funcionalidade de anexar documentação relevante e verificar a validade de certidões.

## Estrutura do Projeto

### 1. Banco de Dados (SQL)

- **Nome do Banco de Dados**: `BeneficiosFiscais`
- **Tabelas Principais**:
  - `AtivosImobilizados`: Armazena informações sobre os ativos imobilizados.
    - `ID`: Identificador único.
    - `NomeAtivo`: Nome do ativo.
    - `NCM`: Código NCM (Nomenclatura Comum do Mercosul).
    - `DataAquisicao`: Data de aquisição do ativo.
    - `Valor`: Valor do ativo.
    - `DataInicioBeneficio`: Data de início do benefício fiscal.
    - `DataFimBeneficio`: Data de término do benefício fiscal.
  - `DocumentosComprobatorios`: Armazena documentos comprovantes, incluindo certidões cuja validade precisa ser verificada.
    - `ID`: Identificador único.
    - `AtivoID`: Relacionamento com a tabela `AtivosImobilizados`.
    - `TipoDocumento`: Tipo do documento (nota fiscal, contrato, certidão, etc.).
    - `DataEmissao`: Data de emissão do documento.
    - `DataValidade`: Data de validade da certidão (se aplicável).
    - `NumeroDocumento`: Número do documento.
    - `StatusValidacao`: Status de validação (aprovado, pendente, rejeitado).
  - `Beneficios`: Informações detalhadas dos benefícios concedidos.
    - `ID`: Identificador único.
    - `AtivoID`: Relacionamento com a tabela `AtivosImobilizados`.
    - `PercentualBeneficio`: Percentual do benefício fiscal.
    - `ValorBeneficio`: Valor concedido como benefício.
    - `StatusBeneficio`: Status do benefício (ativo, encerrado, etc.).

### 2. Planilhas Excel

#### 2.1. Database.xlsx

- **Objetivo**: Centralizar os dados de entrada que serão manipulados no Excel antes de serem enviados para o banco de dados SQL, e também gerenciar a anexação de documentação relevante.
- **Planilhas Internas**:
  - `EntradaDeDados`: Interface para entrada manual de dados (ativos, documentos, benefícios). Também permite anexar documentação diretamente na planilha.
  - `ConsultaDeDados`: Consultas SQL pré-definidas para buscar dados diretamente do banco de dados.
  - `DocumentacaoAnexada`: Gerenciamento de documentos anexados e verificação de validade de certidões.

#### 2.2. Analysis.xlsx

- **Objetivo**: Realizar análises detalhadas com base nos dados armazenados no banco de dados SQL.
- **Planilhas Internas**:
  - `AnaliseDeBeneficios`: Automatizar cálculos e fornecer visualizações sobre a eficácia dos benefícios fiscais.
  - `ValidadeCertidoes`: Verificar a validade das certidões anexadas e gerar relatórios automáticos sobre o status dessas certidões.

### 3. Módulos VBA

#### 3.1. db_integration.bas

- **Função**: Integração com o banco de dados SQL, incluindo operações de leitura (SELECT), inserção (INSERT), atualização (UPDATE) e exclusão (DELETE) de registros. Também gerencia a anexação e validação de documentos.
- **Principais Procedimentos**:
  - `ConectarBancoDeDados`: Conecta ao banco de dados SQL.
  - `InserirAtivo`: Insere novos ativos na tabela `AtivosImobilizados`.
  - `InserirDocumento`: Insere novos documentos na tabela `DocumentosComprobatorios`.
  - `ValidarCertidoes`: Verifica a validade das certidões e atualiza o status no banco de dados.
  - `GerarRelatorioBeneficios`: Gera relatórios automáticos com base nos dados do banco.

#### 3.2. analysis_automation.bas

- **Função**: Automatizar a análise de dados usando macros VBA.
- **Principais Procedimentos**:
  - `CalcularValorTotalBeneficios`: Calcula o valor total dos benefícios fiscais concedidos.
  - `GerarRelatorioDesempenho`: Cria relatórios sobre o desempenho e validade dos benefícios.
  - `GerarRelatorioValidadeCertidoes`: Gera relatórios sobre a validade das certidões, destacando as que estão próximas do vencimento ou já vencidas.

### 4. Power BI Dashboard

#### 4.1. Conexão com SQL

- **Objetivo**: Conectar o Power BI diretamente ao banco de dados SQL para extrair e visualizar dados, incluindo a anexação de documentos e verificação de validade.
- **Relatórios**:
  - **Resumo Geral dos Benefícios**: Gráficos e tabelas com o total de benefícios concedidos, ativos imobilizados e status dos documentos comprovantes.
  - **Análise por NCM**: Análise de benefícios por código NCM, identificando as categorias de ativos que recebem mais incentivos.
  - **Validade dos Benefícios e Certidões**: Visualização da vigência dos benefícios e validade das certidões, mostrando o período de início e término dos benefícios fiscais e destacando as certidões próximas do vencimento.

---

Este projeto visa agilizar e automatizar o processo de análise dos benefícios fiscais para ativos imobilizados, proporcionando maior eficiência e precisão na gestão das informações.
