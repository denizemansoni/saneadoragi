📑 Consolidador de Contratos Benner (Auto-Clean)
🚀 Descrição
Este script automatiza a limpeza e o agrupamento de registros de contratos extraídos de portais jurídicos. Ele elimina a necessidade de edição manual no Excel, realizando a exclusão de colunas desnecessárias e a consolidação de múltiplos contratos em uma única linha por processo.

Antes: Edição manual de colunas + conversão de arquivo + agrupamento.

Depois: Download do CSV + Execução do script = Resultado pronto.

🛠️ Pré-requisitos
Python 3.8+

Bibliotecas: pandas, openpyxl

📦 Instalação Rápida
Criar e ativar ambiente virtual (Windows):

Bash
python -m venv .venv
.venv\Scripts\activate
Instalar dependências:

Bash
pip install pandas openpyxl
📂 Como Usar
Baixe o relatório do portal no formato .csv.

Renomeie o arquivo para 696-428.csv (ou o nome definido no script) e coloque-o na pasta do projeto.

Execute o script:

Bash
python unircontratos.py
O script gerará automaticamente o arquivo contratosbenner.xlsx com os dados limpos e organizados.

🧠 Inteligência do Processo (Resumo Técnico)
O script executa um fluxo de ETL (Extract, Transform, Load) completo:

Leitura Blindada: Utiliza utf-8-sig para evitar erros de caracteres invisíveis na primeira coluna.

Auto-Clean: - Remove automaticamente as colunas D, E, F, G, H e I (posições 3 a 8).

Realiza o strip() em todos os nomes de colunas e células para eliminar espaços fantasmas.

Agrupamento Estratégico: Identifica registros únicos através das chaves Filial, Pasta e Cadastrado em.

Consolidação: Une múltiplos números de contrato em uma única célula, separados por vírgula, removendo duplicatas.

Conversão de Formato: Exporta o resultado final diretamente para .xlsx.

📊 Formato de Saída
O arquivo contratosbenner.xlsx conterá:

As colunas de identificação preservadas.

A coluna Contrato contendo todos os itens agrupados.
