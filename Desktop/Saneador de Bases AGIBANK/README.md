---

# 📑 README - Robô de Bate e Saneamento AGIBANK

## 🎯 Objetivo do Sistema

Este script foi desenvolvido para automatizar o cruzamento de bases ("bate") entre o relatório judicial enviado pelo cliente (**Agibank**) e os dados dos sistemas internos do escritório. O objetivo principal é **identificar novos processos que precisam ser cadastrados**, eliminando o trabalho manual de conferência.

---

## 🛡️ Filosofia de Desenvolvimento: Descarte Zero (0% de Perda)

Uma das maiores preocupações em automações jurídicas é o risco de o código "ignorar" ou "apagar" linhas importantes de forma silenciosa. **Este robô foi projetado sob a regra estrita de exclusão zero.**

### Como o robô trata 100% das linhas do cliente?

1. **Varredura Completa:** O script lê absolutamente todas as linhas presentes no arquivo `PBI JUDICIAL.xlsx`. Nenhuma linha é deletada da memória por filtros de data, palavras-chave ou regras de corte.
2. **Separação por Não-Existência:** A única filtragem realizada é uma divisão lógica. O robô testa cada processo:
* **Se o processo NÃO existe no sistema interno:** Ele é separado e enviado para o relatório final de **"Casos Novos"**.
* **Se o processo JÁ existe no sistema interno:** Ele apenas deixa de ser listado no relatório de saída (afinal, se já está no sistema, não precisa ser cadastrado novamente). Mesmo assim, ele foi lido e processado pelo motor do script.



> ⚠️ **Nota de Transparência Técnica:** Modificações de texto (como a limpeza da coluna `nr_cpf_adverso`, que remove as palavras "CPF:" ou "CNPJ" para deixar apenas os números limpos) alteram o conteúdo da célula para fins de padronização, mas **JAMAIS excluem a linha** do relatório.

---

## 📂 Estrutura de Pastas de Trabalho

Para que o script funcione no formato executável (`.exe`), ele utiliza uma arquitetura de caminhos dinâmicos. Ao ser rodado pela primeira vez, ele cria automaticamente a seguinte estrutura no mesmo local onde o `.exe` está salvo:

* **`📂 Entrada/`**: Local onde o usuário deve depositar os relatórios antes de rodar o programa.
* O arquivo do cliente deve **começar** com: `PBI JUDICIAL` (ex: `PBI JUDICIAL_JUNHO.xlsx`)
* Os arquivos internos devem **começar** com: `ATIVOS AGI` e `ENCERRADOS AGI`


* **`📂 Saida/`**: Local onde o robô salvará o resultado final purificado.
* Nome do arquivo gerado: `CADASTRO_AGIBANK_DD-MM-AAAA.xlsx`



---

## 📈 Recursos e Melhorias Implementadas

* **Prefixos Dinâmicos:** O robô não exige nomes de arquivos 100% exatos. Se o cliente mudar a data ou salvar como `PBI JUDICIAL (1).xlsx`, o script reconhecerá o arquivo pelo início do nome, evitando quebras e chamados de suporte.
* **Ordenação Logística por UF:** Os casos novos são ordenados de forma alfabética crescente (A-Z) pela coluna `ds_estado` para facilitar a distribuição e triagem da equipe de cadastro.
* **Controle de Visualização (Temporizador de 10s):** Ao finalizar o processamento (ou encontrar algum erro crítico), o terminal exibe uma contagem regressiva de 10 segundos na tela, garantindo tempo hábil para o usuário ler o resumo da operação antes da janela fechar.

---

## 🚀 Como Executar

1. Cole os relatórios atualizados do cliente e do sistema interno dentro da pasta `Entrada`.
2. Dê um duplo clique no executável do robô.
3. Aguarde o aviso de conclusão na tela e a contagem regressiva.
4. Abra a pasta `Saida` para coletar sua planilha de trabalho 100% higienizada e pronta para cadastro.