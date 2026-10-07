import pandas as pd

# Nome do arquivo que você baixa do portal
arquivo_entrada = "696-428.csv"
arquivo_final = "contratosbenner.xlsx"

def executar_automacao():
    try:
        # 1. Leitura com 'utf-8-sig' para remover caracteres invisíveis no início do arquivo
        # Usamos sep=';' porque seu CSV utiliza ponto e vírgula
        df = pd.read_csv(arquivo_entrada, sep=';', encoding='utf-8-sig')
        
        # 2. LIMPEZA CRÍTICA: Remove espaços extras nos nomes das colunas
        # Isso garante que 'Filial' seja lido como 'Filial' e não ' Filial' ou 'Filial\ufeff'
        df.columns = df.columns.str.strip()

        # 3. EXCLUSÃO DAS COLUNAS D, E, F, G, H e I (pela posição)
        # No Python: A=0, B=1, C=2, D=3, E=4, F=5, G=6, H=7, I=8
        # O comando abaixo remove as colunas da posição 3 até a 8
        colunas_para_remover = df.columns[3:9]
        df = df.drop(columns=colunas_para_remover)

        # 4. Normalização dos textos (remove espaços dentro das células)
        df = df.apply(lambda col: col.map(lambda x: x.strip() if isinstance(x, str) else x))

        # 5. Configuração das Chaves de Agrupamento
        chaves = ["Filial", "Pasta", "Cadastrado em"]

        # 6. Agrupamento e União dos Contratos
        # Aqui ele une os contratos da mesma pasta em uma única linha, separados por vírgula
        agrupado = df.groupby(chaves, as_index=False).agg({
            "Contrato": lambda x: ", ".join(x.astype(str).unique())
        })

        # 7. Salvar o resultado final
        agrupado.to_excel(arquivo_final, index=False)
        
        print(f"✅ SUCESSO: O arquivo '{arquivo_final}' foi gerado com os contratos unidos!")

    except Exception as e:
        print(f"❌ ERRO DURANTE A EXECUÇÃO: {e}")

if __name__ == "__main__":
    executar_automacao()