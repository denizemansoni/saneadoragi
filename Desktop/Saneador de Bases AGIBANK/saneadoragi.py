import pandas as pd
import os
import sys
import time
from datetime import datetime
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment

# --- ARQUITETURA DE CAMINHOS DINÂMICOS (Suporte para executável .EXE) ---
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

PASTA_ENTRADA = os.path.join(BASE_DIR, "Entrada")
PASTA_SAIDA = os.path.join(BASE_DIR, "Saida")

# Garante a existência das pastas no ambiente
os.makedirs(PASTA_ENTRADA, exist_ok=True)
os.makedirs(PASTA_SAIDA, exist_ok=True)

# --- CONFIGURAÇÃO DE PREFIXOS DINÂMICOS DE ARQUIVOS ---
PREFIXO_CLIENTE = "PBI JUDICIAL"
PREFIXOS_INTERNOS = ["ATIVOS AGI", "ENCERRADOS AGI"]
PREFIXOS_EXCECOES = ["EXCECOES", "DIVERGENTES", "CASOS_ESPECIAIS"]

def extrair_numeros(txt):
    """Remove caracteres não numéricos para criar chaves puras de comparação."""
    if pd.isna(txt): return ""
    return "".join(filter(str.isdigit, str(txt)))

def carregar_excel_seguro(caminho):
    """Carrega arquivos Excel contornando possíveis linhas de cabeçalhos institucionais."""
    try:
        df = pd.read_excel(caminho)
        if df.empty: return pd.DataFrame()
        if "Unnamed" in str(df.columns[0]) or "RELATÓRIO" in str(df.columns[0]):
            df = pd.read_excel(caminho, skiprows=1)
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception:
        return pd.DataFrame()

def executar():
    print("🚀 Iniciando triagem AGIBANK...")
    
    if not os.path.exists(PASTA_ENTRADA):
        return "❌ Erro: Pasta 'Entrada' não localizada no diretório."
        
    arquivos_na_pasta = os.listdir(PASTA_ENTRADA)
    
    # -------------------------------------------------------------
    # 1. CARREGAMENTO E MAPEAMENTO DAS BASES INTERNAS
    # -------------------------------------------------------------
    arquivos_internos_reais = [
        os.path.join(PASTA_ENTRADA, f) for f in arquivos_na_pasta 
        if any(f.startswith(pref) for pref in PREFIXOS_INTERNOS) and f.endswith('.xlsx')
    ]
    
    mapa_sistema_processos = {}  # { 'processo_limpo': 'Nome_Arquivo' }
    mapa_sistema_pastas = {}     # { 'pasta_limpa': 'Nome_Arquivo' }

    for arq in arquivos_internos_reais:
        nome_base = os.path.basename(arq)
        df = carregar_excel_seguro(arq)
        
        col_proc = next((c for c in df.columns if "NÚMERO DO PROCESSO" in c.upper() or "NR_PROCESSO" in c.upper()), None)
        col_pasta = next((c for c in df.columns if "PASTA" in c.upper() or "CÓDIGO" in c.upper() or "CODIGO" in c.upper()), None)

        if col_proc:
            for p in df[col_proc].dropna().astype(str):
                p_limpo = extrair_numeros(p)
                if p_limpo:
                    mapa_sistema_processos[p_limpo] = nome_base
        if col_pasta:
            for pst in df[col_pasta].dropna().astype(str):
                pst_limpa = extrair_numeros(pst)
                if pst_limpa:
                    mapa_sistema_pastas[pst_limpa] = nome_base

    # -------------------------------------------------------------
    # 2. CARREGAMENTO DAS DIVERGÊNCIAS / EXCEÇÕES
    # -------------------------------------------------------------
    arquivos_excecao = [
        os.path.join(PASTA_ENTRADA, f) for f in arquivos_na_pasta 
        if any(f.upper().startswith(pref) for pref in PREFIXOS_EXCECOES) and f.endswith('.xlsx')
    ]

    dict_excecoes_proc = {}   # { 'processo_limpo': 'motivo/origem' }
    dict_excecoes_pasta = {}  # { 'pasta_limpa': 'motivo/origem' }

    for arq_exc in arquivos_excecao:
        nome_exc = os.path.basename(arq_exc)
        print(f"📌 Lendo base de Exceções/Divergências: {nome_exc}")
        df_exc = carregar_excel_seguro(arq_exc)
        
        col_proc = next((c for c in df_exc.columns if "PROCESSO" in c.upper() or "NR_PROCESSO" in c.upper()), None)
        col_pasta = next((c for c in df_exc.columns if "PASTA" in c.upper() or "CODIGO" in c.upper()), None)

        if col_proc:
            for val in df_exc[col_proc].dropna().astype(str):
                chave = extrair_numeros(val)
                if chave: dict_excecoes_proc[chave] = nome_exc
        if col_pasta:
            for val in df_exc[col_pasta].dropna().astype(str):
                chave = extrair_numeros(val)
                if chave: dict_excecoes_pasta[chave] = nome_exc

    # -------------------------------------------------------------
    # 3. LEITURA DA BASE DO CLIENTE
    # -------------------------------------------------------------
    arq_cliente_nome = next((f for f in arquivos_na_pasta if f.startswith(PREFIXO_CLIENTE) and f.endswith('.xlsx')), None)
    
    if not arq_cliente_nome:
        return f"❌ Erro: Nenhum arquivo começando com '{PREFIXO_CLIENTE}' foi encontrado na pasta Entrada."
        
    caminho_cliente = os.path.join(PASTA_ENTRADA, arq_cliente_nome)
    print(f"📂 Lendo relatório do Cliente: {arq_cliente_nome}")
    df_cliente = carregar_excel_seguro(caminho_cliente)

    if 'nr_processo' not in df_cliente.columns:
        return f"❌ Erro: Coluna 'nr_processo' não localizada no arquivo do cliente."

    # Criando DataFrame de Auditoria Mantendo 100% das Linhas Originais
    df_auditoria = df_cliente.copy()
    
    # Gerando chaves internas para cruzamento
    df_auditoria['CHAVE_PROC_LIMPO'] = df_auditoria['nr_processo'].apply(extrair_numeros)
    
    col_pasta_cli = next((c for c in df_auditoria.columns if "pasta" in c.lower() or "cd_pasta" in c.lower()), None)
    if col_pasta_cli:
        df_auditoria['CHAVE_PASTA_LIMPA'] = df_auditoria[col_pasta_cli].apply(extrair_numeros)
    else:
        df_auditoria['CHAVE_PASTA_LIMPA'] = ""

    # -------------------------------------------------------------
    # 4. MOTOR DE CLASSIFICAÇÃO E AUDITORIA
    # -------------------------------------------------------------
    status_list = []
    mensagem_list = []

    for index, row in df_auditoria.iterrows():
        proc = row['CHAVE_PROC_LIMPO']
        pasta = row['CHAVE_PASTA_LIMPA']
        
        # Validação 1: Cadastro nas Bases Internas (Ativos / Encerrados)
        if proc in mapa_sistema_processos:
            origem = mapa_sistema_processos[proc]
            status_list.append("JÁ CADASTRADO")
            mensagem_list.append(f"Processo localizado na base interna ({origem})")
        elif pasta and pasta in mapa_sistema_pastas:
            origem = mapa_sistema_pastas[pasta]
            status_list.append("JÁ CADASTRADO (POR PASTA)")
            mensagem_list.append(f"Pasta cadastrada na base interna ({origem})")
            
        # Validação 2: Presença na Planilha de Exceções / Divergências
        elif proc in dict_excecoes_proc:
            origem_exc = dict_excecoes_proc[proc]
            status_list.append("ALERTA: DIVERGÊNCIA")
            mensagem_list.append(f"Caso com divergência não corrigível registrado em {origem_exc}")
        elif pasta and pasta in dict_excecoes_pasta:
            origem_exc = dict_excecoes_pasta[pasta]
            status_list.append("ALERTA: DIVERGÊNCIA")
            mensagem_list.append(f"Pasta com divergência não corrigível registrada em {origem_exc}")
            
        # Validação 3: Casos Novos
        else:
            status_list.append("CASO NOVO")
            mensagem_list.append("Pronto para cadastramento no sistema")

    # Adicionando os dados de auditoria ao DataFrame Completo
    df_auditoria['STATUS_AUDITORIA'] = status_list
    df_auditoria['MENSAGEM_AUDITORIA'] = mensagem_list

    # -------------------------------------------------------------
    # 5. SEPARAÇÃO DAS ABAS DE SAÍDA
    # -------------------------------------------------------------
    # Aba 1: Casos Novos
    df_casos_novos = df_auditoria[df_auditoria['STATUS_AUDITORIA'] == 'CASO NOVO'].copy()
    if 'ds_estado' in df_casos_novos.columns:
        df_casos_novos = df_casos_novos.sort_values(by=['ds_estado'], ascending=True)
    
    df_casos_novos['STATUS_CADASTRO'] = "Caso Novo - Liberado para Cadastro"

    # Aba 3: Divergências / Alertas
    df_divergencias = df_auditoria[df_auditoria['STATUS_AUDITORIA'].str.contains("DIVERGÊNCIA")].copy()

    # Limpeza de colunas temporárias de chave
    cols_para_remover = ['CHAVE_PROC_LIMPO', 'CHAVE_PASTA_LIMPA']
    df_auditoria = df_auditoria.drop(columns=cols_para_remover, errors='ignore')
    df_casos_novos = df_casos_novos.drop(columns=cols_para_remover, errors='ignore')
    df_divergencias = df_divergencias.drop(columns=cols_para_remover, errors='ignore')

    # Trata visualização do CPF/CNPJ
    col_doc = 'nr_cpf_adverso'
    for df_target in [df_auditoria, df_casos_novos, df_divergencias]:
        if col_doc in df_target.columns:
            df_target[col_doc] = df_target[col_doc].astype(str).str.replace(r'(?i)CPF|CNPJ', '', regex=True).str.strip()

    # -------------------------------------------------------------
    # 6. EXPORTAÇÃO E FORMATAÇÃO COM OPENPYXL
    # -------------------------------------------------------------
    data_hoje = datetime.now().strftime('%d-%m-%Y')
    nome_arquivo = os.path.join(PASTA_SAIDA, f"RELATORIO_AUDITORIA_AGIBANK_{data_hoje}.xlsx")
    
    writer = pd.ExcelWriter(nome_arquivo, engine='openpyxl')
    
    # Gravação das Abas
    df_casos_novos.to_excel(writer, index=False, sheet_name='Casos Novos')
    df_auditoria.to_excel(writer, index=False, sheet_name='Auditoria Completa')
    if not df_divergencias.empty:
        df_divergencias.to_excel(writer, index=False, sheet_name='Casos com Divergencia')

    # Estilização Corporativa
    cor_cabecalho = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
    cor_alerta = PatternFill(start_color="FCE5CD", end_color="FCE5CD", fill_type="solid")
    fonte_corpo = Font(name='Calibri Light', size=10)
    fonte_cabecalho = Font(name='Calibri Light', size=10, bold=True)
    borda = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    alinhamento_padrao = Alignment(horizontal='left', vertical='center')

    for sheet_name in writer.sheets:
        ws = writer.sheets[sheet_name]
        
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                cell.font = fonte_corpo
                cell.border = borda
                cell.alignment = alinhamento_padrao
                
                if cell.row == 1:
                    cell.fill = cor_cabecalho
                    cell.font = fonte_cabecalho
                else:
                    header_name = str(ws.cell(1, cell.column).value).lower()
                    if "data" in header_name or "dt_" in header_name:
                        cell.number_format = 'dd/mm/yyyy'
                    elif "valor" in header_name or "vl_" in header_name:
                        cell.number_format = '#,##0.00'
                    
                    # Destaque para linhas/células de divergência
                    if "status" in header_name and "DIVERGÊNCIA" in str(cell.value):
                        cell.fill = cor_alerta

        # Largura automática padrão
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 25

    writer.close()
    
    return (
        f"🎯 PROCESSAMENTO E AUDITORIA CONCLUÍDOS!\n\n"
        f"📊 Linhas do Relatório do Cliente: {len(df_auditoria)}\n"
        f"✨ Casos Novos para Cadastrar: {len(df_casos_novos)}\n"
        f"⚠️ Casos com Divergência/Exceção: {len(df_divergencias)}\n\n"
        f"📂 Arquivo gerado em: {nome_arquivo}"
    )

if __name__ == "__main__":
    mensagem_final = executar()
    
    print("\n" + "="*60)
    print(mensagem_final)
    print("="*60)
    
    print("\nO programa será fechado automaticamente em:")
    for i in range(10, 0, -1):
        print(f"{i} segundos...", end="\r")
        time.sleep(1)