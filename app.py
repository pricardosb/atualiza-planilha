import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import io

st.set_page_config(page_title="Integrador Profissional", layout="wide")
st.title("⚡ Integrador: Edição Direta no Destino")

# --- 1. CARREGAMENTO DOS ARQUIVOS ---
col1, col2 = st.columns(2)
with col1:
    source_file = st.file_uploader("1. Arquivo de ORIGEM (Fonte)", type=["xlsx", "xls", "csv", "txt"])
with col2:
    dest_file = st.file_uploader("2. Arquivo de DESTINO (Original)", type=["xlsx", "xls"])
    header_dest = st.number_input("Linha do cabeçalho no Destino:", value=11, min_value=1)

if source_file and dest_file:
    # --- LEITURA DA ORIGEM ---
    if "source_df" not in st.session_state or st.session_state.get("last_source") != source_file.name:
        if source_file.name.endswith(".csv"): raw = pd.read_csv(source_file)
        elif source_file.name.endswith(".txt"): raw = pd.read_csv(source_file, sep=None, engine='python')
        else: raw = pd.read_excel(source_file)
        st.session_state["source_df"] = raw
        st.session_state["last_source"] = source_file.name
    
    df_origem = st.session_state["source_df"]

    # --- 2. SELEÇÃO DA ABA (PASTA) DO DESTINO ---
    # Carrega o workbook apenas para listar as abas
    wb = load_workbook(dest_file, data_only=True)
    sheet_names = wb.sheetnames
    target_sheet = st.selectbox("3. Escolha a ABA (Pasta) onde inserir os dados:", sheet_names)
    
    # --- 3. PESQUISA ORDENADA ---
    st.subheader("4. Seleção de Registros da Origem")
    col_busca = st.selectbox("Coluna para pesquisa na Origem:", df_origem.columns)
    
    # Gera opções ORDENADAS
    opcoes = [f"{val} (Linha {idx})" for idx, val in df_origem[col_busca].items()]
    opcoes.sort() 
    
    selected_options = st.multiselect("🔍 Selecione os registros (Ordem alfabética):", opcoes)
    selected_indices = [int(item.split("(Linha ")[1].replace(")", "")) for item in selected_options]
    
    st.write(f"📌 Total selecionados: {len(selected_indices)}")

    # --- 4. MAPEAMENTO MANUAL ---
    st.subheader("5. Mapeamento de Colunas")
    
    # Identifica colunas da aba destino na linha do cabeçalho
    ws = wb[target_sheet]
    header_row = list(ws.iter_rows(min_row=header_dest, max_row=header_dest, values_only=True))[0]
    # Filtra colunas com nomes (ignorando None/vazio)
    dest_cols = [str(col) for col in header_row if col is not None]
    
    mapping = {}
    cols_ui = st.columns(3)
    for i, d_col in enumerate(dest_cols):
        with cols_ui[i % 3]:
            map_val = st.selectbox(f"Destino '{d_col}' recebe de:", ["--- Não mapear ---"] + list(df_origem.columns), key=f"map_{d_col}")
            if map_val != "--- Não mapear ---":
                mapping[d_col] = map_val

    # --- 5. INSERÇÃO ---
    st.subheader("6. Opções de Inserção")
    modo = st.radio("Onde inserir?", ["Final da planilha", "Em uma linha específica"])
    target_row = st.number_input("Linha específica:", min_value=header_dest + 1, value=header_dest + 1)

    if st.button("🚀 Processar e Atualizar Arquivo Original"):
        if not selected_indices:
            st.error("Selecione registros!")
        else:
            # Prepara os dados selecionados
            dados_para_inserir = df_origem.iloc[selected_indices]
            
            # Mapeia colunas para encontrar índices no Excel
            # Cria um dicionário de {nome_coluna_destino: indice_coluna_excel}
            header_map = {str(val): i+1 for i, val in enumerate(header_row) if val is not None}
            
            # Insere os dados
            if modo == "Final da planilha":
                current_row = ws.max_row + 1
            else:
                current_row = target_row
                ws.insert_rows(current_row, amount=len(dados_para_inserir))
            
            # Grava as linhas
            for idx, row in dados_para_inserir.iterrows():
                for dest_col, orig_col in mapping.items():
                    if dest_col in header_map:
                        ws.cell(row=current_row, column=header_map[dest_col], value=row[orig_col])
                current_row += 1
            
            # Salva
            buffer = io.BytesIO()
            wb.save(buffer)
            
            st.success("Arquivo atualizado com sucesso!")
            st.download_button("📥 Baixar Arquivo Atualizado", data=buffer.getvalue(), file_name="destino_atualizado.xlsx")
