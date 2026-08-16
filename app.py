import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Integrador Profissional", layout="wide")
st.title("⚡ Integrador: Edição de Arquivos XLS/XLSX")

# --- 1. CARREGAMENTO ---
col1, col2 = st.columns(2)
with col1:
    source_file = st.file_uploader("1. Arquivo de ORIGEM", type=["xlsx", "xls", "csv", "txt"])
with col2:
    dest_file = st.file_uploader("2. Arquivo de DESTINO (.xls ou .xlsx)", type=["xlsx", "xls"])
    header_dest = st.number_input("Linha do cabeçalho no Destino:", value=11, min_value=1)

if source_file and dest_file:
    # --- LEITURA ORIGEM ---
    if "source_df" not in st.session_state or st.session_state.get("last_source") != source_file.name:
        if source_file.name.endswith(".csv"): raw = pd.read_csv(source_file)
        elif source_file.name.endswith(".txt"): raw = pd.read_csv(source_file, sep=None, engine='python')
        else: raw = pd.read_excel(source_file)
        st.session_state["source_df"] = raw
        st.session_state["last_source"] = source_file.name
    
    df_origem = st.session_state["source_df"]

    # --- CARREGAMENTO DESTINO (LEITURA DE TODAS AS ABAS) ---
    try:
        # Lê todas as abas do arquivo destino
        all_sheets = pd.read_excel(dest_file, sheet_name=None, header=header_dest-1)
        target_sheet = st.selectbox("3. Escolha a ABA (Pasta) onde inserir:", list(all_sheets.keys()))
        dest_df = all_sheets[target_sheet]
    except Exception as e:
        st.error(f"Erro ao ler o arquivo de Destino: {e}")
        st.stop()

    # --- PESQUISA ORDENADA ---
    st.subheader("4. Seleção de Registros")
    col_busca = st.selectbox("Coluna para pesquisa na Origem:", df_origem.columns)
    opcoes = sorted([f"{val} (Linha {idx})" for idx, val in df_origem[col_busca].items()])
    
    selected_options = st.multiselect("🔍 Selecione os registros (Ordem alfabética):", opcoes)
    selected_indices = [int(item.split("(Linha ")[1].replace(")", "")) for item in selected_options]
    
    # --- MAPEAMENTO ---
    st.subheader("5. Mapeamento")
    # Colunas que existem no seu destino
    dest_cols = [c for c in dest_df.columns if "Unnamed" not in str(c)]
    
    mapping = {}
    cols_ui = st.columns(3)
    for i, d_col in enumerate(dest_cols):
        with cols_ui[i % 3]:
            map_val = st.selectbox(f"Destino '{d_col}' recebe de:", ["--- Não mapear ---"] + list(df_origem.columns), key=f"map_{i}")
            if map_val != "--- Não mapear ---":
                mapping[d_col] = map_val

    # --- INSERÇÃO ---
    st.subheader("6. Atualização")
    if st.button("🚀 Processar e Gerar Arquivo"):
        if not selected_indices:
            st.error("Selecione registros!")
        else:
            # Prepara dados
            dados = df_origem.iloc[selected_indices][list(mapping.values())].copy()
            dados.columns = list(mapping.keys())
            
            # Atualiza apenas a aba escolhida no dicionário de abas
            all_sheets[target_sheet] = pd.concat([dest_df, dados], ignore_index=True)
            
            # Salva o arquivo (o formato .xlsx é o recomendado para garantir que nada quebre)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                for sheet_name, df in all_sheets.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=header_dest-1)
            
            st.success("Arquivo processado com sucesso!")
            st.download_button("📥 Baixar Arquivo Atualizado (.xlsx)", data=buffer.getvalue(), file_name="destino_atualizado.xlsx")
