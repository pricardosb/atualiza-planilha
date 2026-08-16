import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Gestão de Dados Pro", layout="wide")
st.title("⚡ Painel de Controle de Integração")

# --- BLOCO DE ARQUIVOS ---
col1, col2 = st.columns(2)
with col1:
    source_file = st.file_uploader("1. Enviar Origem", type=["xlsx", "xls", "csv", "txt"])
    origem_sem_cabecalho = st.checkbox("Origem NÃO tem cabeçalho", value=False)
with col2:
    dest_file = st.file_uploader("2. Enviar Destino", type=["xlsx", "xls"])
    header_dest = st.number_input("Em qual linha está o cabeçalho no Destino?", value=11, min_value=1)

if source_file and dest_file:
    # 1. Carregar Origem
    if "source_df" not in st.session_state or st.session_state.get("last_file") != source_file.name:
        header_val = None if origem_sem_cabecalho else 0
        if source_file.name.endswith(".csv"): raw = pd.read_csv(source_file, header=header_val)
        elif source_file.name.endswith(".txt"): raw = pd.read_csv(source_file, header=header_val, sep=None, engine='python')
        else: raw = pd.read_excel(source_file, header=header_val)
        
        if origem_sem_cabecalho: raw.columns = [f"Col {i+1}" for i in range(len(raw.columns))]
        raw.insert(0, "Selecionar", False)
        st.session_state["source_df"] = raw
        st.session_state["last_file"] = source_file.name

    # 2. Carregar Destino
    xls = pd.ExcelFile(dest_file)
    sheets = {s: pd.read_excel(dest_file, sheet_name=s, header=header_dest-1) for s in xls.sheet_names}
    selected_sheet = st.selectbox("3. Escolha a Aba de Destino:", list(sheets.keys()))
    dest_df = sheets[selected_sheet]

    # 3. MAPEAMENTO MANUAL (Obrigatório escolher)
    st.subheader("4. Mapear Colunas")
    source_cols = [c for c in st.session_state["source_df"].columns if c != "Selecionar"]
    mapping = {}
    
    cols_ui = st.columns(4)
    for i, col_name in enumerate(dest_df.columns):
        if "Unnamed" in str(col_name): continue
        with cols_ui[i % 4]:
            sel = st.selectbox(f"Destino '{col_name}':", ["--- Não mapear ---"] + source_cols, key=f"map_{col_name}")
            if sel != "--- Não mapear ---": mapping[col_name] = sel

    # 4. PESQUISA FUNCIONAL (Input box real)
    st.subheader("5. Pesquisa e Seleção")
    c1, c2 = st.columns([1, 2])
    with c1: search_col = st.selectbox("Buscar na coluna:", source_cols)
    with c2: search_term = st.text_input("🔍 Digite aqui para filtrar os dados:")

    # Filtra view, mas mantém estado no source_df
    df_to_edit = st.session_state["source_df"].copy()
    if search_term:
        mask = df_to_edit[search_col].astype(str).str.contains(search_term, case=False, na=False)
        df_to_edit = df_to_edit[mask | (df_to_edit["Selecionar"] == True)]

    # Edição segura
    edited = st.data_editor(df_to_edit, use_container_width=True, hide_index=True)
    
    # Sincroniza o que foi marcado/desmarcado
    for idx in edited.index:
        st.session_state["source_df"].at[idx, "Selecionar"] = edited.at[idx, "Selecionar"]

    # 5. MODO DE INSERÇÃO
    st.subheader("6. Opções de Inserção")
    modo = st.radio("Onde inserir?", ["Final da planilha", "Em uma linha específica"])
    target_row = 0
    if modo == "Em uma linha específica":
        target_row = st.number_input("Inserir a partir da linha (relativo ao cabeçalho):", min_value=0)

    if st.button("🚀 Processar e Atualizar"):
        final_data = st.session_state["source_df"][st.session_state["source_df"]["Selecionar"] == True]
        if final_data.empty: st.error("Nenhum dado selecionado!")
        else:
            # Prepara dados
            new_rows = final_data[list(mapping.values())].copy()
            new_rows.columns = list(mapping.keys())
            
            # Concatena preservando a estrutura
            if modo == "Final da planilha":
                updated = pd.concat([dest_df, new_rows], ignore_index=True)
            else:
                updated = pd.concat([dest_df.iloc[:target_row], new_rows, dest_df.iloc[target_row:]], ignore_index=True)
            
            # Gera download
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                updated.to_excel(writer, sheet_name=selected_sheet, index=False, startrow=header_dest)
            
            st.success("Tudo pronto! Planilha processada.")
            st.download_button("📥 Baixar Planilha Atualizada", data=buffer.getvalue(), file_name="destino_atualizado.xlsx")
