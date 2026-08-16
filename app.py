import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Sistema de Gestão de Dados", layout="wide")
st.title("⚡ Painel de Integração (Controle Total)")

col1, col2 = st.columns(2)
with col1:
    source_file = st.file_uploader("1. Enviar Arquivo de Origem", type=["xlsx", "xls", "csv", "txt"])
    origem_sem_cabecalho = st.checkbox("O arquivo de origem NÃO tem cabeçalho", value=False)

with col2:
    dest_file = st.file_uploader("2. Enviar Arquivo de Destino", type=["xlsx", "xls"])
    header_dest = st.number_input("Em qual linha está o cabeçalho no Destino? (Linha do Excel)", value=11, min_value=1)

@st.cache_data(show_spinner=False)
def ler_destino(file, header_line):
    idx = header_line - 1
    xls = pd.ExcelFile(file)
    all_sheets = {sheet: pd.read_excel(file, sheet_name=sheet, header=idx) for sheet in xls.sheet_names}
    return xls.sheet_names, all_sheets

if source_file and dest_file:
    try:
        # Leitura da Origem
        header_val = None if origem_sem_cabecalho else 0
        if source_file.name.lower().endswith(".csv"):
            raw_df = pd.read_csv(source_file, header=header_val)
        elif source_file.name.lower().endswith(".txt"):
            raw_df = pd.read_csv(source_file, header=header_val, sep=None, engine='python')
        else:
            raw_df = pd.read_excel(source_file, header=header_val)
            
        if origem_sem_cabecalho:
            raw_df.columns = [f"Coluna {i+1}" for i in range(len(raw_df.columns))]

        # Inicializa estado se o arquivo mudar
        if "file_name" not in st.session_state or st.session_state["file_name"] != source_file.name:
            raw_df.insert(0, "Selecionar", False)
            st.session_state["source_df"] = raw_df
            st.session_state["file_name"] = source_file.name

        sheet_names, all_dest_dfs = ler_destino(dest_file, header_dest)
        selected_sheet = st.selectbox("3. Escolha a Aba de Destino:", sheet_names)
        dest_df = all_dest_dfs[selected_sheet]

        # --- MAPEAMENTO MANUAL ---
        st.subheader("4. Mapeamento de Colunas (Selecione manualmente)")
        source_cols = [c for c in st.session_state["source_df"].columns if c != "Selecionar"]
        
        # Mostra colunas da origem para conferência
        with st.expander("Ver colunas da Origem (Para conferir os nomes)"):
            st.write(source_cols)

        mapping = {}
        cols_ui = st.columns(3)
        for i, dest_col in enumerate(dest_df.columns):
            if "Unnamed" in str(dest_col): continue
            
            with cols_ui[i % 3]:
                # Inicia vazio para obrigar escolha
                escolha = st.selectbox(
                    f"Destino '{dest_col}':", 
                    options=["--- Não mapear ---"] + source_cols,
                    key=f"map_{dest_col}"
                )
                if escolha != "--- Não mapear ---":
                    mapping[dest_col] = escolha

        # --- PESQUISA ROBUSTA ---
        st.subheader("5. Pesquisa na Origem")
        search_col = st.selectbox("Buscar na coluna:", options=source_cols)
        search_term = st.text_input("Digite o termo de busca:", "")
        
        df_display = st.session_state["source_df"].copy()
        if search_term:
            # Filtro robusto
            mask = df_display[search_col].astype(str).str.contains(search_term, case=False, na=False)
            df_display = df_display[mask]

        # Edição e Sincronização
        edited_df = st.data_editor(df_display, use_container_width=True, hide_index=True)
        
        # Sincroniza a coluna "Selecionar" de volta para o session_state principal
        if st.button("Confirmar Alterações na Seleção"):
            # Atualiza apenas os índices que foram alterados na tabela editada
            for idx in edited_df.index:
                st.session_state["source_df"].loc[idx, "Selecionar"] = edited_df.loc[idx, "Selecionar"]
            st.success("Seleções atualizadas!")

        final_selected = st.session_state["source_df"][st.session_state["source_df"]["Selecionar"] == True]
        st.write(f"📌 Total de registros marcados: **{len(final_selected)}**")
        
        # --- PROCESSAMENTO ---
        st.subheader("6. Finalizar")
        if st.button("🚀 Processar"):
            if not mapping:
                st.warning("Nenhum mapeamento foi feito!")
            elif len(final_selected) == 0:
                st.error("Nenhum registro selecionado.")
            else:
                # Prepara dados
                cols_origem = list(mapping.values())
                new_data = final_selected[cols_origem].copy()
                new_data.columns = list(mapping.keys())
                
                # Concatena
                updated_sheet = pd.concat([dest_df, new_data], ignore_index=True)
                
                # Download
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                    updated_sheet.to_excel(writer, sheet_name=selected_sheet, index=False, startrow=header_dest)
                
                st.download_button("📥 Baixar Arquivo Processado", data=buffer.getvalue(), file_name="resultado.xlsx")

    except Exception as e:
        st.error(f"Erro: {e}")
