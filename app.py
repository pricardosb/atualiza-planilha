import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Sistema de Gestão de Dados", layout="wide")
st.title("⚡ Painel de Integração e Seleção")

col1, col2 = st.columns(2)
with col1:
    source_file = st.file_uploader("1. Enviar Arquivo de Origem", type=["xlsx", "xls", "csv"])
with col2:
    dest_file = st.file_uploader("2. Enviar Arquivo de Destino (Cabeçalho na linha 11)", type=["xlsx", "xls"])

@st.cache_data(show_spinner=False)
def ler_origem(file):
    if file.name.lower().endswith(".csv"):
        return pd.read_csv(file)
    return pd.read_excel(file)

@st.cache_data(show_spinner=False)
def ler_destino(file):
    xls = pd.ExcelFile(file)
    all_sheets = {sheet: pd.read_excel(file, sheet_name=sheet, header=10) for sheet in xls.sheet_names}
    return xls.sheet_names, all_sheets

if source_file and dest_file:
    try:
        # Garante que a base de dados da origem fica gravada na memória sem perder seleções
        if "file_name" not in st.session_state or st.session_state["file_name"] != source_file.name:
            raw_df = ler_origem(source_file)
            raw_df.insert(0, "Selecionar", False)
            st.session_state["source_df"] = raw_df
            st.session_state["file_name"] = source_file.name

        sheet_names, all_dest_dfs = ler_destino(dest_file)
        
        st.success("⚡ Arquivos carregados com sucesso!")
        
        # --- Seleção da Aba ---
        selected_sheet = st.selectbox("3. Escolha a Aba (Planilha) de Destino:", sheet_names)
        dest_df = all_dest_dfs[selected_sheet]
        
        # --- Mapeamento (De/Para) ---
        st.subheader("4. Mapeamento (Você escolhe o que cada campo do Destino recebe)")
        mapping = {}
        source_cols = [c for c in st.session_state["source_df"].columns if c != "Selecionar"]
        
        for dest_col in dest_df.columns:
            if "Unnamed" in str(dest_col): continue
            mapping[dest_col] = st.selectbox(
                f"O campo '{dest_col}' do Destino recebe da Origem:", 
                options=source_cols, 
                key=f"map_{dest_col}"
            )
            
        # --- Pesquisa e Seleção Instantânea ---
        st.subheader("5. Selecionar Dados da Origem")
        search = st.text_input("🔍 Pesquisa Instantânea (digite para filtrar a visualização):", "")
        
        df_atual = st.session_state["source_df"]
        
        if search:
            mask = df_atual.astype(str).apply(lambda row: row.str.contains(search, case=False, regex=False).any(), axis=1)
            # Mostra o que deu match OU o que o usuário já marcou anteriormente
            df_display = df_atual[mask | (df_atual["Selecionar"] == True)]
        else:
            df_display = df_atual
            
        # Tabela interativa para seleção
        edited_df = st.data_editor(df_display, use_container_width=True, hide_index=True, key="editor_origem")
        
        # Sincroniza as marcações feitas na tela com a memória principal
        for idx in edited_df.index:
            st.session_state["source_df"].loc[idx, "Selecionar"] = edited_df.loc[idx, "Selecionar"]
            
        # Conta quantos foram marcados no total geral
        final_selected = st.session_state["source_df"][st.session_state["source_df"]["Selecionar"] == True]
        st.write(f"📌 Total de registros marcados para envio: **{len(final_selected)}**")
        
        # --- Inserção ---
        st.subheader("6. Finalizar")
        mode = st.radio("Onde salvar?", ["Final do arquivo", "Em uma linha específica"])
        target_line = 0
        if mode == "Em uma linha específica":
            target_line = st.number_input("Número da linha (após o cabeçalho):", min_value=0, max_value=len(dest_df))
            
        if st.button("🚀 Processar e Gerar Novo Arquivo"):
            if len(final_selected) == 0:
                st.error("Nenhum registro foi marcado! Marque a caixinha 'Selecionar' na tabela de origem.")
            else:
                cols_origem = list(mapping.values())
                new_data = final_selected[cols_origem].copy()
                new_data.columns = list(mapping.keys())
                
                current_dest = all_dest_dfs[selected_sheet].copy()
                
                if mode == "Final do arquivo":
                    updated_sheet = pd.concat([current_dest, new_data], ignore_index=True)
                else:
                    updated_sheet = pd.concat([current_dest.iloc[:target_line], new_data, current_dest.iloc[target_line:]], ignore_index=True)
                
                output_dfs = all_dest_dfs.copy()
                output_dfs[selected_sheet] = updated_sheet
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                    for sh_name, sh_df in output_dfs.items():
                        sh_df.to_excel(writer, sheet_name=sh_name, index=False, startrow=10)
                
                st.success("Tudo pronto! Arquivo gerado com sucesso.")
                st.download_button("📥 Baixar Planilha Atualizada", data=buffer.getvalue(), file_name="arquivo_final.xlsx")

    except Exception as e:
        st.error(f"Ocorreu um erro: {e}")
