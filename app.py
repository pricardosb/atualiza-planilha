import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Sistema de Gestão de Dados", layout="wide")
st.title("⚡ Painel de Integração (Excel, CSV e TXT)")

col1, col2 = st.columns(2)
with col1:
    source_file = st.file_uploader("1. Enviar Arquivo de Origem", type=["xlsx", "xls", "csv", "txt"])
    origem_sem_cabecalho = st.checkbox("O arquivo de origem NÃO tem cabeçalho", value=False)

with col2:
    dest_file = st.file_uploader("2. Enviar Arquivo de Destino (Cabeçalho na linha 11)", type=["xlsx", "xls"])

@st.cache_data(show_spinner=False)
def ler_destino(file):
    xls = pd.ExcelFile(file)
    all_sheets = {sheet: pd.read_excel(file, sheet_name=sheet, header=10) for sheet in xls.sheet_names}
    return xls.sheet_names, all_sheets

if source_file and dest_file:
    try:
        header_val = None if origem_sem_cabecalho else 0
        nome_arquivo = source_file.name.lower()
        
        # Leitor inteligente para Origem (suporta CSV, TXT e Excel)
        if nome_arquivo.endswith(".csv"):
            raw_df = pd.read_csv(source_file, header=header_val)
        elif nome_arquivo.endswith(".txt"):
            try:
                raw_df = pd.read_csv(source_file, header=header_val, sep=None, engine='python')
            except:
                raw_df = pd.read_csv(source_file, header=header_val, sep='\t')
        else:
            raw_df = pd.read_excel(source_file, header=header_val)
            
        # Se não tiver cabeçalho, renomeia as colunas automaticamente
        if origem_sem_cabecalho or raw_df.columns.dtype != 'O':
            raw_df.columns = [f"Coluna {i+1}" for i in range(len(raw_df.columns))]

        # Gerencia estado da sessão para manter as seleções
        if "file_name" not in st.session_state or st.session_state["file_name"] != source_file.name or st.session_state.get("sem_cabecalho_antigo") != origem_sem_cabecalho:
            raw_df.insert(0, "Selecionar", False)
            st.session_state["source_df"] = raw_df
            st.session_state["file_name"] = source_file.name
            st.session_state["sem_cabecalho_antigo"] = origem_sem_cabecalho

        sheet_names, all_dest_dfs = ler_destino(dest_file)
        
        st.success("⚡ Arquivos carregados com sucesso!")
        
        # --- Seleção da Aba ---
        selected_sheet = st.selectbox("3. Escolha a Aba (Planilha) de Destino:", sheet_names)
        dest_df = all_dest_dfs[selected_sheet]
        
        # --- Mapeamento Inteligente (De/Para) ---
        st.subheader("4. Mapeamento de Colunas (Destino x Origem)")
        mapping = {}
        source_cols = [c for c in st.session_state["source_df"].columns if c != "Selecionar"]
        
        for dest_col in dest_df.columns:
            if "Unnamed" in str(dest_col): continue
            
            default_idx = 0
            for i, sc in enumerate(source_cols):
                if str(dest_col).strip().upper() in str(sc).strip().upper():
                    default_idx = i
                    break
            
            mapping[dest_col] = st.selectbox(
                f"Destino '{dest_col}' recebe da Origem:", 
                options=source_cols, 
                index=default_idx,
                key=f"map_{dest_col}"
            )
            
        # --- Pesquisa Direta e Precisa ---
        st.subheader("5. Pesquisa Direta na Origem")
        
        col_busca_1, col_busca_2 = st.columns([1, 2])
        with col_busca_1:
            coluna_pesquisa = st.selectbox("Pesquisar na coluna:", options=source_cols)
        with col_busca_2:
            search = st.text_input(f"🔍 Digite o termo para buscar em '{coluna_pesquisa}':", "")
        
        df_atual = st.session_state["source_df"]
        
        if search:
            mask = df_atual[coluna_pesquisa].astype(str).str.contains(search, case=False, na=False)
            df_display = df_atual[mask | (df_atual["Selecionar"] == True)]
        else:
            df_display = df_atual
            
        edited_df = st.data_editor(df_display, use_container_width=True, hide_index=True, key="editor_origem")
        
        for idx in edited_df.index:
            st.session_state["source_df"].loc[idx, "Selecionar"] = edited_df.loc[idx, "Selecionar"]
            
        final_selected = st.session_state["source_df"][st.session_state["source_df"]["Selecionar"] == True]
        st.write(f"📌 Total de registros marcados para envio: **{len(final_selected)}**")
        
        # --- Inserção ---
        st.subheader("6. Finalizar e Inserir no Destino")
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
        st.error(f"Ocorreu um erro ao processar o arquivo de Origem: {e}")
