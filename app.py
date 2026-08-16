import io
import pandas as pd
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(page_title="Integrador Profissional", layout="wide")
st.title("⚡ Integrador: Origem -> Destino")

# --- 1. CARREGAMENTO DOS ARQUIVOS ---
col1, col2 = st.columns(2)
with col1:
    source_file = st.file_uploader("1. Arquivo de ORIGEM (Fonte)", type=["xlsx", "xls", "csv", "txt"])
with col2:
    dest_file = st.file_uploader("2. Arquivo de DESTINO (Onde inserir)", type=["xlsx", "xls"])
    header_dest = st.number_input("Em qual linha está o cabeçalho no Destino?", value=11, min_value=1)

if source_file and dest_file:
    # --- LEITURA DA ORIGEM ---
    if "source_df" not in st.session_state or st.session_state.get("last_source") != source_file.name:
        if source_file.name.endswith(".csv"): raw = pd.read_csv(source_file)
        elif source_file.name.endswith(".txt"): raw = pd.read_csv(source_file, sep=None, engine='python')
        else: raw = pd.read_excel(source_file)
        st.session_state["source_df"] = raw
        st.session_state["last_source"] = source_file.name

    df_origem = st.session_state["source_df"]

    # --- 2. PESQUISA DINÂMICA (Multiselect) ---
    st.subheader("3. Seleção de Registros")
    
    col_busca = st.selectbox("Escolha qual coluna usar para pesquisar na Origem:", df_origem.columns)
    
    # Cria uma lista de opções para o Multiselect (Ex: "João (Linha 5)", "Maria (Linha 6)")
    # Isso resolve a "caixa de rolagem" e "selecionar na própria caixa"
    opcoes = [f"{val} (Linha {idx})" for idx, val in df_origem[col_busca].items()]
    
    selected_options = st.multiselect("🔍 Digite o nome/termo e selecione abaixo:", opcoes)
    
    # Extrai os índices das linhas selecionadas
    selected_indices = [int(item.split("(Linha ")[1].replace(")", "")) for item in selected_options]
    st.write(f"📌 **Total de registros selecionados: {len(selected_indices)}**")

    if selected_indices:
        st.write("Registros selecionados:")
        st.dataframe(df_origem.iloc[selected_indices], use_container_width=True)

    # --- 3. MAPEAMENTO MANUAL ---
    st.subheader("4. Mapeamento de Colunas (Origem x Destino)")
    # Lê colunas do destino para saber o que mapear
    temp_dest = pd.read_excel(dest_file, header=header_dest-1, nrows=0)
    dest_cols = [c for c in temp_dest.columns if "Unnamed" not in str(c)]
    
    mapping = {}
    cols_ui = st.columns(3)
    for i, d_col in enumerate(dest_cols):
        with cols_ui[i % 3]:
            # Default para '--- Não mapear ---'
            map_val = st.selectbox(f"Destino '{d_col}' recebe de:", ["--- Não mapear ---"] + list(df_origem.columns), key=f"map_{d_col}")
            if map_val != "--- Não mapear ---":
                mapping[d_col] = map_val

    # --- 4. INSERÇÃO ---
    st.subheader("5. Inserir no Destino")
    modo = st.radio("Onde inserir?", ["Final da planilha", "Em uma linha específica"])
    target_row = 0
    if modo == "Em uma linha específica":
        target_row = st.number_input("Inserir na linha:", min_value=1)

    if st.button("🚀 Inserir Dados no Destino"):
        if not selected_indices:
            st.error("Selecione pelo menos um registro!")
        elif not mapping:
            st.error("Faça pelo menos um mapeamento!")
        else:
            # Prepara os dados selecionados
            dados_para_inserir = df_origem.iloc[selected_indices].copy()
            # Renomeia para bater com o destino
            dados_para_inserir = dados_para_inserir[list(mapping.values())]
            dados_para_inserir.columns = list(mapping.keys())
            
            # Lê o arquivo destino completo
            dest_df = pd.read_excel(dest_file, header=header_dest-1)
            
            # Concatenação lógica
            if modo == "Final da planilha":
                updated_df = pd.concat([dest_df, dados_para_inserir], ignore_index=True)
            else:
                updated_df = pd.concat([dest_df.iloc[:target_row], dados_para_inserir, dest_df.iloc[target_row:]], ignore_index=True)
            
            # Salva o arquivo (o to_excel com startrow=header_dest mantém a formatação acima dele)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                # Carrega o arquivo destino para preservar estilos, se possível
                updated_df.to_excel(writer, index=False, startrow=header_dest)
            
            st.success("Dados inseridos com sucesso!")
            st.download_button("📥 Baixar Arquivo Atualizado", data=buffer.getvalue(), file_name="destino_processado.xlsx")
