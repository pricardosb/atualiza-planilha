import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Integrador XLS Profissional", layout="wide")
st.title("⚡ Integrador XLS: Preservando Títulos e Inserção Precisa")

# --- 1. CARREGAMENTO DOS ARQUIVOS ---
col1, col2 = st.columns(2)
with col1:
    source_file = st.file_uploader("1. Arquivo de ORIGEM (Excel, CSV ou TXT)", type=["xlsx", "xls", "csv", "txt"])
with col2:
    dest_file = st.file_uploader("2. Arquivo de DESTINO (.xls)", type=["xls", "xlsx"])
    header_dest = st.number_input("Linha do cabeçalho no Destino:", value=11, min_value=1)

if source_file and dest_file:
    # --- LEITURA DA ORIGEM ---
    if "source_df" not in st.session_state or st.session_state.get("last_source") != source_file.name:
        try:
            if source_file.name.endswith(".csv"): raw = pd.read_csv(source_file)
            elif source_file.name.endswith(".txt"): raw = pd.read_csv(source_file, sep=None, engine='python')
            else: raw = pd.read_excel(source_file)
            st.session_state["source_df"] = raw
            st.session_state["last_source"] = source_file.name
        except Exception as e:
            st.error(f"Erro ao ler a origem: {e}")
            st.stop()
    
    df_origem = st.session_state["source_df"]

    # --- LEITURA DO DESTINO (PRESERVANDO TUDO ACIMA DO CABEÇALHO) ---
    try:
        # Lê o arquivo mantendo todas as linhas (sem descartar o topo)
        all_sheets_raw = pd.read_excel(dest_file, sheet_name=None, header=None)
        sheet_names = list(all_sheets_raw.keys())
        target_sheet = st.selectbox("3. Escolha a ABA (Pasta) de Destino:", sheet_names)
        
        sheet_df = all_sheets_raw[target_sheet]
    except Exception as e:
        st.error(f"Erro ao ler o arquivo de destino: {e}")
        st.stop()

    # --- 2. PESQUISA ORDENADA COM CONTAGEM ---
    st.subheader("4. Seleção de Registros da Origem")
    col_busca = st.selectbox("Escolha a coluna para pesquisar na Origem:", df_origem.columns)
    
    opcoes = [f"{val} (Linha {idx})" for idx, val in df_origem[col_busca].items()]
    opcoes.sort() 
    
    selected_options = st.multiselect("🔍 Digite e selecione os registros:", opcoes)
    selected_indices = [int(item.split("(Linha ")[1].replace(")", "")) for item in selected_options]
    
    st.markdown(f"### 📌 Total de registros selecionados: **{len(selected_indices)}**")

    # --- 3. MAPEAMENTO MANUAL ---
    st.subheader("5. Mapeamento de Colunas (Origem x Destino)")
    
    header_row_idx = header_dest - 1
    header_values = sheet_df.iloc[header_row_idx].tolist()
    
    dest_cols_map = {str(val): i for i, val in enumerate(header_values) if pd.notna(val) and "Unnamed" not in str(val)}
    dest_cols = list(dest_cols_map.keys())
    
    mapping = {}
    cols_ui = st.columns(3)
    for i, d_col in enumerate(dest_cols):
        with cols_ui[i % 3]:
            map_val = st.selectbox(f"Destino '{d_col}' recebe de:", ["--- Não mapear ---"] + list(df_origem.columns), key=f"map_{i}")
            if map_val != "--- Não mapear ---":
                mapping[d_col] = map_val

    # --- 4. OPÇÕES DE INSERÇÃO PRECISA ---
    st.subheader("6. Onde deseja inserir os dados na aba?")
    modo_insercao = st.radio("Escolha o local de inserção:", ["Final da planilha", "A partir de uma linha específica"])
    
    min_linha_val = header_dest + 1
    target_row = min_linha_val
    if modo_insercao == "A partir de uma linha específica":
        target_row = st.number_input(f"Digite o número da linha exata (Mínimo {min_linha_val}):", min_value=min_linha_val, value=min_linha_val)

    # --- 5. PROCESSAMENTO E GERAÇÃO DO ARQUIVO .XLS ---
    st.subheader("7. Finalizar Processo")
    if st.button("🚀 Processar e Gerar Arquivo XLS"):
        if not selected_indices:
            st.error("⚠️ Você precisa selecionar pelo menos um registro na pesquisa acima!")
        elif not mapping:
            st.error("⚠️ Faça pelo menos um mapeamento de colunas!")
        else:
            try:
                # Separa as partes da planilha para não apagar o topo
                top_rows = sheet_df.iloc[:header_row_idx].copy()      # Linhas acima do cabeçalho (títulos intactos)
                header_df = sheet_df.iloc[[header_row_idx]].copy()    # Linha do cabeçalho
                data_rows = sheet_df.iloc[header_row_idx + 1:].copy()  # Dados existentes
                
                # Prepara os novos dados mapeados
                dados_selecionados = df_origem.iloc[selected_indices]
                new_rows_df = pd.DataFrame(columns=sheet_df.columns)
                
                for _, src_row in dados_selecionados.iterrows():
                    new_row_data = {}
                    for dest_col, orig_col in mapping.items():
                        col_idx = dest_cols_map[dest_col]
                        new_row_data[col_idx] = src_row[orig_col]
                    new_rows_df = pd.concat([new_rows_df, pd.DataFrame([new_row_data])], ignore_index=True)
                
                # Insere no local exato escolhido
                if modo_insercao == "Final da planilha":
                    final_data = pd.concat([data_rows, new_rows_df], ignore_index=True)
                else:
                    rel_idx = target_row - (header_row_idx + 2)
                    rel_idx = max(0, min(rel_idx, len(data_rows)))
                    
                    part1 = data_rows.iloc[:rel_idx]
                    part2 = data_rows.iloc[rel_idx:]
                    final_data = pd.concat([part1, new_rows_df, part2], ignore_index=True)
                
                # Reconstrói a aba mantendo o topo e o cabeçalho originais
                updated_sheet_df = pd.concat([top_rows, header_df, final_data], ignore_index=True)
                all_sheets_raw[target_sheet] = updated_sheet_df
                
                # Salva todas as abas em formato .xls
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlwt') as writer:
                    for s_name, s_df in all_sheets_raw.items():
                        s_df.to_excel(writer, sheet_name=s_name, index=False, header=False)
                
                st.success("✅ Arquivo XLS processado com sucesso mantendo os títulos!")
                st.download_button(
                    "📥 Baixar Arquivo Atualizado (.xls)", 
                    data=buffer.getvalue(), 
                    file_name="destino_atualizado.xls",
                    mime="application/vnd.ms-excel"
                )
            except Exception as e:
                st.error(f"Erro ao processar os dados: {e}")
