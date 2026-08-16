import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Integrador Profissional", layout="wide")
st.title("⚡ Integrador de Dados: Origem -> Destino")

# --- 1. CARREGAMENTO DOS ARQUIVOS ---
col1, col2 = st.columns(2)
with col1:
    source_file = st.file_uploader("1. Arquivo de ORIGEM (Excel, CSV ou TXT)", type=["xlsx", "xls", "csv", "txt"])
with col2:
    dest_file = st.file_uploader("2. Arquivo de DESTINO (.xls ou .xlsx)", type=["xlsx", "xls"])
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

    # --- LEITURA DO DESTINO (TODAS AS ABAS) ---
    try:
        all_sheets = pd.read_excel(dest_file, sheet_name=None, header=header_dest-1)
        target_sheet = st.selectbox("3. Escolha a ABA (Pasta) onde inserir os dados:", list(all_sheets.keys()))
        dest_df = all_sheets[target_sheet]
    except Exception as e:
        st.error(f"Erro ao ler o arquivo de Destino: {e}")
        st.stop()

    # --- 2. PESQUISA ORDENADA COM CONTAGEM ---
    st.subheader("4. Seleção de Registros da Origem")
    col_busca = st.selectbox("Escolha a coluna para pesquisar na Origem:", df_origem.columns)
    
    # Gera opções e ORDENA em ordem alfabética
    opcoes = [f"{val} (Linha {idx})" for idx, val in df_origem[col_busca].items()]
    opcoes.sort() 
    
    selected_options = st.multiselect("🔍 Digite e selecione os registros:", opcoes)
    selected_indices = [int(item.split("(Linha ")[1].replace(")", "")) for item in selected_options]
    
    # CONTAGEM EM TEMPO REAL
    st.markdown(f"### 📌 Total de registros selecionados: **{len(selected_indices)}**")

    # --- 3. MAPEAMENTO MANUAL ---
    st.subheader("5. Mapeamento de Colunas (Origem x Destino)")
    dest_cols = [c for c in dest_df.columns if "Unnamed" not in str(c)]
    
    mapping = {}
    cols_ui = st.columns(3)
    for i, d_col in enumerate(dest_cols):
        with cols_ui[i % 3]:
            map_val = st.selectbox(f"Destino '{d_col}' recebe de:", ["--- Não mapear ---"] + list(df_origem.columns), key=f"map_{i}")
            if map_val != "--- Não mapear ---":
                mapping[d_col] = map_val

    # --- 4. OPÇÕES DE INSERÇÃO (FINAL OU LINHA ESPECÍFICA) ---
    st.subheader("6. Onde deseja inserir os dados na aba?")
    modo_insercao = st.radio("Escolha o local de inserção:", ["Final da planilha", "A partir de uma linha específica"])
    
    linha_especifica = 0
    if modo_insercao == "A partir de uma linha específica":
        linha_especifica = st.number_input("Digite o número da linha de destino (após o cabeçalho):", min_value=0, max_value=len(dest_df), value=0)

    # --- 5. PROCESSAMENTO E GERAÇÃO DO ARQUIVO ---
    st.subheader("7. Finalizar Processo")
    if st.button("🚀 Processar e Atualizar Destino"):
        if not selected_indices:
            st.error("⚠️ Você precisa selecionar pelo menos um registro na pesquisa acima!")
        elif not mapping:
            st.error("⚠️ Faça pelo menos um mapeamento de colunas!")
        else:
            try:
                # Prepara os dados selecionados com base no mapeamento
                dados_para_inserir = df_origem.iloc[selected_indices][list(mapping.values())].copy()
                dados_para_inserir.columns = list(mapping.keys())
                
                # Insere de acordo com a escolha do usuário
                if modo_insercao == "Final da planilha":
                    updated_sheet = pd.concat([dest_df, dados_para_inserir], ignore_index=True)
                else:
                    updated_sheet = pd.concat([dest_df.iloc[:linha_especifica], dados_para_inserir, dest_df.iloc[linha_especifica:]], ignore_index=True)
                
                # Atualiza a aba específica no dicionário mantendo as outras intactas
                all_sheets[target_sheet] = updated_sheet
                
                # Salva o arquivo preservando todas as abas originais
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    for sheet_name, df_data in all_sheets.items():
                        df_data.to_excel(writer, sheet_name=sheet_name, index=False, startrow=header_dest-1)
                
                st.success("✅ Arquivo atualizado com sucesso!")
                st.download_button(
                    "📥 Baixar Arquivo Atualizado", 
                    data=buffer.getvalue(), 
                    file_name="destino_atualizado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Erro ao processar os dados: {e}")
