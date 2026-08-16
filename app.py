import io
import pandas as pd
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(page_title="Integrador Profissional", layout="wide")
st.title("⚡ Integrador: Edição Direta no Destino")

# --- 1. CARREGAMENTO ---
col1, col2 = st.columns(2)
with col1:
    source_file = st.file_uploader("1. Arquivo de ORIGEM", type=["xlsx", "xls", "csv", "txt"])
with col2:
    dest_file = st.file_uploader("2. Arquivo de DESTINO (Original)", type=["xlsx", "xls"])
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

    # --- 2. PESQUISA EM ORDEM ---
    st.subheader("3. Seleção de Registros")
    col_busca = st.selectbox("Buscar na coluna da Origem:", df_origem.columns)
    
    # Gerar opções e ORDENAR
    opcoes = [f"{val} (Linha {idx})" for idx, val in df_origem[col_busca].items()]
    opcoes.sort() # <--- ORDENAÇÃO AQUI
    
    selected_options = st.multiselect("🔍 Digite e selecione (Ordem alfabética):", opcoes)
    
    selected_indices = [int(item.split("(Linha ")[1].replace(")", "")) for item in selected_options]
    st.write(f"📌 **Total selecionados: {len(selected_indices)}**")

    # --- 3. MAPEAMENTO ---
    st.subheader("4. Mapeamento")
    # Ler cabeçalho do destino para mostrar as colunas
    dest_df_header = pd.read_excel(dest_file, header=header_dest-1, nrows=0)
    dest_cols = [c for c in dest_df_header.columns if "Unnamed" not in str(c)]
    
    mapping = {}
    cols_ui = st.columns(3)
    for i, d_col in enumerate(dest_cols):
        with cols_ui[i % 3]:
            map_val = st.selectbox(f"Destino '{d_col}' recebe de:", ["--- Não mapear ---"] + list(df_origem.columns), key=f"map_{d_col}")
            if map_val != "--- Não mapear ---":
                mapping[d_col] = map_val

    # --- 4. INSERÇÃO NO DESTINO (USANDO OPENPYXL) ---
    st.subheader("5. Atualização")
    modo = st.radio("Onde inserir?", ["Final da planilha", "Em uma linha específica"])
    target_row = st.number_input("Linha (se escolhido específica):", min_value=1, value=1)

    if st.button("🚀 Atualizar Arquivo Original"):
        if not selected_indices:
            st.error("Selecione registros!")
        else:
            # Prepara dados
            dados_para_inserir = df_origem.iloc[selected_indices][list(mapping.values())]
            
            # Carrega o arquivo original usando openpyxl (Preserva tudo)
            wb = load_workbook(dest_file)
            sheet_names = wb.sheetnames
            selected_sheet = st.selectbox("Selecione a aba para inserir:", sheet_names)
            ws = wb[selected_sheet]

            # Inserir dados
            if modo == "Final da planilha":
                for row_data in dados_para_inserir.itertuples(index=False):
                    ws.append(list(row_data))
            else:
                # Insere linhas no meio (move as de baixo para baixo)
                ws.insert_rows(target_row, amount=len(dados_para_inserir))
                for i, row_data in enumerate(dados_para_inserir.itertuples(index=False)):
                    for col_idx, value in enumerate(row_data, start=1):
                        ws.cell(row=target_row + i, column=col_idx, value=value)

            # Salvar
            buffer = io.BytesIO()
            wb.save(buffer)
            
            st.success("Atualização realizada com sucesso!")
            st.download_button("📥 Baixar Arquivo Atualizado", data=buffer.getvalue(), file_name="destino_atualizado.xlsx")
