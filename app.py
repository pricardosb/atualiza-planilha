import io
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from copy import copy

st.set_page_config(page_title="Integrador Profissional", layout="wide")
st.title("⚡ Integrador: Ordem Automática na 1ª Coluna e Herança de Estilos")

# --- 1. CARREGAMENTO DOS ARQUIVOS ---
col1, col2 = st.columns(2)
with col1:
    source_file = st.file_uploader("1. Arquivo de ORIGEM (Excel, CSV ou TXT)", type=["xlsx", "xls", "csv", "txt"])
    origem_tem_cabecalho = st.checkbox("O arquivo de ORIGEM tem cabeçalho na 1ª linha?", value=True)
with col2:
    dest_file = st.file_uploader("2. Arquivo de DESTINO (.xlsx)", type=["xlsx"])
    header_dest = st.number_input("Linha do cabeçalho no Destino:", value=11, min_value=1)

if source_file and dest_file:
    # --- LEITURA DA ORIGEM ---
    cache_key = f"{source_file.name}_{origem_tem_cabecalho}"
    if "source_df" not in st.session_state or st.session_state.get("last_cache_key") != cache_key:
        try:
            hdr = 0 if origem_tem_cabecalho else None
            if source_file.name.endswith(".csv"): 
                raw = pd.read_csv(source_file, header=hdr)
            elif source_file.name.endswith(".txt"): 
                raw = pd.read_csv(source_file, sep=None, engine='python', header=hdr)
            else: 
                raw = pd.read_excel(source_file, header=hdr)
            
            if not origem_tem_cabecalho:
                raw.columns = [f"Coluna {i+1}" for i in range(len(raw.columns))]
                
            st.session_state["source_df"] = raw
            st.session_state["last_cache_key"] = cache_key
        except Exception as e:
            st.error(f"Erro ao ler a origem: {e}")
            st.stop()
    
    df_origem = st.session_state["source_df"]

    # --- LEITURA DO DESTINO COM OPENPYXL ---
    try:
        dest_file.seek(0)
        wb = load_workbook(io.BytesIO(dest_file.getvalue()))
        sheet_names = wb.sheetnames
        target_sheet = st.selectbox("3. Escolha a ABA (Pasta) de Destino:", sheet_names)
        ws = wb[target_sheet]
    except Exception as e:
        st.error(f"Erro ao carregar arquivo de destino: {e}")
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
    
    header_cells = ws[header_dest]
    dest_cols = [cell.value for cell in header_cells if cell.value is not None]
    col_name_to_idx = {cell.value: cell.column for cell in header_cells if cell.value is not None}
    
    mapping = {}
    cols_ui = st.columns(3)
    for i, d_col in enumerate(dest_cols):
        with cols_ui[i % 3]:
            # A primeira coluna é a Ordem, informamos o usuário visualmente no selectbox
            label_text = f"Destino '{d_col}' (Ordem):" if i == 0 else f"Destino '{d_col}' recebe de:"
            map_val = st.selectbox(label_text, ["--- Não mapear ---"] + list(df_origem.columns), key=f"map_{i}")
            if map_val != "--- Não mapear ---":
                mapping[d_col] = map_val

    # --- 4. OPÇÕES DE INSERÇÃO PRECISA ---
    st.subheader("6. Onde deseja inserir os dados na aba?")
    modo_insercao = st.radio("Escolha o local de inserção:", ["Final da planilha", "A partir de uma linha específica"])
    
    min_linha_val = header_dest + 1
    target_row = min_linha_val
    if modo_insercao == "A partir de uma linha específica":
        target_row = st.number_input(f"Digite o número da linha exata (Mínimo {min_linha_val}):", min_value=min_linha_val, value=min_linha_val)

    # --- 5. PROCESSAMENTO E GERAÇÃO DO ARQUIVO ---
    st.subheader("7. Finalizar Processo")
    if st.button("🚀 Processar e Atualizar Destino"):
        if not selected_indices:
            st.error("⚠️ Você precisa selecionar pelo menos um registro na pesquisa acima!")
        else:
            try:
                dados_para_inserir = df_origem.iloc[selected_indices]
                num_new = len(dados_para_inserir)
                
                # Define a linha inicial de inserção
                if modo_insercao == "Final da planilha":
                    start_row = ws.max_row + 1
                else:
                    start_row = target_row
                    ws.insert_rows(start_row, amount=num_new)
                
                # Linha de referência para copiar a formatação
                ref_row_idx = start_row - 1 if start_row > header_dest + 1 else header_dest + 1
                if ref_row_idx < header_dest + 1:
                    ref_row_idx = header_dest + 1

                # A primeira coluna do destino é obrigatoriamente a Ordem
                ordem_col_name = dest_cols[0]
                ordem_col_idx = col_name_to_idx[ordem_col_name]

                # Descobre o número inicial da sequência com base na linha imediatamente anterior
                base_seq = 0
                if start_row > header_dest + 1:
                    prev_val = ws.cell(row=start_row - 1, column=ordem_col_idx).value
                    try:
                        if prev_val is not None:
                            base_seq = int(prev_val)
                    except:
                        base_seq = 0

                # 1. Preenche os dados inseridos e a respectiva Ordem sequencial
                current_row = start_row
                for i, (_, row) in enumerate(dados_para_inserir.iterrows()):
                    # Preenche a 1ª coluna com a ordem sequencial correta
                    seq_val = base_seq + 1 + i
                    target_cell_ordem = ws.cell(row=current_row, column=ordem_col_idx, value=seq_val)
                    
                    # Copia formatação para a célula de ordem
                    ref_cell_ordem = ws.cell(row=ref_row_idx, column=ordem_col_idx)
                    if ref_cell_ordem.font: target_cell_ordem.font = copy(ref_cell_ordem.font)
                    if ref_cell_ordem.border: target_cell_ordem.border = copy(ref_cell_ordem.border)
                    if ref_cell_ordem.fill: target_cell_ordem.fill = copy(ref_cell_ordem.fill)
                    if ref_cell_ordem.alignment: target_cell_ordem.alignment = copy(ref_cell_ordem.alignment)
                    if ref_cell_ordem.number_format: target_cell_ordem.number_format = ref_cell_ordem.number_format

                    # Preenche as demais colunas mapeadas
                    for dest_col, orig_col in mapping.items():
                        if dest_col != ordem_col_name and dest_col in col_name_to_idx:
                            col_idx = col_name_to_idx[dest_col]
                            target_cell = ws.cell(row=current_row, column=col_idx, value=row[orig_col])
                            
                            # Copia formatação pré-existente
                            ref_cell = ws.cell(row=ref_row_idx, column=col_idx)
                            if ref_cell.font: target_cell.font = copy(ref_cell.font)
                            if ref_cell.border: target_cell.border = copy(ref_cell.border)
                            if ref_cell.fill: target_cell.fill = copy(ref_cell.fill)
                            if ref_cell.alignment: target_cell.alignment = copy(ref_cell.alignment)
                            if ref_cell.number_format: target_cell.number_format = ref_cell.number_format
                            
                    current_row += 1

                # 2. Atualiza a numeração de todas as linhas restantes abaixo para manter a sequência contínua
                next_seq = base_seq + 1 + num_new
                for r in range(start_row + num_new, ws.max_row + 1):
                    ws.cell(row=r, column=ordem_col_idx, value=next_seq)
                    next_seq += 1
                
                buffer = io.BytesIO()
                wb.save(buffer)
                buffer.seek(0)
                
                st.success("✅ Arquivo atualizado com sucesso! Ordem sequencial ajustada e estilos aplicados.")
                st.download_button(
                    "📥 Baixar Arquivo Atualizado", 
                    data=buffer.getvalue(), 
                    file_name="destino_atualizado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Erro ao processar os dados: {e}")
