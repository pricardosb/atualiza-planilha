import io
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from copy import copy

st.set_page_config(page_title="Integrador Profissional", layout="wide")
st.title("⚡ Integrador: Leitura Automática de Qualquer Formato")

# --- 1. CARREGAMENTO DOS ARQUIVOS ---
col1, col2 = st.columns(2)
with col1:
    source_file = st.file_uploader("1. Arquivo de ORIGEM (Excel, CSV ou TXT)", type=["xlsx", "xls", "csv", "txt"])
    origem_tem_cabecalho = st.checkbox("Origem tem cabeçalho?", value=True)
with col2:
    dest_file = st.file_uploader("2. Arquivo de DESTINO (.xlsx)", type=["xlsx"])
    header_dest = st.number_input("Linha do cabeçalho no Destino:", value=11, min_value=1)

if source_file and dest_file:
    # --- LEITURA DA ORIGEM INTELIGENTE (DETECTA QUALQUER FORMATO) ---
    cache_key_src = f"{source_file.name}_{origem_tem_cabecalho}"
    if "source_df" not in st.session_state or st.session_state.get("last_cache_key_src") != cache_key_src:
        hdr = 0 if origem_tem_cabecalho else None
        raw = None
        
        try:
            # Tenta ler como Excel primeiro (.xlsx, .xls, .xlsm)
            source_file.seek(0)
            raw = pd.read_excel(source_file, header=hdr)
        except Exception:
            # Se falhar, é porque é um arquivo de texto/CSV/TXT. Vamos testar as codificações (UTF-8, Latin1, etc.)
            encodings = ['utf-8', 'latin1', 'cp1252', 'iso-8859-1']
            file_name_lower = source_file.name.lower()
            
            for enc in encodings:
                try:
                    source_file.seek(0)
                    if file_name_lower.endswith('.csv'):
                        raw = pd.read_csv(source_file, header=hdr, encoding=enc)
                    else:
                        # Para .txt ou arquivos com delimitadores variados
                        raw = pd.read_csv(source_file, sep=None, engine='python', header=hdr, encoding=enc)
                    break
                except Exception:
                    continue
                    
        if raw is None:
            st.error("❌ Não foi possível ler o arquivo de origem. Verifique se o arquivo está corrompido.")
            st.stop()
            
        if not origem_tem_cabecalho:
            raw.columns = [f"Col {i+1}" for i in range(len(raw.columns))]
            
        st.session_state["source_df"] = raw
        st.session_state["last_cache_key_src"] = cache_key_src
        
    df_origem = st.session_state["source_df"]

    # --- 2. SELEÇÃO DOS DADOS COM CONTADOR ---
    st.subheader("3. Seleção de Dados da Origem")
    col_busca = st.selectbox("Coluna para identificar os registros:", df_origem.columns)
    
    opcoes = [f"{val} (Linha {idx})" for idx, val in df_origem[col_busca].items()]
    selected_options = st.multiselect("🔍 Escolha os registros que deseja enviar:", opcoes)
    
    selected_indices = [int(item.split("(Linha ")[1].replace(")", "")) for item in selected_options]
    st.metric("Total de itens selecionados", len(selected_indices))

    # --- LEITURA DO DESTINO (COM CACHE) ---
    if "dest_wb" not in st.session_state or st.session_state.get("last_cache_key_dest") != dest_file.name:
        dest_file.seek(0)
        st.session_state["dest_wb"] = load_workbook(io.BytesIO(dest_file.getvalue()), data_only=True)
        st.session_state["last_cache_key_dest"] = dest_file.name
    wb = st.session_state["dest_wb"]
    
    target_sheet = st.selectbox("4. Escolha a ABA de Destino:", wb.sheetnames)
    ws = wb[target_sheet]

    # --- 3. MAPEAMENTO POR COLUNA ---
    st.subheader("5. Mapeamento de Colunas")
    max_col = ws.max_column
    mapping = {}
    
    cols_ui = st.columns(4)
    for i in range(1, max_col + 1):
        header_val = ws.cell(row=header_dest, column=i).value
        label = f"Col {i} ({header_val})" if header_val else f"Col {i} (Em branco)"
        
        with cols_ui[(i-1) % 4]:
            options = ["--- Não mapear ---", "⚠️ Auto-incrementar (Seq)"] + list(df_origem.columns)
            map_val = st.selectbox(label, options, key=f"map_{i}")
            if map_val != "--- Não mapear ---":
                mapping[i] = map_val

    # --- 4. OPÇÕES DE LOCAL DE INSERÇÃO ---
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
            st.error("⚠️ Selecione pelo menos um registro na lista acima!")
        else:
            try:
                wb_write = load_workbook(io.BytesIO(dest_file.getvalue()))
                ws_write = wb_write[target_sheet]
                num_new = len(selected_indices)
                
                if modo_insercao == "Final da planilha":
                    start_row = ws_write.max_row + 1
                    try:
                        base_seq = int(ws_write.cell(row=ws_write.max_row, column=1).value or 0)
                    except:
                        base_seq = 0
                else:
                    start_row = target_row
                    ws_write.insert_rows(start_row, amount=num_new)
                    try:
                        base_seq = int(ws_write.cell(row=start_row - 1, column=1).value or 0)
                    except:
                        base_seq = 0

                ref_row_idx = start_row - 1 if start_row > header_dest + 1 else header_dest + 1
                if ref_row_idx < header_dest + 1:
                    ref_row_idx = header_dest + 1

                current_row = start_row
                seq_val = base_seq

                for idx in selected_indices:
                    seq_val += 1
                    for col_idx, origem_col in mapping.items():
                        if origem_col == "⚠️ Auto-incrementar (Seq)":
                            cell_val = seq_val
                        else:
                            cell_val = df_origem.iloc[idx][origem_col]
                        
                        target_cell = ws_write.cell(row=current_row, column=col_idx, value=cell_val)
                        
                        ref_cell = ws_write.cell(row=ref_row_idx, column=col_idx)
                        if ref_cell.font: target_cell.font = copy(ref_cell.font)
                        if ref_cell.border: target_cell.border = copy(ref_cell.border)
                        if ref_cell.fill: target_cell.fill = copy(ref_cell.fill)
                        if ref_cell.alignment: target_cell.alignment = copy(ref_cell.alignment)
                        if ref_cell.number_format: target_cell.number_format = ref_cell.number_format
                            
                    current_row += 1

                if modo_insercao == "A partir de uma linha específica":
                    next_seq = seq_val
                    for r in range(start_row + num_new, ws_write.max_row + 1):
                        next_seq += 1
                        ws_write.cell(row=r, column=1, value=next_seq)

                buffer = io.BytesIO()
                wb_write.save(buffer)
                buffer.seek(0)
                
                st.success(f"✅ Processamento concluído com sucesso! {num_new} registros adicionados.")
                st.download_button(
                    "📥 Baixar Arquivo Atualizado", 
                    data=buffer.getvalue(), 
                    file_name="destino_atualizado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Erro ao processar: {e}")
