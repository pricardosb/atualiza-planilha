import io
import pandas as pd
import numpy as np
import streamlit as st
from openpyxl import load_workbook
from copy import copy

# --- FUNÇÕES DE SUPORTE ---
def copiar_estilo_completo(origem, destino):
    if origem.has_style:
        destino.font = copy(origem.font)
        destino.border = copy(origem.border)
        destino.fill = copy(origem.fill)
        destino.number_format = copy(origem.number_format)
        destino.protection = copy(origem.protection)
        destino.alignment = copy(origem.alignment)

def deduplicar_colunas(colunas):
    vistos = {}
    novas_colunas = []
    for col in colunas:
        col_str = str(col).strip()
        if col_str in vistos:
            vistos[col_str] += 1
            novas_colunas.append(f"{col_str} ({vistos[col_str]})")
        else:
            vistos[col_str] = 1
            novas_colunas.append(col_str)
    return novas_colunas

def extrair_valor_limpo(df, idx, col_name):
    try:
        val = df.iloc[idx][col_name]
        if isinstance(val, pd.Series):
            val = val.iloc[0]
        if pd.isna(val):
            return None
        if hasattr(val, 'item'):
            val = val.item()
        return val
    except Exception:
        return None

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Integrador Profissional", layout="wide")
st.title("⚡ Integrador Profissional")

# --- CARREGAMENTO ---
col1, col2 = st.columns(2)
with col1:
    st.subheader("1. Arquivo de ORIGEM")
    source_file = st.file_uploader("Selecione o arquivo de ORIGEM", type=["xlsx", "xls", "csv", "txt"], key="src_upload")
    origem_tem_cabecalho = st.checkbox("Origem tem cabeçalho na 1ª linha?", value=True)
with col2:
    st.subheader("2. Arquivo de DESTINO")
    dest_file = st.file_uploader("Selecione o arquivo de DESTINO (.xlsx)", type=["xlsx"], key="dest_upload")
    header_dest = st.number_input("Linha do cabeçalho no Destino:", value=11, min_value=1)

if source_file:
    cache_key_src = f"{source_file.name}_{origem_tem_cabecalho}"
    if "source_df" not in st.session_state or st.session_state.get("last_cache_key_src") != cache_key_src:
        hdr = 0 if origem_tem_cabecalho else None
        raw = None
        try:
            source_file.seek(0)
            raw = pd.read_excel(source_file, header=hdr)
        except:
            for enc in ['utf-8', 'latin1', 'cp1252', 'iso-8859-1']:
                try:
                    source_file.seek(0)
                    raw = pd.read_csv(source_file, sep=None, engine='python', header=hdr, encoding=enc)
                    break
                except: continue
        if raw is not None:
            if not origem_tem_cabecalho: 
                raw.columns = [f"Col {i+1}" for i in range(len(raw.columns))]
            else: 
                raw.columns = deduplicar_colunas(raw.columns)
            st.session_state["source_df"] = raw
            st.session_state["last_cache_key_src"] = cache_key_src

if dest_file:
    if "wb_data" not in st.session_state or st.session_state.get("last_dest_name") != dest_file.name:
        dest_file.seek(0)
        st.session_state["wb_data"] = dest_file.getvalue()
        st.session_state["last_dest_name"] = dest_file.name

# --- FLUXO PRINCIPAL ---
df_origem = st.session_state.get("source_df")
wb_data = st.session_state.get("wb_data")

if df_origem is not None and wb_data is not None:
    wb = load_workbook(io.BytesIO(wb_data))
    
    target_sheet = st.selectbox("Escolha a ABA na Planilha de Destino a ser Atualizada:", wb.sheetnames)
    ws = wb[target_sheet]

    # --- 3. SELEÇÃO DE REGISTROS ---
    st.subheader("3. Seleção de Registros")
    col_busca = st.selectbox("Coluna identificadora (para seleção):", df_origem.columns)
    
    opcoes_selecao = [f"{val} (Linha {idx})" for idx, val in df_origem[col_busca].items()]
    selected_options = st.multiselect("🔍 Escolha os registros:", opcoes_selecao)
    selected_indices = [int(item.split("(Linha ")[1].replace(")", "")) for item in selected_options]

    if selected_indices:
        st.info(f"📊 **{len(selected_indices)}** registro(s) selecionado(s) para atualização.")

    st.write("---")

    # --- 4. CORRELAÇÃO ORIGEM X DESTINO ---
    st.subheader("4. Correlação ORIGEM X DESTINO")
    mapping = {}
    cols_ui = st.columns(4)
    opcoes_mapeamento = ["--- Não mapear ---", "⚠️ Auto-incrementar (Seq)"] + list(df_origem.columns)
    
    for i in range(1, ws.max_column + 1):
        header_val = ws.cell(row=header_dest, column=i).value
        with cols_ui[(i-1) % 4]:
            map_val = st.selectbox(f"Col {i} ({header_val or 'S/ Título'})", opcoes_mapeamento, key=f"map_{i}")
            if map_val != "--- Não mapear ---":
                mapping[i] = map_val

    st.write("---")

    # --- LOCAL DE INSERÇÃO ---
    modo_insercao = st.radio("Local de inserção:", ["Final da planilha", "A partir de uma linha específica"])
    target_row = st.number_input("Linha:", min_value=header_dest+1, value=header_dest+1) if modo_insercao == "A partir de uma linha específica" else ws.max_row + 1

    # --- PROCESSAMENTO ---
    if st.button("🚀 Processar e Atualizar"):
        if not selected_indices: st.error("Selecione itens!"); st.stop()
        
        # 1. Obter valor base de sequência da linha anterior
        ref_row_idx = (target_row - 1) if modo_insercao == "A partir de uma linha específica" else (ws.max_row)
        
        base_seq = 0
        if ref_row_idx >= header_dest:
            val_acima = ws.cell(row=ref_row_idx, column=1).value
            try:
                base_seq = int(val_acima)
            except:
                base_seq = 0
        
        # 2. Inserir linhas caso seja no meio
        if modo_insercao == "A partir de uma linha específica":
            ws.insert_rows(target_row, amount=len(selected_indices))
        
        # 3. Escrever dados e numeração
        current_row = target_row
        seq_val = base_seq

        for idx in selected_indices:
            seq_val += 1
            for col_idx in range(1, ws.max_column + 1):
                target_cell = ws.cell(row=current_row, column=col_idx)
                ref_cell = ws.cell(row=ref_row_idx, column=col_idx)
                
                # Copia formatação integral da célula de referência
                copiar_estilo_completo(ref_cell, target_cell)
                
                # Gravação robusta
                if col_idx == 1 or mapping.get(col_idx) == "⚠️ Auto-incrementar (Seq)":
                    target_cell.value = seq_val
                elif col_idx in mapping:
                    origem_col = mapping[col_idx]
                    target_cell.value = extrair_valor_limpo(df_origem, idx, origem_col)
                else:
                    target_cell.value = None # Limpa colunas não mapeadas
            
            current_row += 1

        # 4. Re-sequenciar linhas abaixo (caso inserido no meio)
        if modo_insercao == "A partir de uma linha específica":
            for r in range(current_row, ws.max_row + 1):
                val_atual = ws.cell(row=r, column=1).value
                if val_atual is not None:
                    seq_val += 1
                    ws.cell(row=r, column=1, value=seq_val)

        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        st.success("✅ Processamento concluído com sucesso!")
        st.download_button("📥 Baixar Arquivo Atualizado", buffer.getvalue(), "destino_atualizado.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
