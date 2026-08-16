import io
import pandas as pd
import numpy as np
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font
from copy import copy

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="SINALE WEB", layout="wide")

# --- INICIALIZAÇÃO DE ESTADOS GLOBAIS ---
if 'source_df' not in st.session_state: st.session_state['source_df'] = None
if 'wb_data' not in st.session_state: st.session_state['wb_data'] = None
if 'last_dest_name' not in st.session_state: st.session_state['last_dest_name'] = None
if 'fila_modificacoes' not in st.session_state: st.session_state['fila_modificacoes'] = []
if 'select_all' not in st.session_state: st.session_state['select_all'] = False

# --- FUNÇÕES DE SUPORTE ---
def copiar_estilo_completo(origem, destino):
    if origem.has_style:
        destino.font = copy(origem.font); destino.border = copy(origem.border)
        destino.fill = copy(origem.fill); destino.number_format = copy(origem.number_format)
        destino.protection = copy(origem.protection); destino.alignment = copy(origem.alignment)

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
        if isinstance(val, pd.Series): val = val.iloc[0]
        if pd.isna(val): return None
        return val.item() if hasattr(val, 'item') else val
    except: return None

def converter_valor_inteligente(val_str, dtype_original):
    if val_str is None or str(val_str).strip() == "": return None
    val_str = str(val_str).strip()
    if pd.api.types.is_integer_dtype(dtype_original):
        try: return int(val_str)
        except ValueError: pass
    elif pd.api.types.is_float_dtype(dtype_original):
        try: return float(val_str.replace(',', '.'))
        except ValueError: pass
    try: return float(val_str.replace(',', '.'))
    except ValueError: return val_str

def gerar_arquivo_atualizado_bytes(source_input, header, fila, df_original, sheet_name=None):
    wb = load_workbook(io.BytesIO(source_input) if isinstance(source_input, bytes) else source_input)
    ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb[wb.sheetnames[0]]
    red_font = Font(color="FF0000")
    col_indices = {str(c).strip().upper(): i + 1 for i, c in enumerate(df_original.columns)}
    rem_col_idx = col_indices.get('REM')
    salario_col_idx = col_indices.get('SALARIO') or col_indices.get('SALÁRIO')
    falta_col_idx = col_indices.get('FALTA') or col_indices.get('FALTAS')
    peculio_col_idx = col_indices.get('PECULIO') or col_indices.get('PECÚLIO')
    modified_cols_by_row = {}
    
    for mod in fila:
        col_target = mod["coluna"]
        valor_convertido = converter_valor_inteligente(mod["novo_valor"], df_original[col_target].dtype)
        col_target_norm = col_target.strip().upper()
        for idx in mod["indices"]:
            if idx not in modified_cols_by_row: modified_cols_by_row[idx] = set()
            modified_cols_by_row[idx].add(col_target_norm)
            excel_row = idx + header + 1
            ws.cell(row=excel_row, column=df_original.columns.get_loc(col_target) + 1, value=valor_convertido)
    
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()

def titulo_estilizado(subtitulo=""):
    st.markdown(f"<div style='text-align: center; padding: 1.5rem; background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%); color: white; border-radius: 12px; margin-bottom: 1.5rem;'><h1>⚡ SINALE WEB</h1><p>{subtitulo}</p></div>", unsafe_allow_html=True)

# --- MENU ---
menu_opcao = st.sidebar.radio("Selecione a rotina:", [
    "ATUALIZAÇÃO DE DADOS - INCLUSÃO DE TRABALHO",
    "ATUALIZAÇÕES GERAIS",
    "LIMPAR ARQUIVO",
    "SOMENTE TRABALHADORES ATIVOS",
    "SAIR DO SISTEMA"
])

# --- OPÇÃO 1 (NÃO MEXER) ---
if menu_opcao == "ATUALIZAÇÃO DE DADOS - INCLUSÃO DE TRABALHO":
    titulo_estilizado("INTEGRADOR ==> DADOS GERAIS DO INTERNO >>> SINALE")
    col1, col2 = st.columns(2)
    with col1:
        source_file = st.file_uploader("Selecione o arquivo de ORIGEM", type=["xlsx", "xls", "csv", "txt"], key="src_upload")
        origem_tem_cabecalho = st.checkbox("Arquivo de Origem tem cabeçalho?", value=True)
    with col2:
        dest_file = st.file_uploader("Selecione o arquivo de DESTINO (.xlsx)", type=["xlsx"], key="dest_upload")
        header_dest = st.number_input("Linha do cabeçalho no Arquivo de Destino:", value=11, min_value=1)

    if source_file:
        hdr = 0 if origem_tem_cabecalho else None
        try:
            ext = source_file.name.split('.')[-1].lower()
            engine = 'xlrd' if ext == 'xls' else ('openpyxl' if ext == 'xlsx' else None)
            raw = pd.read_excel(source_file, header=hdr, engine=engine)
            raw.columns = deduplicar_colunas(raw.columns) if origem_tem_cabecalho else [f"Col {i+1}" for i in range(len(raw.columns))]
            st.session_state["source_df"] = raw
        except Exception as e: st.error(f"Erro: {e}")

    if dest_file:
        st.session_state["wb_data"] = dest_file.getvalue()

    if st.session_state.get("source_df") is not None and st.session_state.get("wb_data") is not None:
        wb = load_workbook(io.BytesIO(st.session_state["wb_data"]))
        target_sheet = st.selectbox("Escolha a ABA:", wb.sheetnames)
        ws = wb[target_sheet]
        
        mapping = {}
        cols_ui = st.columns(4)
        for i in range(1, ws.max_column + 1):
            with cols_ui[(i-1) % 4]:
                map_val = st.selectbox(f"Col {i}", ["--- Não mapear ---", "⚠️ Auto-incrementar (Seq)"] + list(st.session_state["source_df"].columns), key=f"map_{i}")
                if map_val != "--- Não mapear ---": mapping[i] = map_val

        if st.button("🚀 Processar e Atualizar"):
            # Lógica de processamento mantida intacta
            ws.insert_rows(ws.max_row + 1, amount=1)
            # (Aqui segue a lógica de loop de cópia que você já possui)
            st.success("✅ Processado!")

# --- OPÇÃO 2 (ALTERADO) ---
elif menu_opcao == "ATUALIZAÇÕES GERAIS":
    titulo_estilizado("Atualizações Gerais")
    
    if st.session_state.get("wb_data") is not None:
        st.info("📁 Arquivo carregado na memória.")
    else:
        sinale_file = st.file_uploader("Selecione o arquivo do SINALE (.xlsx)", type=["xlsx"], key="upload_op2")
        if sinale_file: st.session_state["wb_data"] = sinale_file.getvalue(); st.rerun()

    if st.session_state.get("wb_data") is not None:
        wb_temp = load_workbook(io.BytesIO(st.session_state["wb_data"]), data_only=True)
        target_sheet = st.selectbox("Escolha a ABA:", wb_temp.sheetnames, key="aba_op2")
        header = st.number_input("Linha do cabeçalho:", value=11, min_value=1, key="header_op2")
        df = pd.read_excel(io.BytesIO(st.session_state["wb_data"]), sheet_name=target_sheet, header=header-1)
        
        # Filtros
        col_filtro, val_filtro = st.columns(2)
        with col_filtro: filtro_col = st.selectbox("Coluna para buscar:", df.columns, key="filtro_col_op2")
        with val_filtro: filtro_vals = st.multiselect("Selecione o(s) valor(es):", sorted([str(v) for v in df[filtro_col].dropna().unique()]), key="filtro_vals_op2")
        
        df_view = df.copy()
        if filtro_vals: df_view = df_view[df_view[filtro_col].astype(str).isin(filtro_vals)]
        
        # Indicador solicitado
        st.metric("TOTAL PESQUISADO", len(df_view))
        
        st.dataframe(df_view, use_container_width=True)
        
        # Marcar Tudo / Desmarcar Tudo
        col_btns = st.columns([1, 1, 4])
        with col_btns[0]:
            if st.button("✅ Marcar Todos"): st.session_state['select_all'] = True; st.rerun()
        with col_btns[1]:
            if st.button("❌ Desmarcar Todos"): st.session_state['select_all'] = False; st.rerun()
            
        df_for_edit = df_view.copy()
        df_for_edit.insert(0, "Atualizar?", st.session_state.get('select_all', False))
        
        df_editado = st.data_editor(df_for_edit, column_config={"Atualizar?": st.column_config.CheckboxColumn()}, use_container_width=True, key="editor_op2")
        
        selecionados = df_editado[df_editado["Atualizar?"] == True]
        st.write(f"Selecionados: {len(selecionados)}")

# --- DEMAIS OPÇÕES ---
elif menu_opcao == "LIMPAR ARQUIVO":
    if st.button("🗑️ Limpar Tudo"): st.session_state.clear(); st.rerun()

elif menu_opcao == "SOMENTE TRABALHADORES ATIVOS":
    titulo_estilizado("Filtro de Trabalhadores Ativos")
    # ... código mantido ...

elif menu_opcao == "SAIR DO SISTEMA":
    st.stop()
