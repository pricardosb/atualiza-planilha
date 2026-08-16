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
    for mod in fila:
        col_target = mod["coluna"]
        valor_convertido = converter_valor_inteligente(mod["novo_valor"], df_original[col_target].dtype)
        for idx in mod["indices"]:
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

# --- OPÇÃO 1: MANTIDA 100% IGUAL AO ORIGINAL ---
if menu_opcao == "ATUALIZAÇÃO DE DADOS - INCLUSÃO DE TRABALHO":
    titulo_estilizado("INTEGRADOR ==> DADOS GERAIS DO INTERNO >>> SINALE")
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("1. Arquivo de ORIGEM")
        source_file = st.file_uploader("Selecione o arquivo de ORIGEM", type=["xlsx", "xls", "csv", "txt"], key="src_upload")
        origem_tem_cabecalho = st.checkbox("Arquivo de Origem tem cabeçalho?", value=True)
    with col2:
        st.subheader("2. Arquivo de DESTINO")
        dest_file = st.file_uploader("Selecione o arquivo de DESTINO (.xlsx)", type=["xlsx"], key="dest_upload")
        header_dest = st.number_input("Linha do cabeçalho no Arquivo de Destino:", value=11, min_value=1)

    if source_file:
        cache_key_src = f"{source_file.name}_{origem_tem_cabecalho}"
        if "source_df" not in st.session_state or st.session_state.get("last_cache_key_src") != cache_key_src:
            hdr = 0 if origem_tem_cabecalho else None
            try:
                source_file.seek(0)
                ext = source_file.name.split('.')[-1].lower()
                engine_util = 'xlrd' if ext == 'xls' else ('openpyxl' if ext == 'xlsx' else None)
                raw = pd.read_excel(source_file, header=hdr, engine=engine_util)
                raw.columns = deduplicar_colunas(raw.columns) if origem_tem_cabecalho else [f"Col {i+1}" for i in range(len(raw.columns))]
                st.session_state["source_df"] = raw
                st.session_state["last_cache_key_src"] = cache_key_src
            except Exception as e: st.error(f"Erro ao ler arquivo: {e}")

    if dest_file:
        if "wb_data" not in st.session_state or st.session_state.get("last_dest_name") != dest_file.name:
            dest_file.seek(0)
            st.session_state["wb_data"] = dest_file.getvalue()
            st.session_state["last_dest_name"] = dest_file.name

    df_origem = st.session_state.get("source_df")
    wb_data = st.session_state.get("wb_data")

    if df_origem is not None and wb_data is not None:
        wb = load_workbook(io.BytesIO(wb_data))
        target_sheet = st.selectbox("Escolha a ABA na Planilha de Destino a ser Atualizada:", wb.sheetnames)
        ws = wb[target_sheet]

        st.subheader("3. Seleção de Registros")
        col_busca = st.selectbox("Coluna identificadora (para seleção):", df_origem.columns)
        opcoes_selecao = [f"{val} (Linha {idx})" for idx, val in df_origem[col_busca].items()]
        selected_options = st.multiselect("🔍 Escolha os registros:", opcoes_selecao)
        selected_indices = [int(item.split("(Linha ")[1].replace(")", "")) for item in selected_options]

        if selected_indices:
            st.info(f"📊 **{len(selected_indices)}** registro(s) selecionado(s) para atualização.")

        st.write("---")
        st.subheader("4. Correlação dos dados dos Arquivos ORIGEM X DESTINO")
        mapping = {}
        cols_ui = st.columns(4)
        opcoes_mapeamento = ["--- Não mapear ---", "⚠️ Auto-incrementar (Seq)"] + list(df_origem.columns)
        for i in range(1, ws.max_column + 1):
            header_val = ws.cell(row=header_dest, column=i).value
            with cols_ui[(i-1) % 4]:
                map_val = st.selectbox(f"Col {i} ({header_val or 'S/ Título'})", opcoes_mapeamento, key=f"map_{i}")
                if map_val != "--- Não mapear ---": mapping[i] = map_val

        st.write("---")
        st.subheader("5. Local da Atualização")
        modo_insercao = st.radio("Local de inserção:", ["Final da planilha", "A partir de uma linha específica"])
        target_row = st.number_input("Linha:", min_value=header_dest+1, value=header_dest+1) if modo_insercao == "A partir de uma linha específica" else ws.max_row + 1

        st.write("---")
        if st.button("🚀 Processar e Atualizar"):
            if not selected_indices: st.error("Selecione itens!"); st.stop()
            ref_row_idx = (target_row - 1) if modo_insercao == "A partir de uma linha específica" else ws.max_row
            base_seq = 0
            if ref_row_idx >= header_dest:
                val_acima = ws.cell(row=ref_row_idx, column=1).value
                try: base_seq = int(val_acima)
                except: base_seq = 0
            if modo_insercao == "A partir de uma linha específica": ws.insert_rows(target_row, amount=len(selected_indices))
            current_row = target_row
            seq_val = base_seq
            for idx in selected_indices:
                seq_val += 1
                ref_row_idx = current_row - 1
                for col_idx in range(1, ws.max_column + 1):
                    target_cell = ws.cell(row=current_row, column=col_idx)
                    ref_cell = ws.cell(row=ref_row_idx, column=col_idx)
                    copiar_estilo_completo(ref_cell, target_cell)
                    if col_idx == 1 or mapping.get(col_idx) == "⚠️ Auto-incrementar (Seq)": target_cell.value = seq_val
                    elif col_idx in mapping: target_cell.value = extrair_valor_limpo(df_origem, idx, mapping[col_idx])
                    else: target_cell.value = ref_cell.value
                current_row += 1
            buffer = io.BytesIO()
            wb.save(buffer)
            st.session_state["wb_data"] = buffer.getvalue()
            st.success("✅ Processamento concluído com sucesso!")
            st.download_button("📥 Baixar Versão Atualizada", st.session_state["wb_data"], "sinale_atualizado.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# --- OPÇÃO 2: ATUALIZAÇÕES GERAIS (MODIFICADA COM SEUS PEDIDOS) ---
elif menu_opcao == "ATUALIZAÇÕES GERAIS":
    titulo_estilizado("Atualizações Gerais")
    if st.session_state.get("wb_data") is not None: st.info("📁 Arquivo carregado.")
    else:
        sinale_file = st.file_uploader("Selecione o arquivo do SINALE (.xlsx)", type=["xlsx"])
        if sinale_file: st.session_state["wb_data"] = sinale_file.getvalue(); st.rerun()

    if st.session_state.get("wb_data") is not None:
        wb_temp = load_workbook(io.BytesIO(st.session_state["wb_data"]), data_only=True)
        target_sheet = st.selectbox("Escolha a ABA:", wb_temp.sheetnames)
        header = st.number_input("Linha do cabeçalho:", value=11, min_value=1)
        df = pd.read_excel(io.BytesIO(st.session_state["wb_data"]), sheet_name=target_sheet, header=header-1)
        
        col_filtro, val_filtro = st.columns(2)
        with col_filtro: filtro_col = st.selectbox("Coluna para buscar:", df.columns)
        with val_filtro: filtro_vals = st.multiselect("Selecione o(s) valor(es):", sorted([str(v) for v in df[filtro_col].dropna().unique()]))
        
        df_view = df.copy()
        if filtro_vals: df_view = df_view[df_view[filtro_col].astype(str).isin(filtro_vals)]
        
        # TOTAL PESQUISADO
        st.metric("TOTAL PESQUISADO", len(df_view))
        
        col_btns = st.columns([1, 1, 4])
        with col_btns[0]:
            if st.button("✅ Marcar Todos"): st.session_state['select_all'] = True; st.rerun()
        with col_btns[1]:
            if st.button("❌ Desmarcar Todos"): st.session_state['select_all'] = False; st.rerun()
            
        df_for_edit = df_view.copy()
        df_for_edit.insert(0, "Atualizar?", st.session_state.get('select_all', False))
        df_editado = st.data_editor(df_for_edit, column_config={"Atualizar?": st.column_config.CheckboxColumn()}, use_container_width=True)
        selecionados = df_editado[df_editado["Atualizar?"] == True]
        
        if not selecionados.empty:
            st.write("---")
            col_target = st.selectbox("Escolha o campo para alterar:", df.columns)
            
            # MOSTRAR VALOR ANTIGO
            valores_antigos = selecionados[col_target].dropna().unique()
            st.info(f"📌 Valor(es) atual(is): **{', '.join(map(str, valores_antigos))}**")
            
            # CAMPO NOVO VALOR
            novo_val = st.text_input(f"Digite o NOVO valor para '{col_target}':")
            
            if st.button("➕ Adicionar à Fila"):
                st.session_state['fila_modificacoes'].append({"indices": selecionados.index.tolist(), "coluna": col_target, "novo_valor": novo_val, "aba": target_sheet})
                st.rerun()

        if st.session_state['fila_modificacoes']:
            if st.button("📥 Baixar Arquivo Atualizado"):
                file_bytes = gerar_arquivo_atualizado_bytes(st.session_state["wb_data"], header, st.session_state['fila_modificacoes'], df, sheet_name=target_sheet)
                st.download_button("Clique aqui para baixar", file_bytes, "sinale_atualizado_final.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# --- DEMAIS OPÇÕES ---
elif menu_opcao == "LIMPAR ARQUIVO":
    if st.button("🗑️ Limpar Tudo"): st.session_state.clear(); st.rerun()
elif menu_opcao == "SOMENTE TRABALHADORES ATIVOS":
    titulo_estilizado("Filtro de Trabalhadores Ativos")
elif menu_opcao == "SAIR DO SISTEMA":
    st.stop()
