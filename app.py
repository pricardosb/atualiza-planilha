import io
import pandas as pd
import numpy as np
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font
from copy import copy

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="SINALE WEB", layout="wide")

# --- FUNÇÕES DE SUPORTE (ORIGINAIS) ---
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
    if val_str is None or str(val_str).strip() == "":
        return None
    val_str = str(val_str).strip()
    
    if pd.api.types.is_integer_dtype(dtype_original):
        try: return int(val_str)
        except ValueError: pass
    elif pd.api.types.is_float_dtype(dtype_original):
        try: return float(val_str.replace(',', '.'))
        except ValueError: pass
    elif pd.api.types.is_datetime64_any_dtype(dtype_original):
        try: return pd.to_datetime(val_str)
        except: pass

    if '.' not in val_str and ',' not in val_str:
        try: return int(val_str)
        except ValueError: pass
        
    try: return float(val_str.replace(',', '.'))
    except ValueError: pass
    
    return val_str

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

# --- OPÇÃO 1: INCLUSÃO DE TRABALHO (ORIGINAL COMPLETA) ---
if menu_opcao == "ATUALIZAÇÃO DE DADOS - INCLUSÃO DE TRABALHO":
    titulo_estilizado("Rotina: Inclusão de Trabalho")
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
                raw = pd.read_excel(source_file, header=hdr)
                raw.columns = deduplicar_colunas(raw.columns) if origem_tem_cabecalho else [f"Col {i+1}" for i in range(len(raw.columns))]
                st.session_state["source_df"] = raw
                st.session_state["last_cache_key_src"] = cache_key_src
            except: st.error("Erro ao ler arquivo.")

    if dest_file:
        if "wb_data" not in st.session_state or st.session_state.get("last_dest_name") != dest_file.name:
            dest_file.seek(0)
            st.session_state["wb_data"] = dest_file.getvalue()
            st.session_state["last_dest_name"] = dest_file.name

    df_origem = st.session_state.get("source_df")
    wb_data = st.session_state.get("wb_data")

    if df_origem is not None and wb_data is not None:
        wb = load_workbook(io.BytesIO(wb_data))
        target_sheet = st.selectbox("Escolha a ABA:", wb.sheetnames)
        ws = wb[target_sheet]
        st.subheader("3. Seleção de Registros")
        col_busca = st.selectbox("Coluna identificadora:", df_origem.columns)
        opcoes_selecao = [f"{val} (Linha {idx})" for idx, val in df_origem[col_busca].items()]
        selected_options = st.multiselect("🔍 Escolha os registros:", opcoes_selecao)
        selected_indices = [int(item.split("(Linha ")[1].replace(")", "")) for item in selected_options]
        st.subheader("4. Correlação ORIGEM X DESTINO")
        mapping = {}
        cols_ui = st.columns(4)
        opcoes_mapeamento = ["--- Não mapear ---", "⚠️ Auto-incrementar (Seq)"] + list(df_origem.columns)
        for i in range(1, ws.max_column + 1):
            header_val = ws.cell(row=header_dest, column=i).value
            with cols_ui[(i-1) % 4]:
                map_val = st.selectbox(f"Col {i} ({header_val or 'S/ Título'})", opcoes_mapeamento, key=f"map_{i}")
                if map_val != "--- Não mapear ---": mapping[i] = map_val
        if st.button("🚀 Processar e Atualizar"):
            ref_row_idx = ws.max_row
            seq_val = int(ws.cell(row=ref_row_idx, column=1).value) if ws.cell(row=ref_row_idx, column=1).value else 0
            for idx in selected_indices:
                seq_val += 1
                new_row = ws.max_row + 1
                for col_idx in range(1, ws.max_column + 1):
                    target_cell = ws.cell(row=new_row, column=col_idx)
                    ref_cell = ws.cell(row=ref_row_idx, column=col_idx)
                    copiar_estilo_completo(ref_cell, target_cell)
                    if col_idx == 1 or mapping.get(col_idx) == "⚠️ Auto-incrementar (Seq)": target_cell.value = seq_val
                    elif col_idx in mapping: target_cell.value = extrair_valor_limpo(df_origem, idx, mapping[col_idx])
            buffer = io.BytesIO(); wb.save(buffer)
            st.success("Processado!"); st.download_button("📥 Baixar", buffer.getvalue(), "sinale_atualizado.xlsx")

# --- OPÇÃO 2: ATUALIZAÇÕES GERAIS ---
elif menu_opcao == "ATUALIZAÇÕES GERAIS":
    titulo_estilizado("Atualizações Gerais")
    sinale_file = st.file_uploader("Selecione o arquivo do SINALE (.xlsx)", type=["xlsx"])
    header = st.number_input("Linha do cabeçalho:", value=11, min_value=1)
    
    if sinale_file:
        # Inicializa a fila na session_state se não existir ou se trocar de arquivo
        if 'fila_modificacoes' not in st.session_state:
            st.session_state['fila_modificacoes'] = []
        if 'last_sinale_name' not in st.session_state or st.session_state['last_sinale_name'] != sinale_file.name:
            st.session_state['fila_modificacoes'] = []
            st.session_state['last_sinale_name'] = sinale_file.name

        wb = load_workbook(sinale_file)
        ws = wb[wb.sheetnames[0]]
        df = pd.read_excel(sinale_file, header=header-1)
        
        # 1. VISUALIZAÇÃO E PESQUISA INTELIGENTE
        st.subheader("🔍 Filtros de Visualização")
        cols_para_ver = st.multiselect("Quais campos deseja visualizar?", df.columns.tolist(), default=df.columns.tolist())
        
        col_filtro, val_filtro = st.columns(2)
        with col_filtro: filtro_col = st.selectbox("Coluna para buscar:", df.columns)
        valores_existentes = sorted([str(v) for v in df[filtro_col].dropna().unique()])
        with val_filtro: filtro_vals = st.multiselect("Selecione o(s) valor(es) para filtrar:", valores_existentes)
        
        df_view = df.copy()
        if filtro_vals:
            df_view = df_view[df_view[filtro_col].astype(str).isin(filtro_vals)]
        
        st.metric("Total de Registros Encontrados", len(df_view))
        st.dataframe(df_view[cols_para_ver], use_container_width=True)
        
        # 2. SELEÇÃO PARA ATUALIZAÇÃO
        st.subheader("✏️ Seleção para Atualizar")
        
        if 'select_all' not in st.session_state: st.session_state['select_all'] = False
        
        cols_btns = st.columns([1, 1, 4])
        with cols_btns[0]:
            if st.button("✅ Marcar Todos"): st.session_state['select_all'] = True; st.rerun()
        with cols_btns[1]:
            if st.button("❌ Desmarcar Todos"): st.session_state['select_all'] = False; st.rerun()
            
        df_for_edit = df_view.copy()
        df_for_edit.insert(0, "Atualizar?", st.session_state['select_all'])
        
        df_editado = st.data_editor(df_for_edit, column_config={"Atualizar?": st.column_config.CheckboxColumn()}, use_container_width=True)
        
        # 3. CONTAGEM, VALOR ATUAL E ADIÇÃO À FILA
        selecionados = df_editado[df_editado["Atualizar?"] == True]
        st.metric("Total de Registros Marcados", len(selecionados))
        
        if not selecionados.empty:
            col_target = st.selectbox("Selecione a coluna que deseja alterar:", df.columns)
            
            valores_atuais = selecionados[col_target].dropna().unique()
            if len(valores_atuais) == 1:
                st.info(f"💡 **Valor Atual no(s) registro(s) selecionado(s):** `{valores_atuais[0]}`")
            elif len(valores_atuais) > 1:
                st.warning(f"⚠️ **Atenção:** Os registros selecionados possuem **{len(valores_atuais)} valores diferentes** nesta coluna: `{list(valores_atuais)}`")
            else:
                st.info("💡 **Valor Atual:** *(Vazio / Nulo)*")
            
            novo_val = st.text_input("Digite o novo valor:")
            
            if st.button("➕ Adicionar à Fila de Modificações"):
                st.session_state['fila_modificacoes'].append({
                    "indices": selecionados.index.tolist(),
                    "coluna": col_target,
                    "novo_valor": novo_val
                })
                st.success("Modificação adicionada à fila com sucesso! Você pode continuar fazendo novas modificações.")
                st.rerun()

        # 4. EXIBIÇÃO DA FILA E PROCESSAMENTO FINAL
        if st.session_state['fila_modificacoes']:
            st.markdown("---")
            st.subheader("📋 Fila de Modificações Pendentes")
            
            df_fila_resumo = pd.DataFrame([
                {"Qtd Registros": len(item["indices"]), "Coluna": item["coluna"], "Novo Valor": item["novo_valor"]}
                for item in st.session_state['fila_modificacoes']
            ])
            st.dataframe(df_fila_resumo, use_container_width=True)
            
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                if st.button("🗑️ Limpar Fila de Modificações"):
                    st.session_state['fila_modificacoes'] = []
                    st.rerun()
            with col_f2:
                if st.button("🚀 Processar Todas e Baixar Arquivo Final"):
                    red_font = Font(color="FF0000")
                    
                    # Identificação da coluna REM para a regra de negócio
                    rem_col_idx = None
                    for idx_c, c_name in enumerate(df.columns):
                        if str(c_name).strip().upper() == 'REM':
                            rem_col_idx = idx_c + 1
                            break
                    
                    for mod in st.session_state['fila_modificacoes']:
                        col_target = mod["coluna"]
                        novo_val = mod["novo_valor"]
                        valor_convertido = converter_valor_inteligente(novo_val, df[col_target].dtype)
                        
                        is_rz = col_target.strip().upper() in ['RZ', 'R.Z.']
                        is_saida = col_target.strip().upper() in ['SAIDA', 'SAÍDA']
                        
                        for idx in mod["indices"]:
                            excel_row = idx + header + 1
                            col_alvo_idx = df.columns.get_loc(col_target) + 1
                            
                            # Atualiza o campo principal
                            ws.cell(row=excel_row, column=col_alvo_idx, value=valor_convertido)
                            
                            # REGRA 2: Se RZ >= 3, atualiza REM com divisão inteira por 3
                            if is_rz and rem_col_idx is not None:
                                try:
                                    num_rz = float(str(valor_convertido).replace(',', '.'))
                                    if num_rz >= 3:
                                        rem_val = int(num_rz // 3)
                                        ws.cell(row=excel_row, column=rem_col_idx, value=rem_val)
                                except:
                                    pass
                                    
                            # REGRA 1: Se o campo alterado for SAÍDA, deixa os caracteres da linha inteira em vermelho
                            if is_saida:
                                for c_idx in range(1, ws.max_column + 1):
                                    cell_obj = ws.cell(row=excel_row, column=c_idx)
                                    cur_font = cell_obj.font
                                    if cur_font:
                                        cell_obj.font = Font(name=cur_font.name, size=cur_font.size, bold=cur_font.bold, italic=cur_font.italic, color="FF0000")
                                    else:
                                        cell_obj.font = red_font
                    
                    buffer = io.BytesIO()
                    wb.save(buffer)
                    st.success("Todas as modificações foram processadas com sucesso!")
                    st.download_button("📥 Baixar Arquivo Atualizado Final", buffer.getvalue(), "sinale_atualizado_final.xlsx")

# --- OUTRAS OPÇÕES ---
elif menu_opcao == "LIMPAR ARQUIVO":
    st.write("Funcionalidade de Limpeza...")
elif menu_opcao == "SOMENTE TRABALHADORES ATIVOS":
    st.write("Funcionalidade de Filtro de Ativos...")
elif menu_opcao == "SAIR DO SISTEMA":
    st.stop()
