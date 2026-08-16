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
    """Converte valores respeitando o tipo original da coluna (Data, Número ou Texto)"""
    if val_str is None or str(val_str).strip() == "":
        return None
    val_str = str(val_str).strip()
    
    # 1. Tratamento de Datas
    if pd.api.types.is_datetime64_any_dtype(dtype_original):
        try:
            return pd.to_datetime(val_str, dayfirst=True).to_pydatetime()
        except:
            return val_str

    # 2. Tratamento de Inteiros
    if pd.api.types.is_integer_dtype(dtype_original):
        try:
            return int(float(val_str.replace(',', '.')))
        except:
            return val_str

    # 3. Tratamento de Floats (Decimais)
    if pd.api.types.is_float_dtype(dtype_original):
        try:
            return float(val_str.replace(',', '.'))
        except:
            return val_str

    # 4. Texto (Padrão)
    return val_str

def gerar_arquivo_atualizado_bytes(source_input, header, fila, df_original, sheet_name=None):
    if isinstance(source_input, bytes):
        wb = load_workbook(io.BytesIO(source_input))
    else:
        source_input.seek(0)
        wb = load_workbook(source_input)
        
    ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb[wb.sheetnames[0]]
    
    red_font = Font(color="FF0000")
    
    col_indices = {}
    for idx_c, c_name in enumerate(df_original.columns):
        c_norm = str(c_name).strip().upper()
        col_indices[c_norm] = idx_c + 1
        
    rem_col_idx = col_indices.get('REM')
    salario_col_idx = col_indices.get('SALARIO') or col_indices.get('SALÁRIO')
    falta_col_idx = col_indices.get('FALTA') or col_indices.get('FALTAS')
    peculio_col_idx = col_indices.get('PECULIO') or col_indices.get('PECÚLIO')
    
    modified_cols_by_row = {}
    
    for mod in fila:
        col_target = mod["coluna"]
        novo_val = mod["novo_valor"]
        
        # Uso da função melhorada de conversão
        valor_convertido = converter_valor_inteligente(novo_val, df_original[col_target].dtype)
        
        col_target_norm = col_target.strip().upper()
        is_rz = col_target_norm in ['RZ', 'R.Z.']
        is_saida = col_target_norm in ['SAIDA', 'SAÍDA']
        
        for idx in mod["indices"]:
            if idx not in modified_cols_by_row:
                modified_cols_by_row[idx] = set()
            modified_cols_by_row[idx].add(col_target_norm)
            
            excel_row = idx + header + 1
            col_alvo_idx = df_original.columns.get_loc(col_target) + 1
            
            ws.cell(row=excel_row, column=col_alvo_idx, value=valor_convertido)
            
            if is_rz and rem_col_idx is not None:
                try:
                    num_rz = float(str(valor_convertido).replace(',', '.'))
                    if num_rz >= 3:
                        rem_val = int(num_rz // 3)
                        ws.cell(row=excel_row, column=rem_col_idx, value=rem_val)
                except:
                    pass
                    
            if is_saida:
                for c_idx in range(1, ws.max_column + 1):
                    cell_obj = ws.cell(row=excel_row, column=c_idx)
                    cur_font = cell_obj.font
                    if cur_font:
                        cell_obj.font = Font(name=cur_font.name, size=cur_font.size, bold=cur_font.bold, italic=cur_font.italic, color="FF0000")
                    else:
                        cell_obj.font = red_font
                        
    if peculio_col_idx is not None and salario_col_idx is not None:
        for idx, cols_mod in modified_cols_by_row.items():
            if any(c in ['SALARIO', 'SALÁRIO'] for c in cols_mod):
                excel_row = idx + header + 1
                sal_val_cell = ws.cell(row=excel_row, column=salario_col_idx).value
                try:
                    salario = float(str(sal_val_cell).replace(',', '.'))
                except:
                    salario = 0.0
                
                falta_atualizada = any(c in ['FALTA', 'FALTAS'] for c in cols_mod)
                
                if falta_atualizada and falta_col_idx is not None:
                    falta_val_cell = ws.cell(row=excel_row, column=falta_col_idx).value
                    try:
                        falta = float(str(falta_val_cell).replace(',', '.'))
                    except:
                        falta = 0.0
                    peculio = (salario - falta) * 0.25
                else:
                    peculio = salario * 0.25
                    
                ws.cell(row=excel_row, column=peculio_col_idx, value=peculio)

    buffer = io.BytesIO()
    wb.save(buffer)
    file_bytes = buffer.getvalue()
    
    st.session_state["wb_data"] = file_bytes
    return file_bytes

def titulo_estilizado(subtitulo=""):
    st.markdown(f"<div style='text-align: center; padding: 1.5rem; background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%); color: white; border-radius: 12px; margin-bottom: 1.5rem;'><h1>⚡ SINALE WEB</h1><p>{subtitulo}</p></div>", unsafe_allow_html=True)

# --- MENU ---
menu_opcao = st.sidebar.radio("Selecione a rotina:", [
    "ATUALIZAÇÃO DE DADOS - INCLUSÃO DE TRABALHO",
    "ATUALIZAÇÕES GERAIS",
    "GERAR RELATORIOS",
    "SOMENTE TRABALHADORES ATIVOS",
    "SAIR DO SISTEMA"
])

# --- OPÇÃO 1 ---
if menu_opcao == "ATUALIZAÇÃO DE DADOS - INCLUSÃO DE TRABALHO":
    titulo_estilizado("INTEGRADOR ==> DADOS GERAIS DO INTERNO >>> SINALE")
    
    # ... [Resto da lógica da Opção 1 mantida] ...
    if st.checkbox("🗑️ Descartar dados da memória e carregar novos arquivos", value=False, key="desc_op1"):
        st.session_state['source_df'] = None
        st.session_state['wb_data'] = None
        st.session_state['last_dest_name'] = None
        st.session_state['fila_modificacoes'] = []
        st.success("Memória limpa com sucesso!")
        st.rerun()
    # (Inserção omitida por brevidade, mas deve ser a mesma do bloco anterior)
    # ...

# --- OPÇÃO 2 ---
# --- OPÇÃO 2: ATUALIZAÇÕES GERAIS ---
# --- OPÇÃO 2: ATUALIZAÇÕES GERAIS ---
elif menu_opcao == "ATUALIZAÇÕES GERAIS":
    titulo_estilizado("Atualizações Gerais")
    
    if st.session_state.get("wb_data") is not None:
        st.info("📁 Arquivo carregado automaticamente da memória.")
        if st.checkbox("🗑️ Descartar dados da memória e carregar novo arquivo", value=False, key="desc_op2"):
            st.session_state["wb_data"] = None
            st.session_state['fila_modificacoes'] = []
            st.rerun()
    else:
        sinale_file = st.file_uploader("Selecione o arquivo do SINALE (.xlsx)", type=["xlsx"], key="upload_op2")
        if sinale_file:
            st.session_state["wb_data"] = sinale_file.getvalue()
            st.rerun()

    if st.session_state.get("wb_data") is not None:
        wb_temp = load_workbook(io.BytesIO(st.session_state["wb_data"]), data_only=True)
        target_sheet = st.selectbox("Escolha a ABA do arquivo para trabalhar:", wb_temp.sheetnames, key="aba_op2")
        header = st.number_input("Linha do cabeçalho:", value=11, min_value=1, key="header_op2")
        
        df = pd.read_excel(io.BytesIO(st.session_state["wb_data"]), sheet_name=target_sheet, header=header-1)
        
        st.subheader("🔍 Filtros de Visualização")
        col_filtro = st.selectbox("Coluna para buscar:", df.columns, key="filtro_col_op2")
        valores_existentes = sorted([str(v) for v in df[col_filtro].dropna().unique()])
        filtro_vals = st.multiselect("Selecione o(s) valor(es) para filtrar:", valores_existentes, key="filtro_vals_op2")
        
        df_view = df.copy()
        if filtro_vals: df_view = df_view[df_view[col_filtro].astype(str).isin(filtro_vals)]
        
        st.metric("Total de Registros Encontrados", len(df_view))
        
        # --- BOTÕES DE MARCAÇÃO ---
        if 'select_all' not in st.session_state: st.session_state['select_all'] = False
        
        cols_btns = st.columns([1, 1, 4])
        with cols_btns[0]:
            if st.button("✅ Marcar Todos"): st.session_state['select_all'] = True; st.rerun()
        with cols_btns[1]:
            if st.button("❌ Desmarcar Todos"): st.session_state['select_all'] = False; st.rerun()
        
        # --- EDITOR ---
        st.subheader("✏️ Seleção para Atualizar")
        df_for_edit = df_view.copy()
        df_for_edit.insert(0, "Atualizar?", st.session_state['select_all'])
        
        df_editado = st.data_editor(
            df_for_edit, 
            column_config={"Atualizar?": st.column_config.CheckboxColumn()}, 
            use_container_width=True, 
            key="editor_op2"
        )
        
        selecionados = df_editado[df_editado["Atualizar?"] == True]
        st.metric("Total de Registros Marcados", len(selecionados))
        
        if not selecionados.empty:
            col_target = st.selectbox("Selecione a coluna que deseja alterar:", df.columns, key="col_target_op2")
            
            # --- MOSTRAR VALOR ANTIGO (RESTAURADO) ---
            valores_antigos_str = ", ".join([str(v) for v in selecionados[col_target].dropna().unique()])
            if not valores_antigos_str: valores_antigos_str = "Vazio"
            st.info(f"📌 **Valor(es) atual(is) / antigo(s)** no campo **'{col_target}'** para os registros selecionados: **{valores_antigos_str}**")
            
            novo_val = st.text_input("Digite o novo valor (Data, Número ou Texto):", key="novo_val_op2")
            
            if st.button("➕ Adicionar à Fila de Modificações", key="btn_add_fila"):
                st.session_state['fila_modificacoes'].append({
                    "indices": selecionados.index.tolist(),
                    "coluna": col_target,
                    "novo_valor": novo_val,
                    "valor_antigo": valores_antigos_str,
                    "vl_busca": ", ".join(filtro_vals) if filtro_vals else "Todos",
                    "aba": target_sheet
                })
                # Limpa o campo após adicionar
                st.session_state['novo_val_op2'] = "" 
                st.success("Modificação adicionada à fila!"); st.rerun()

        # Exibição da fila e download...
        if st.session_state['fila_modificacoes']:
            st.markdown("---")
            st.subheader("📋 Fila de Modificações Pendentes")
            file_bytes = gerar_arquivo_atualizado_bytes(io.BytesIO(st.session_state["wb_data"]), header, st.session_state['fila_modificacoes'], df, sheet_name=target_sheet)
            st.download_button("📥 Baixar Arquivo Atualizado", file_bytes, "sinale_atualizado_final.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
# --- OPÇÕES 3, 4 e 5 ---
elif menu_opcao == "GERAR RELATORIOS":
    titulo_estilizado("Gerar Relatórios")
    # ...
elif menu_opcao == "SOMENTE TRABALHADORES ATIVOS":
    titulo_estilizado("Filtro de Trabalhadores Ativos")
    # ...
elif menu_opcao == "SAIR DO SISTEMA":
    st.stop()
