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

def gerar_arquivo_atualizado_bytes(source_input, header, fila, df_original):
    if isinstance(source_input, bytes):
        wb = load_workbook(io.BytesIO(source_input))
    else:
        source_input.seek(0)
        wb = load_workbook(source_input)
        
    ws = wb[wb.sheetnames[0]]
    
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
    "LIMPAR ARQUIVO",
    "SOMENTE TRABALHADORES ATIVOS",
    "SAIR DO SISTEMA"
])

# --- OPÇÃO 1 ---
if menu_opcao == "ATUALIZAÇÃO DE DADOS - INCLUSÃO DE TRABALHO":
    titulo_estilizado("INTEGRADOR ==> DADOS GERAIS DO INTERNO >>> SINALE")
    
    # Opção para descartar memória antes de carregar
    if st.checkbox("🗑️ Descartar dados da memória e carregar novos arquivos", value=False, key="desc_op1"):
        st.session_state['source_df'] = None
        st.session_state['wb_data'] = None
        st.session_state['last_dest_name'] = None
        st.session_state['fila_modificacoes'] = []
        st.success("Memória limpa com sucesso! Faça o upload dos novos arquivos abaixo.")
        st.rerun()

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
            if not selected_indices: 
                st.error("Selecione itens!")
                st.stop()
            
            ref_row_idx = (target_row - 1) if modo_insercao == "A partir de uma linha específica" else ws.max_row
            
            base_seq = 0
            if ref_row_idx >= header_dest:
                val_acima = ws.cell(row=ref_row_idx, column=1).value
                try: base_seq = int(val_acima)
                except: base_seq = 0
            
            if modo_insercao == "A partir de uma linha específica":
                ws.insert_rows(target_row, amount=len(selected_indices))
            
            current_row = target_row
            seq_val = base_seq

            for idx in selected_indices:
                seq_val += 1
                for col_idx in range(1, ws.max_column + 1):
                    target_cell = ws.cell(row=current_row, column=col_idx)
                    ref_cell = ws.cell(row=ref_row_idx, column=col_idx)
                    copiar_estilo_completo(ref_cell, target_cell)
                    
                    if col_idx == 1 or mapping.get(col_idx) == "⚠️ Auto-incrementar (Seq)":
                        target_cell.value = seq_val
                    elif col_idx in mapping:
                        target_cell.value = extrair_valor_limpo(df_origem, idx, mapping[col_idx])
                    else:
                        target_cell.value = None
                current_row += 1

            if modo_insercao == "A partir de uma linha específica":
                for r in range(current_row, ws.max_row + 1):
                    val_atual = ws.cell(row=r, column=1).value
                    if val_atual is not None:
                        seq_val += 1
                        ws.cell(row=r, column=1, value=seq_val)

            buffer = io.BytesIO()
            wb.save(buffer)
            st.session_state["wb_data"] = buffer.getvalue()
            st.success("✅ Processamento concluído com sucesso e salvo na memória!")
            st.download_button("📥 Baixar Versão Atualizada", st.session_state["wb_data"], "sinale_atualizado.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.warning("⚠️ Carregue os arquivos de Origem e Destino para habilitar os campos.")

# --- OPÇÃO 2 ---
elif menu_opcao == "ATUALIZAÇÕES GERAIS":
    titulo_estilizado("Atualizações Gerais")
    
    if st.session_state.get("wb_data") is not None:
        st.info("📁 Arquivo carregado automaticamente da memória.")
        
        if st.checkbox("🗑️ Descartar dados da memória e carregar novo arquivo", value=False, key="desc_op2"):
            st.session_state["wb_data"] = None
            st.session_state['fila_modificacoes'] = []
            st.success("Memória limpa com sucesso!")
            st.rerun()
            
        sinale_file = io.BytesIO(st.session_state["wb_data"])
    else:
        st.warning("⚠️ Nenhum arquivo de destino encontrado na memória. Faça o upload abaixo.")
        sinale_file = st.file_uploader("Selecione o arquivo do SINALE (.xlsx)", type=["xlsx"])
        if sinale_file:
            st.session_state["wb_data"] = sinale_file.getvalue()
            st.session_state['last_sinale_name'] = sinale_file.name
            sinale_file = io.BytesIO(st.session_state["wb_data"])
            st.rerun()

    if st.session_state.get("wb_data") is not None:
        header = st.number_input("Linha do cabeçalho:", value=11, min_value=1)
        
        sinale_file.seek(0)
        df = pd.read_excel(sinale_file, header=header-1)
        
        if 'last_sinale_name' not in st.session_state:
            st.session_state['last_sinale_name'] = "arquivo_memoria.xlsx"

        st.subheader("🔍 Filtros de Visualização")
        cols_para_ver = st.multiselect("Quais campos deseja visualizar?", df.columns.tolist(), default=df.columns.tolist())
        
        col_filtro, val_filtro = st.columns(2)
        with col_filtro: filtro_col = st.selectbox("Coluna para buscar:", df.columns)
        valores_existentes = sorted([str(v) for v in df[filtro_col].dropna().unique()])
        with val_filtro: filtro_vals = st.multiselect("Selecione o(s) valor(es) para filtrar:", valores_existentes)
        
        df_view = df.copy()
        if filtro_vals: df_view = df_view[df_view[filtro_col].astype(str).isin(filtro_vals)]
        
        st.metric("Total de Registros Encontrados", len(df_view))
        st.dataframe(df_view[cols_para_ver], use_container_width=True)
        
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
        
        selecionados = df_editado[df_editado["Atualizar?"] == True]
        st.metric("Total de Registros Marcados", len(selecionados))
        
        if not selecionados.empty:
            col_target = st.selectbox("Selecione a coluna que deseja alterar:", df.columns)
            
            # MOSTRA O VALOR ANTIGO ANTES DE DIGITAR O NOVO VALOR
            valores_antigos_str = ", ".join([str(v) for v in selecionados[col_target].dropna().unique()])
            if not valores_antigos_str:
                valores_antigos_str = "Vazio"
            st.info(f"📌 **Valor(es) atual(is) / antigo(s)** no campo **'{col_target}'** para os registros selecionados: **{valores_antigos_str}**")
            
            novo_val = st.text_input("Digite o novo valor:")
            
            if st.button("➕ Adicionar à Fila de Modificações"):
                vl_busca_str = ", ".join(filtro_vals) if filtro_vals else "Todos"
                st.session_state['fila_modificacoes'].append({
                    "indices": selecionados.index.tolist(),
                    "coluna": col_target,
                    "novo_valor": novo_val,
                    "valor_antigo": valores_antigos_str,
                    "vl_busca": vl_busca_str
                })
                st.success("Modificação adicionada à fila!"); st.rerun()

        if st.session_state['fila_modificacoes']:
            st.markdown("---")
            st.subheader("📋 Fila de Modificações Pendentes")
            df_fila_resumo = pd.DataFrame([
                {
                    "QTD REG": len(item["indices"]),
                    "VL BUSCA": item["vl_busca"],
                    "CAMPO": item["coluna"],
                    "VL ANTIGO": item["valor_antigo"],
                    "NOVO VALOR": item["novo_valor"]
                }
                for item in st.session_state['fila_modificacoes']
            ])
            st.dataframe(df_fila_resumo, use_container_width=True)
            
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                if st.button("🗑️ Limpar Fila"): st.session_state['fila_modificacoes'] = []; st.rerun()
            with col_f2:
                file_bytes = gerar_arquivo_atualizado_bytes(io.BytesIO(st.session_state["wb_data"]), header, st.session_state['fila_modificacoes'], df)
                st.download_button("📥 Baixar Arquivo Atualizado", file_bytes, "sinale_atualizado_final.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# --- OPÇÃO 3: LIMPAR ARQUIVO ---
elif menu_opcao == "LIMPAR ARQUIVO":
    titulo_estilizado("Limpeza de Memória do Sistema")
    st.write("Clique abaixo para zerar os arquivos salvos em memória e começar um novo ciclo.")
    if st.button("🗑️ Limpar Dados da Memória"):
        st.session_state['source_df'] = None
        st.session_state['wb_data'] = None
        st.session_state['last_dest_name'] = None
        st.session_state['fila_modificacoes'] = []
        st.success("Memória limpa com sucesso!")
        st.rerun()

# --- OPÇÃO 4: SOMENTE TRABALHADORES ATIVOS ---
elif menu_opcao == "SOMENTE TRABALHADORES ATIVOS":
    titulo_estilizado("Filtro de Trabalhadores Ativos")
    
    if st.session_state.get("wb_data") is not None:
        st.info("📁 Utilizando arquivo presente na memória.")
        
        if st.checkbox("🗑️ Descartar dados da memória e carregar novo arquivo", value=False, key="desc_op4"):
            st.session_state["wb_data"] = None
            st.success("Memória limpa com sucesso!")
            st.rerun()
            
        wb = load_workbook(io.BytesIO(st.session_state["wb_data"]), data_only=True)
        aba_ativos = st.selectbox("Selecione a aba:", wb.sheetnames, key="aba_atv")
        
        df_ativos = pd.read_excel(io.BytesIO(st.session_state["wb_data"]), sheet_name=aba_ativos, header=10)
        
        colunas_disp = df_ativos.columns.tolist()
        col_filtro = st.selectbox("Selecione a coluna de status/situação:", colunas_disp)
        
        if col_filtro:
            valores_unicos = df_ativos[col_filtro].dropna().unique().tolist()
            valor_escolhido = st.selectbox("Selecione o valor correspondente a 'Ativo':", valores_unicos)
            
            if st.button("Filtrar Registros Ativos"):
                df_filtrado = df_ativos[df_ativos[col_filtro] == valor_escolhido]
                st.success(f"Foram encontrados {len(df_filtrado)} registros ativos.")
                st.dataframe(df_filtrado, use_container_width=True)
    else:
        st.warning("⚠️ Nenhum arquivo carregado na memória. Faça o upload abaixo.")
        up_memo = st.file_uploader("Arquivo Excel:", type=["xlsx"], key="up_memo_ativos")
        if up_memo:
            st.session_state["wb_data"] = up_memo.getvalue()
            st.rerun()

# --- OPÇÃO 5: SAIR DO SISTEMA ---
elif menu_opcao == "SAIR DO SISTEMA":
    titulo_estilizado("Sessão Encerrada")
    st.info("Você pode fechar esta aba do navegador com segurança.")
    st.stop()
