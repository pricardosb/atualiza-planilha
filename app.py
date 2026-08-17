import io
import calendar
import datetime
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
if 'file_settings' not in st.session_state: st.session_state['file_settings'] = {}
if 'pesquisa_df' not in st.session_state: st.session_state['pesquisa_df'] = None

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

def calcular_pascoa(ano):
    a = ano % 19
    b = ano // 100
    c = ano % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    L = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * L) // 451
    mes = (h + L - 7 * m + 114) // 31
    dia = ((h + L - 7 * m + 114) % 31) + 1
    return datetime.date(ano, mes, dia)

def obter_estatisticas_mes(ano, mes):
    cal = calendar.monthcalendar(ano, mes)
    pascoa = calcular_pascoa(ano)
    feriados = [
        datetime.date(ano, 1, 1),
        pascoa - datetime.timedelta(days=47), # Carnaval
        pascoa - datetime.timedelta(days=2),  # Sexta-feira Santa
        datetime.date(ano, 4, 21),
        datetime.date(ano, 5, 1),
        pascoa + datetime.timedelta(days=60), # Corpus Christi
        datetime.date(ano, 9, 7),
        datetime.date(ano, 10, 12),
        datetime.date(ano, 11, 2),
        datetime.date(ano, 11, 15),
        datetime.date(ano, 11, 20), # Consciência Negra
        datetime.date(ano, 12, 25)
    ]
    feriados_mes = [f for f in feriados if f.month == mes and f.year == ano]
    
    dias_seg_sex_total = 0
    dias_seg_sab_total = 0
    feriados_seg_sex = 0
    feriados_seg_sab = 0
    lista_feriados_detalhes = []
    
    for semana in cal:
        for i in range(7):
            dia = semana[i]
            if dia != 0:
                data_atual = datetime.date(ano, mes, dia)
                wd = data_atual.weekday() # 0=Seg, ..., 5=Sáb, 6=Dom
                if wd < 5:
                    dias_seg_sex_total += 1
                    dias_seg_sab_total += 1
                elif wd == 5:
                    dias_seg_sab_total += 1
                    
                if data_atual in feriados_mes:
                    if wd < 5:
                        feriados_seg_sex += 1
                        feriados_seg_sab += 1
                        lista_feriados_detalhes.append((data_atual, "Seg a Sex"))
                    elif wd == 5:
                        feriados_seg_sab += 1
                        lista_feriados_detalhes.append((data_atual, "Sábado"))

    return {
        "seg_sex_total": dias_seg_sex_total,
        "seg_sex_feriados": feriados_seg_sex,
        "seg_sex_uteis": dias_seg_sex_total - feriados_seg_sex,
        "seg_sab_total": dias_seg_sab_total,
        "seg_sab_feriados": feriados_seg_sab,
        "seg_sab_uteis": dias_seg_sab_total - feriados_seg_sab,
        "feriados_detalhes": lista_feriados_detalhes
    }

def gerar_arquivo_atualizado_bytes(source_input, header, fila, df_original, sheet_name=None):
    wb = load_workbook(io.BytesIO(source_input) if isinstance(source_input, bytes) else source_input)
    ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb[wb.sheetnames[0]]
    for mod in fila:
        col_target = mod["coluna"]
        valor_convertido = converter_valor_inteligente(mod["novo_valor"], df_original[col_target].dtype)
        for idx in mod["indices"]:
            excel_row = idx + header + 1
            ws.cell(row=excel_row, column=df_original.columns.get_loc(col_target) + 1, value=valor_convertido)
            
            if col_target.strip().upper() in ["SAIDA", "SAÍDA"]:
                for col_idx in range(1, ws.max_column + 1):
                    cell = ws.cell(row=excel_row, column=col_idx)
                    current_font = cell.font
                    if current_font:
                        cell.font = Font(
                            name=current_font.name,
                            size=current_font.size,
                            bold=current_font.bold,
                            italic=current_font.italic,
                            strike=current_font.strike,
                            underline=current_font.underline,
                            color="FF0000"
                        )
                    else:
                        cell.font = Font(color="FF0000")
                        
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()

def titulo_estilizado(subtitulo=""):
    st.markdown(f"<div style='text-align: center; padding: 1.5rem; background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%); color: white; border-radius: 12px; margin-bottom: 1.5rem;'><h1>⚡ SINALE WEB</h1><p>{subtitulo}</p></div>", unsafe_allow_html=True)

def extrair_mes_ano_m9(file_bytes, sheets_available):
    """Extrai a informação de MÊS/ANO na célula M9 (Linha 9, Coluna M/13) da aba 'COM REMUNERAÇÃO'."""
    try:
        # Tenta localizar a aba 'COM REMUNERAÇÃO'
        target_sheet = None
        for s in sheets_available:
            if "COM REMUNER" in s.strip().upper():
                target_sheet = s
                break
        if not target_sheet and sheets_available:
            target_sheet = sheets_available[0]
            
        file_bytes.seek(0)
        # Lê até a linha 9 (sem cabeçalho)
        df_cell = pd.read_excel(file_bytes, sheet_name=target_sheet, header=None, nrows=9)
        val = df_cell.iloc[8, 12] # Linha 9 (índice 8), Coluna M (índice 12)
        
        if pd.isna(val) or str(val).strip() == "":
            return "SEM MÊS/ANO"
            
        if isinstance(val, (datetime.datetime, datetime.date)):
            return val.strftime("%m/%Y")
            
        val_str = str(val).strip()
        return val_str
    except Exception:
        return "SEM MÊS/ANO"

# --- MENU ---
menu_opcao = st.sidebar.radio("Selecione a rotina:", [
    "INCLUSÃO DE TRABALHO",
    "ATUALIZAÇÕES GERAIS",
    "PESQUISA PARA REMIÇÃO",
    "LIMPAR ARQUIVO",
    "SOMENTE TRABALHADORES ATIVOS",
    "SAIR DO SISTEMA"
])

# --- OPÇÃO 1: INCLUSÃO DE TRABALHO ---
if menu_opcao == "INCLUSÃO DE TRABALHO":
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

# --- OPÇÃO 2: ATUALIZAÇÕES GERAIS ---
elif menu_opcao == "ATUALIZAÇÕES GERAIS":
    titulo_estilizado("Atualizações Gerais")
    
    if st.session_state.get("wb_data") is not None:
        st.info("📁 Arquivo carregado automaticamente da memória.")
        if st.checkbox("🗑️ Descartar dados da memória e carregar novo arquivo", value=False, key="desc_op2"):
            st.session_state["wb_data"] = None
            st.session_state['fila_modificacoes'] = []
            st.success("Memória limpa com sucesso!")
            st.rerun()
    else:
        st.warning("⚠️ Nenhum arquivo de destino encontrado na memória. Faça o upload abaixo.")
        sinale_file = st.file_uploader("Selecione o arquivo do SINALE (.xlsx)", type=["xlsx"], key="upload_op2")
        if sinale_file:
            st.session_state["wb_data"] = sinale_file.getvalue()
            st.session_state['last_sinale_name'] = sinale_file.name
            st.rerun()

    if st.session_state.get("wb_data") is not None:
        wb_temp = load_workbook(io.BytesIO(st.session_state["wb_data"]), data_only=True)
        target_sheet = st.selectbox("Escolha a ABA do arquivo para trabalhar:", wb_temp.sheetnames, key="aba_op2")
        header = st.number_input("Linha do cabeçalho:", value=11, min_value=1, key="header_op2")
        df = pd.read_excel(io.BytesIO(st.session_state["wb_data"]), sheet_name=target_sheet, header=header-1)

        st.subheader("🔍 Filtros de Visualização")
        cols_para_ver = st.multiselect("Quais campos deseja visualizar?", df.columns.tolist(), default=df.columns.tolist())
        col_filtro, val_filtro = st.columns(2)
        with col_filtro: filtro_col = st.selectbox("Coluna para buscar:", df.columns, key="filtro_col_op2")
        valores_existentes = sorted([str(v) for v in df[filtro_col].dropna().unique()])
        with val_filtro: filtro_vals = st.multiselect("Selecione o(s) valor(es) para filtrar:", valores_existentes, key="filtro_vals_op2")
        
        df_view = df.copy()
        if filtro_vals: df_view = df_view[df_view[filtro_col].astype(str).isin(filtro_vals)]
        st.metric("Total de Registros Encontrados", len(df_view))
        st.dataframe(df_view[cols_para_ver], use_container_width=True)
        
        st.subheader("✏️ Seleção para Atualizar")
        if 'select_all' not in st.session_state: st.session_state['select_all'] = False
        cols_btns = st.columns([1, 1, 4])
        with cols_btns[0]:
            if st.button("✅ Marcar Todos", key="btn_marcar_t"): st.session_state['select_all'] = True; st.rerun()
        with cols_btns[1]:
            if st.button("❌ Desmarcar Todos", key="btn_desmarcar_t"): st.session_state['select_all'] = False; st.rerun()
        
        df_for_edit = df_view.copy()
        df_for_edit.insert(0, "Atualizar?", st.session_state['select_all'])
        df_editado = st.data_editor(df_for_edit, column_config={"Atualizar?": st.column_config.CheckboxColumn()}, use_container_width=True, key="editor_op2")
        
        selecionados = df_editado[df_editado["Atualizar?"] == True]
        st.metric("Total de Registros Marcados", len(selecionados))
        
        if not selecionados.empty:
            col_target = st.selectbox("Selecione a coluna que deseja alterar:", df.columns, key="col_target_op2")
            if col_target.strip().upper() == "DIAS":
                st.markdown("---")
                st.subheader("📅 Cálculo Automático de Dias Úteis (Seg a Sáb / Seg a Sex)")
                c_mes, c_ano = st.columns(2)
                meses_dict = {"Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4, "Maio": 5, "Junho": 6, "Julho": 7, "Agosto": 8, "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12}
                with c_mes: mes_escolhido_nome = st.selectbox("Selecione o Mês:", list(meses_dict.keys()), key="sel_mes_dias"); mes_num = meses_dict[mes_escolhido_nome]
                with c_ano: ano_escolhido = st.number_input("Digite o Ano:", min_value=2020, max_value=2035, value=datetime.date.today().year, key="sel_ano_dias")
                stats = obter_estatisticas_mes(ano_escolhido, mes_num)
                st.info(f"**Resumo para {mes_escolhido_nome}/{ano_escolhido}:**\n* **Segunda a Sábado:** {stats['seg_sab_total']} brutos | **Úteis:** **{stats['seg_sab_uteis']}**\n* **Segunda a Sexta:** {stats['seg_sex_total']} brutos | **Úteis:** **{stats['seg_sex_uteis']}**")
            
            valores_antigos_str = ", ".join([str(v) for v in selecionados[col_target].dropna().unique()])
            st.info(f"📌 **Valor(es) atual(is) / antigo(s)** no campo **'{col_target}'**: **{valores_antigos_str if valores_antigos_str else 'Vazio'}**")
            novo_val = st.text_input("Digite o novo valor:", key="novo_val_op2")
            
            if st.button("➕ Adicionar à Fila de Modificações", key="btn_add_fila"):
                st.session_state['fila_modificacoes'].append({"indices": selecionados.index.tolist(), "coluna": col_target, "novo_valor": novo_val, "valor_antigo": valores_antigos_str, "vl_busca": ", ".join(filtro_vals) if filtro_vals else "Todos", "aba": target_sheet})
                st.success("Modificação adicionada à fila!"); st.rerun()

        if st.session_state['fila_modificacoes']:
            st.markdown("---")
            st.subheader("📋 Fila de Modificações Pendentes")
            df_fila_resumo = pd.DataFrame([{"Remover?": False, "ID_ITEM": i, "ABA": item.get("aba", "Geral"), "CAMPO": item.get("coluna", ""), "NOVO VALOR": item.get("novo_valor", "")} for i, item in enumerate(st.session_state['fila_modificacoes'])])
            df_fila_editado = st.data_editor(df_fila_resumo, column_config={"Remover?": st.column_config.CheckboxColumn("Remover?"), "ID_ITEM": None}, disabled=["ABA", "CAMPO", "NOVO VALOR"], use_container_width=True, key="editor_fila")
            col_f1, col_f2, col_f3 = st.columns(3)
            with col_f1:
                if st.button("🗑️ Remover Selecionados"):
                    indices = df_fila_editado[df_fila_editado["Remover?"] == True]["ID_ITEM"].tolist()
                    st.session_state['fila_modificacoes'] = [item for i, item in enumerate(st.session_state['fila_modificacoes']) if i not in indices]
                    st.rerun()
            with col_f3:
                file_bytes = gerar_arquivo_atualizado_bytes(io.BytesIO(st.session_state["wb_data"]), header, st.session_state['fila_modificacoes'], df, sheet_name=target_sheet)
                st.download_button("📥 Baixar Arquivo Atualizado", file_bytes, "sinale_atualizado_final.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# --- OPÇÃO 3: PESQUISA PARA REMIÇÃO ---
elif menu_opcao == "PESQUISA PARA REMIÇÃO":
    titulo_estilizado("Pesquisa para Remição")
    
    st.subheader("1. Configuração de Arquivos, Abas e Campos")
    uploaded_files = st.file_uploader("Selecione um ou mais arquivos (.xlsx, .xls, .ods)", type=["xlsx", "xls", "ods"], accept_multiple_files=True, key="search_upload")
    
    if uploaded_files:
        settings = {}
        for f in uploaded_files:
            f.seek(0)
            xl = pd.ExcelFile(f)
            sheets_available = xl.sheet_names
            
            with st.expander(f"📁 Configurações para: {f.name}", expanded=True):
                # Regra de seleção automática de abas: 'COM REMUNERAÇÃO' e 'SEM REMUNERAÇÃO'
                pref_sheets = [s for s in sheets_available if any(p in s.strip().upper() for p in ["COM REMUNER", "SEM REMUNER"])]
                default_sheets = pref_sheets if pref_sheets else (sheets_available[:2] if sheets_available else [])
                
                selected_sheets = st.multiselect(f"Selecione até 2 abas para {f.name}", sheets_available, default=default_sheets, max_selections=2, key=f"sheets_{f.name}")
                
                sheet_config = {}
                for i, sheet in enumerate(selected_sheets):
                    st.markdown(f"**Aba: `{sheet}`**")
                    default_header = 11 if i == 0 else 10
                    header_row = st.number_input(f"Linha do cabeçalho para aba '{sheet}'", value=default_header, min_value=1, key=f"head_{f.name}_{sheet}")
                    
                    try:
                        f.seek(0)
                        df_preview = pd.read_excel(f, sheet_name=sheet, header=header_row-1, nrows=0)
                        cols_aba = list(df_preview.columns)
                    except:
                        cols_aba = []
                    
                    # Lógica de seleção inteligente da coluna de consulta por aba
                    sheet_upper = sheet.strip().upper()
                    default_col = None
                    
                    if "COM REMUNER" in sheet_upper:
                        # Para 'COM REMUNERAÇÃO': Procura 'NOME' ou Coluna I (índice 8)
                        for c in cols_aba:
                            if str(c).strip().upper() == "NOME":
                                default_col = c
                                break
                        if not default_col and len(cols_aba) > 8:
                            default_col = cols_aba[8]
                    elif "SEM REMUNER" in sheet_upper:
                        # Para 'SEM REMUNERAÇÃO': Procura 'NOME DO INTERNO' ou Coluna I (índice 8)
                        for c in cols_aba:
                            if str(c).strip().upper() == "NOME DO INTERNO":
                                default_col = c
                                break
                        if not default_col and len(cols_aba) > 8:
                            default_col = cols_aba[8]
                    else:
                        # Para outros arquivos / abas: Procura 'NOME' ou Coluna I
                        for c in cols_aba:
                            if str(c).strip().upper() == "NOME":
                                default_col = c
                                break
                        if not default_col:
                            for c in cols_aba:
                                if "NOME" in str(c).strip().upper():
                                    default_col = c
                                    break
                        if not default_col and len(cols_aba) > 8:
                            default_col = cols_aba[8]
                        elif not default_col and cols_aba:
                            default_col = cols_aba[0]
                    
                    opcoes_colunas = ["--- Não pesquisar nesta aba ---"] + cols_aba
                    default_idx = opcoes_colunas.index(default_col) if default_col in opcoes_colunas else 0
                    
                    col_escolhida = st.selectbox(f"Selecione o campo (coluna) para a pesquisa na aba '{sheet}':", opcoes_colunas, index=default_idx, key=f"col_search_{f.name}_{sheet}")
                    
                    sheet_config[sheet] = {
                        "header_idx": header_row - 1,
                        "col_busca": col_escolhida if col_escolhida != "--- Não pesquisar nesta aba ---" else None
                    }
                    st.markdown("---")
                
                settings[f.name] = sheet_config
        
        if st.button("🔍 Carregar e Consolidar Dados para Pesquisa", key="btn_consolidar_op3"):
            all_results = []
            for f in uploaded_files:
                f.seek(0)
                xl = pd.ExcelFile(f)
                mes_ano_m9 = extrair_mes_ano_m9(f, xl.sheet_names)
                
                file_cfg = settings.get(f.name, {})
                for sheet, cfg in file_cfg.items():
                    try:
                        f.seek(0)
                        df_tmp = pd.read_excel(f, sheet_name=sheet, header=cfg["header_idx"])
                        target_col = cfg["col_busca"]
                        if target_col and target_col in df_tmp.columns:
                            df_tmp['MÊS/ANO - ABA'] = f"{mes_ano_m9} - {sheet}"
                            df_tmp['Campo Pesquisado'] = target_col
                            all_results.append(df_tmp)
                    except Exception as e:
                        st.error(f"Erro ao ler {f.name} - Aba {sheet}: {e}")
            
            if all_results:
                st.session_state['pesquisa_df'] = pd.concat(all_results, ignore_index=True)
                st.success(f"Dados consolidados com sucesso! **{len(st.session_state['pesquisa_df'])}** registros carregados.")
            else:
                st.warning("Nenhum dado encontrado com as configurações informadas.")
                st.session_state['pesquisa_df'] = None
    else:
        st.session_state['pesquisa_df'] = None

    if st.session_state.get('pesquisa_df') is not None:
        df_pesq = st.session_state['pesquisa_df']
        st.markdown("---")
        st.subheader("🔍 Filtros de Visualização e Busca")
        
        cols_controle = ['MÊS/ANO - ABA', 'Campo Pesquisado']
        outras_cols = [c for c in df_pesq.columns if c not in cols_controle]
        lista_colunas_full = cols_controle + outras_cols
        
        cols_para_ver = st.multiselect("Selecione os campos que deseja visualizar:", options=lista_colunas_full, default=cols_controle, key="cols_ver_op3")
        
        col_filtro, val_filtro = st.columns(2)
        with col_filtro: filtro_col = st.selectbox("Coluna para buscar:", df_pesq.columns, key="filtro_col_op3")
        valores_existentes = sorted([str(v) for v in df_pesq[filtro_col].dropna().unique()])
        with val_filtro: filtro_vals = st.multiselect("Selecione o(s) valor(es) para filtrar:", valores_existentes, key="filtro_vals_op3")
        
        df_view = df_pesq.copy()
        if filtro_vals: df_view = df_view[df_view[filtro_col].astype(str).isin(filtro_vals)]
        st.metric("Total de Registros Encontrados", len(df_view))
        if cols_para_ver: st.dataframe(df_view[cols_para_ver], use_container_width=True)
        else: st.info("ℹ️ Selecione ao menos um campo acima para exibir a tabela de visualização.")

# --- DEMAIS OPÇÕES ---
elif menu_opcao == "LIMPAR ARQUIVO":
    if st.button("🗑️ Limpar Tudo"): st.session_state.clear(); st.rerun()
elif menu_opcao == "SOMENTE TRABALHADORES ATIVOS":
    titulo_estilizado("Filtro de Trabalhadores Ativos")
elif menu_opcao == "SAIR DO SISTEMA":
    st.stop()
