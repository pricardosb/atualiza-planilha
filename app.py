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

def formatar_datas_dataframe(df_input):
    """Remove o componente de horário das colunas de data para exibição limpa (DD/MM/AAAA)."""
    df_out = df_input.copy()
    for col in df_out.columns:
        if pd.api.types.is_datetime64_any_dtype(df_out[col]):
            df_out[col] = df_out[col].dt.strftime('%d/%m/%Y')
        else:
            df_out[col] = df_out[col].apply(
                lambda v: v.strftime('%d/%m/%Y') if isinstance(v, (datetime.datetime, datetime.date, pd.Timestamp)) else (
                    str(v).split(' ')[0] if isinstance(v, str) and (' 00:00:00' in str(v) or 'T00:00:00' in str(v)) else v
                )
            )
    return df_out

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
        pascoa - datetime.timedelta(days=47),
        pascoa - datetime.timedelta(days=2),
        datetime.date(ano, 4, 21),
        datetime.date(ano, 5, 1),
        pascoa + datetime.timedelta(days=60),
        datetime.date(ano, 9, 7),
        datetime.date(ano, 10, 12),
        datetime.date(ano, 11, 2),
        datetime.date(ano, 11, 15),
        datetime.date(ano, 11, 20),
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
                wd = data_atual.weekday()
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
                            name=current_font.name, size=current_font.size, bold=current_font.bold,
                            italic=current_font.italic, strike=current_font.strike, underline=current_font.underline, color="FF0000"
                        )
                    else:
                        cell.font = Font(color="FF0000")
                        
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()

def titulo_estilizado(subtitulo=""):
    st.markdown(f"<div style='text-align: center; padding: 1.5rem; background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%); color: white; border-radius: 12px; margin-bottom: 1.5rem;'><h1>⚡ SINALE WEB</h1><p>{subtitulo}</p></div>", unsafe_allow_html=True)

def extrair_mes_ano_m9(file_bytes_io, sheets_available):
    try:
        target_sheet = None
        for s in sheets_available:
            if "COM REMUNER" in s.strip().upper():
                target_sheet = s
                break
        if not target_sheet and sheets_available:
            target_sheet = sheets_available[0]
            
        file_bytes_io.seek(0)
        df_cell = pd.read_excel(file_bytes_io, sheet_name=target_sheet, header=None, nrows=9)
        val = df_cell.iloc[8, 12]
        
        if pd.isna(val) or str(val).strip() == "": return "SEM MÊS/ANO"
        if isinstance(val, (datetime.datetime, datetime.date)): return val.strftime("%m/%Y")
        return str(val).strip()
    except Exception:
        return "SEM MÊS/ANO"

def obter_nome_coluna_por_letra(df, colunas_disponiveis, letra):
    mapa_letras = {'A': 0, 'B': 1, 'C': 2, 'D': 3, 'E': 4, 'F': 5, 'G': 6, 'H': 7, 'I': 8, 'J': 9, 
                   'K': 10, 'L': 11, 'M': 12, 'N': 13, 'O': 14, 'P': 15, 'Q': 16, 'R': 17, 'S': 18, 
                   'T': 19, 'U': 20, 'V': 21, 'W': 22, 'X': 23, 'Y': 24, 'Z': 25}
    idx = mapa_letras.get(letra.upper())
    if idx is not None and idx < len(colunas_disponiveis):
        return colunas_disponiveis[idx]
    return None

def gerar_config_largura_colunas(df_subset, colunas):
    """Gera configuração de largura de colunas baseada exclusivamente no tamanho do CONTEÚDO."""
    config = {}
    for col in colunas:
        if col in df_subset.columns:
            max_len = df_subset[col].astype(str).str.len().max() if not df_subset[col].empty else 10
            if pd.isna(max_len) or max_len <= 12:
                config[col] = st.column_config.Column(width="small")
            elif max_len <= 35:
                config[col] = st.column_config.Column(width="medium")
            else:
                config[col] = st.column_config.Column(width="large")
    return config

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
        
        df_view_fmt = formatar_datas_dataframe(df_view[cols_para_ver])
        st.dataframe(df_view_fmt, use_container_width=True, hide_index=True)
        
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
        for f_idx, f in enumerate(uploaded_files):
            file_key = f"{f_idx}_{f.name}"
            f_bytes = f.getvalue()
            xl = pd.ExcelFile(io.BytesIO(f_bytes))
            sheets_available = xl.sheet_names
            
            pref_sheets = [s for s in sheets_available if any(p in s.strip().upper() for p in ["COM REMUNER", "SEM REMUNER"])]
            
            if pref_sheets:
                default_sheets = pref_sheets
                is_fallback = False
            else:
                default_sheets = [sheets_available[0]] if sheets_available else []
                is_fallback = True
            
            with st.expander(f"📁 Configurações para: Arquivo {f_idx+1} - {f.name}", expanded=True):
                selected_sheets = st.multiselect(
                    f"Selecione aba(s) para {f.name}", 
                    sheets_available, 
                    default=default_sheets, 
                    key=f"sheets_{file_key}"
                )
                
                sheet_config = {}
                for i, sheet in enumerate(selected_sheets):
                    st.markdown(f"**Aba: `{sheet}`**")
                    
                    sheet_upper = sheet.strip().upper()
                    if any(p in sheet_upper for p in ["COM REMUNER", "SEM REMUNER"]):
                        default_header = 11
                    else:
                        default_header = 10 if is_fallback else 11
                    
                    header_row = st.number_input(
                        f"Linha do cabeçalho para aba '{sheet}'", 
                        value=default_header, 
                        min_value=1, 
                        key=f"head_{file_key}_{sheet}"
                    )
                    
                    try:
                        df_preview = pd.read_excel(io.BytesIO(f_bytes), sheet_name=sheet, header=header_row-1, nrows=0)
                        cols_aba = [str(c).strip() for c in df_preview.columns]
                    except:
                        cols_aba = []
                    
                    default_col = None
                    for c in cols_aba:
                        c_up = str(c).strip().upper()
                        if c_up in ["NOME DO INTERNO", "NOME DO INTERNO "]:
                            default_col = c; break
                    if not default_col:
                        for c in cols_aba:
                            if str(c).strip().upper() == "NOME":
                                default_col = c; break
                    if not default_col:
                        for c in cols_aba:
                            if str(c).strip().upper().startswith("NOME"):
                                default_col = c; break
                    if not default_col:
                        for c in cols_aba:
                            if "NOME" in str(c).strip().upper():
                                default_col = c; break
                    if not default_col and len(cols_aba) > 8: default_col = cols_aba[8]
                    elif not default_col and cols_aba: default_col = cols_aba[0]
                    
                    opcoes_colunas = ["--- Não pesquisar nesta aba ---"] + cols_aba
                    default_idx = opcoes_colunas.index(default_col) if default_col in opcoes_colunas else 0
                    
                    col_escolhida = st.selectbox(
                        f"Selecione o campo (coluna) para a pesquisa na aba '{sheet}':", 
                        opcoes_colunas, 
                        index=default_idx, 
                        key=f"col_search_{file_key}_{sheet}"
                    )
                    
                    sheet_config[sheet] = {
                        "header_idx": header_row - 1,
                        "col_busca": col_escolhida if col_escolhida != "--- Não pesquisar nesta aba ---" else None
                    }
                    st.markdown("---")
                
                settings[file_key] = sheet_config
        
        if st.button("🔍 Carregar e Consolidar Dados para Pesquisa", key="btn_consolidar_op3"):
            all_results = []
            for f_idx, f in enumerate(uploaded_files):
                file_key = f"{f_idx}_{f.name}"
                f_bytes = f.getvalue()
                xl = pd.ExcelFile(io.BytesIO(f_bytes))
                mes_ano_m9 = extrair_mes_ano_m9(io.BytesIO(f_bytes), xl.sheet_names)
                
                file_cfg = settings.get(file_key, {})
                for sheet, cfg in file_cfg.items():
                    try:
                        df_tmp = pd.read_excel(io.BytesIO(f_bytes), sheet_name=sheet, header=cfg["header_idx"])
                        df_tmp.columns = [str(c).strip() for c in df_tmp.columns]
                        df_tmp.columns = deduplicar_colunas(df_tmp.columns)
                        
                        col_pedida = cfg.get("col_busca")
                        target_col = None
                        if col_pedida:
                            for c in df_tmp.columns:
                                if str(c).strip().upper() == str(col_pedida).strip().upper():
                                    target_col = c; break
                        if not target_col:
                            for c in df_tmp.columns:
                                if "NOME DO INTERNO" in str(c).strip().upper():
                                    target_col = c; break
                        if not target_col:
                            for c in df_tmp.columns:
                                if "NOME" in str(c).strip().upper():
                                    target_col = c; break
                        if not target_col and len(df_tmp.columns) > 8:
                            target_col = df_tmp.columns[8]
                        elif not target_col and len(df_tmp.columns) > 0:
                            target_col = df_tmp.columns[0]
                        
                        if target_col and target_col in df_tmp.columns:
                            colunas_originais = list(df_tmp.columns)
                            
                            df_tmp['MÊS/ANO - ABA'] = f"{mes_ano_m9} - {sheet}"
                            df_tmp['Aba Original'] = sheet
                            df_tmp['Campo Pesquisado'] = target_col
                            
                            val_nome = df_tmp[target_col].astype(str).str.strip()
                            df_tmp['Nome (Visualização)'] = val_nome + " - " + sheet
                            df_tmp['NOME_LIMPO'] = val_nome.str.upper()
                            
                            df_tmp = df_tmp[~df_tmp['NOME_LIMPO'].isin(['', 'NAN', 'NONE', '0', 'NAT', 'NC', 'N/C'])].copy()
                            
                            aba_upper = sheet.strip().upper()
                            if "COM REMUNER" in aba_upper:
                                letras_desejadas = ['B', 'I', 'J', 'T', 'U', 'V', 'W']
                            elif "SEM REMUNER" in aba_upper:
                                letras_desejadas = ['I', 'B', 'W', 'R', 'S', 'T', 'U']
                            else:
                                letras_desejadas = ['J', 'C', 'X', 'F', 'S', 'T', 'U', 'V']
                            
                            nomes_colunas_exibir = []
                            for let in letras_desejadas:
                                col_nome = obter_nome_coluna_por_letra(df_tmp, colunas_originais, let)
                                if col_nome:
                                    nomes_colunas_exibir.append(str(col_nome))
                            
                            df_tmp['Ordem_Colunas'] = "|".join(nomes_colunas_exibir)
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
        
        nomes_disponiveis = sorted(df_pesq['Nome (Visualização)'].dropna().unique())
        nomes_selecionados = st.multiselect(
            "🔍 Digite para pesquisar e selecione o(s) nome(s):",
            options=nomes_disponiveis,
            key="busca_nomes_op3"
        )
        
        df_view = df_pesq.copy()
        if nomes_selecionados:
            df_view = df_view[df_view['Nome (Visualização)'].isin(nomes_selecionados)]
            
        st.metric("Total de Registros Encontrados", len(df_view))
        
        if not df_view.empty:
            def categorizar_aba(aba_name):
                aba_upper = str(aba_name).strip().upper()
                if "COM REMUNER" in aba_upper:
                    return "COM REMUNERAÇÃO"
                elif "SEM REMUNER" in aba_upper:
                    return "SEM REMUNERAÇÃO"
                else:
                    return "OUTRAS ABAS"
            
            df_view['Categoria_Aba'] = df_view['Aba Original'].apply(categorizar_aba)
            
            df_display_all = formatar_datas_dataframe(df_view)
            
            grupos_categorias = [
                ("🟢 COM REMUNERAÇÃO", "COM REMUNERAÇÃO"),
                ("🟡 SEM REMUNERAÇÃO", "SEM REMUNERAÇÃO"),
                ("🔵 OUTRAS ABAS", "OUTRAS ABAS")
            ]
            
            for titulo_grupo, cat_key in grupos_categorias:
                df_grupo = df_display_all[df_display_all['Categoria_Aba'] == cat_key]
                
                if not df_grupo.empty:
                    cols_ordem = []
                    for seq in df_grupo['Ordem_Colunas'].dropna().unique():
                        for c in seq.split('|'):
                            if c and c not in cols_ordem:
                                cols_ordem.append(c)
                    
                    colunas_finais = ['MÊS/ANO - ABA']
                    for c in cols_ordem:
                        if c in df_grupo.columns and c not in colunas_finais:
                            colunas_finais.append(c)
                    
                    if len(colunas_finais) == 1:
                        for c in df_grupo.columns:
                            if c not in ['MÊS/ANO - ABA', 'Aba Original', 'Campo Pesquisado', 'Nome (Visualização)', 'NOME_LIMPO', 'Ordem_Colunas', 'Categoria_Aba']:
                                colunas_finais.append(c)
                    
                    col_config_conteudo = gerar_config_largura_colunas(df_grupo, colunas_finais)
                    
                    st.markdown(f"### {titulo_grupo} ({len(df_grupo)} registro(s))")
                    st.dataframe(
                        df_grupo[colunas_finais], 
                        column_config=col_config_conteudo,
                        use_container_width=True, 
                        hide_index=True
                    )
                    st.markdown("---")
        else:
            st.info("ℹ️ Nenhum registro selecionado ou encontrado na pesquisa.")

# --- DEMAIS OPÇÕES ---
elif menu_opcao == "LIMPAR ARQUIVO":
    if st.button("🗑️ Limpar Tudo"): st.session_state.clear(); st.rerun()
elif menu_opcao == "SOMENTE TRABALHADORES ATIVOS":
    titulo_estilizado("Filtro de Trabalhadores Ativos")
elif menu_opcao == "SAIR DO SISTEMA":
    st.stop()
