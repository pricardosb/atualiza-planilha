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
            
            # Se o campo atualizado for SAIDA (ou SAÍDA), pinta toda a linha de vermelho
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

# --- MENU ---
menu_opcao = st.sidebar.radio("Selecione a rotina:", [
    "INCLUSÃO DE TRABALHO",
    "ATUALIZAÇÕES GERAIS",
    "PESQUISA PARA REMIÇÃO",
    "LIMPAR ARQUIVO",
    "SOMENTE TRABALHADORES ATIVOS",
    "SAIR DO SISTEMA"
])

# --- OPÇÃO 1: INCLUSÃO DE TRABALHO (TRAVADA) ---
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

# --- OPÇÃO 2: ATUALIZAÇÕES GERAIS (TRAVADA) ---
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
            
            # Se a coluna for DIAS, exibe seletor de Mês/Ano e calcula os dias úteis / feriados antes de mostrar o valor antigo
            if col_target.strip().upper() == "DIAS":
                st.markdown("---")
                st.subheader("📅 Cálculo Automático de Dias Úteis (Seg a Sáb / Seg a Sex)")
                c_mes, c_ano = st.columns(2)
                meses_dict = {
                    "Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4,
                    "Maio": 5, "Junho": 6, "Julho": 7, "Agosto": 8,
                    "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12
                }
                with c_mes:
                    mes_escolhido_nome = st.selectbox("Selecione o Mês:", list(meses_dict.keys()), key="sel_mes_dias")
                    mes_num = meses_dict[mes_escolhido_nome]
                with c_ano:
                    ano_escolhido = st.number_input("Digite o Ano:", min_value=2020, max_value=2035, value=datetime.date.today().year, key="sel_ano_dias")
                
                stats = obter_estatisticas_mes(ano_escolhido, mes_num)
                
                st.info(f"""
                **Resumo para {mes_escolhido_nome}/{ano_escolhido}:**
                * **Segunda a Sábado:** {stats['seg_sab_total']} brutos | **Feriados (Seg-Sáb):** {stats['seg_sab_feriados']} | **Úteis (Seg a Sábado):** **{stats['seg_sab_uteis']}**
                * **Segunda a Sexta:** {stats['seg_sex_total']} brutos | **Feriados (Seg-Sex):** {stats['seg_sex_feriados']} | **Úteis (Seg a Sexta):** **{stats['seg_sex_uteis']}**
                """)
                
                if stats['feriados_detalhes']:
                    feriados_str = ", ".join([f"{f[0].strftime('%d/%m/%Y')}" for f in stats['feriados_detalhes']])
                    st.write(f"📌 **Feriados no período (Seg a Sábado):** {feriados_str}")
                else:
                    st.write("📌 **Nenhum feriado nacional** de Segunda a Sábado neste mês/ano.")
                st.markdown("---")
            
            valores_antigos_str = ", ".join([str(v) for v in selecionados[col_target].dropna().unique()])
            if not valores_antigos_str:
                valores_antigos_str = "Vazio"
            st.info(f"📌 **Valor(es) atual(is) / antigo(s)** no campo **'{col_target}'** para os registros selecionados: **{valores_antigos_str}**")
            
            novo_val = st.text_input("Digite o novo valor:", key="novo_val_op2")
            
            if st.button("➕ Adicionar à Fila de Modificações", key="btn_add_fila"):
                vl_busca_str = ", ".join(filtro_vals) if filtro_vals else "Todos"
                st.session_state['fila_modificacoes'].append({
                    "indices": selecionados.index.tolist(),
                    "coluna": col_target,
                    "novo_valor": novo_val,
                    "valor_antigo": valores_antigos_str,
                    "vl_busca": vl_busca_str,
                    "aba": target_sheet
                })
                st.success("Modificação adicionada à fila!"); st.rerun()

        if st.session_state['fila_modificacoes']:
            st.markdown("---")
            st.subheader("📋 Fila de Modificações Pendentes")
            
            dados_fila = []
            for idx, item in enumerate(st.session_state['fila_modificacoes']):
                dados_fila.append({
                    "Remover?": False,
                    "ID_ITEM": idx,
                    "QTD REG": len(item["indices"]),
                    "ABA": item.get("aba", "Geral"),
                    "VL BUSCA": item.get("vl_busca", "Todos"),
                    "CAMPO": item.get("coluna", ""),
                    "VL ANTIGO": item.get("valor_antigo", "Vazio"),
                    "NOVO VALOR": item.get("novo_valor", "")
                })
            
            df_fila_resumo = pd.DataFrame(dados_fila)
            
            df_fila_editado = st.data_editor(
                df_fila_resumo,
                column_config={
                    "Remover?": st.column_config.CheckboxColumn("Remover?"),
                    "ID_ITEM": None
                },
                disabled=["QTD REG", "ABA", "VL BUSCA", "CAMPO", "VL ANTIGO", "NOVO VALOR"],
                use_container_width=True,
                key="editor_fila"
            )
            
            col_f1, col_f2, col_f3 = st.columns(3)
            with col_f1:
                if st.button("🗑️ Remover Selecionados", key="btn_remover_sel"):
                    indices_para_remover = df_fila_editado[df_fila_editado["Remover?"] == True]["ID_ITEM"].tolist()
                    if indices_para_remover:
                        st.session_state['fila_modificacoes'] = [
                            item for i, item in enumerate(st.session_state['fila_modificacoes']) 
                            if i not in indices_para_remover
                        ]
                        st.success("Modificação(ões) selecionada(s) removida(s) da fila!")
                        st.rerun()
                    else:
                        st.warning("Marque ao menos um item na coluna 'Remover?' para excluir.")
            with col_f2:
                if st.button("🗑️ Limpar Fila Inteira", key="btn_limpar_fila"):
                    st.session_state['fila_modificacoes'] = []
                    st.success("Fila limpa com sucesso!")
                    st.rerun()
            with col_f3:
                file_bytes = gerar_arquivo_atualizado_bytes(io.BytesIO(st.session_state["wb_data"]), header, st.session_state['fila_modificacoes'], df, sheet_name=target_sheet)
                st.download_button("📥 Baixar Arquivo Atualizado", file_bytes, "sinale_atualizado_final.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# --- OPÇÃO 3: PESQUISA PARA REMIÇÃO ---
elif menu_opcao == "PESQUISA PARA REMIÇÃO":
    titulo_estilizado("Pesquisa para Remição")
    
    st.subheader("1. Carregar Arquivos")
    uploaded_files = st.file_uploader("Selecione um ou mais arquivos (.xlsx)", type=["xlsx"], accept_multiple_files=True, key="search_upload")
    
    if uploaded_files:
        settings = {}
        for f in uploaded_files:
            with st.expander(f"Configurações para: {f.name}"):
                c1, c2 = st.columns(2)
                sheets = st.text_input(f"Abas (separadas por vírgula, máx 2) para {f.name}", value="Plan1", key=f"sheets_{f.name}")
                header_row = st.number_input(f"Linha cabeçalho para {f.name}", value=1, min_value=1, key=f"head_{f.name}")
                settings[f.name] = {"sheets": [s.strip() for s in sheets.split(",")], "header": header_row-1}
        
        st.markdown("---")
        st.subheader("2. Parâmetros de Pesquisa")
        search_col = st.text_input("Nome da coluna para filtrar por valor (ex: 'Status'):")
        search_val = st.text_input("Valor a pesquisar nesta coluna (ex: 'Pendente'):")
        name_filter = st.text_input("Filtro por NOME (nome do interno/trabalhador):")
        
        if st.button("🔍 Iniciar Pesquisa"):
            all_results = []
            
            for f in uploaded_files:
                f.seek(0)
                conf = settings[f.name]
                for sheet in conf["sheets"]:
                    try:
                        df_tmp = pd.read_excel(f, sheet_name=sheet, header=conf["header"])
                        
                        # Filtro por Nome (procura em todas as colunas se não especificado)
                        if name_filter:
                            df_tmp = df_tmp[df_tmp.apply(lambda row: row.astype(str).str.contains(name_filter, case=False).any(), axis=1)]
                        
                        # Filtro por Coluna Específica
                        if search_col and search_val:
                            if search_col in df_tmp.columns:
                                df_tmp = df_tmp[df_tmp[search_col].astype(str) == str(search_val)]
                            else:
                                st.warning(f"Coluna '{search_col}' não encontrada na aba '{sheet}' do arquivo {f.name}")
                        
                        if not df_tmp.empty:
                            df_tmp['__ORIGEM__'] = f"{f.name} ({sheet})"
                            all_results.append(df_tmp)
                    except Exception as e:
                        st.error(f"Erro ao ler {f.name} - Aba {sheet}: {e}")
            
            if all_results:
                final_df = pd.concat(all_results, ignore_index=True)
                st.success(f"Foram encontrados {len(final_df)} resultados.")
                st.dataframe(final_df, use_container_width=True)
            else:
                st.warning("Nenhum resultado encontrado com os critérios fornecidos.")

# --- DEMAIS OPÇÕES ---
elif menu_opcao == "LIMPAR ARQUIVO":
    if st.button("🗑️ Limpar Tudo"): st.session_state.clear(); st.rerun()
elif menu_opcao == "SOMENTE TRABALHADORES ATIVOS":
    titulo_estilizado("Filtro de Trabalhadores Ativos")
elif menu_opcao == "SAIR DO SISTEMA":
    st.stop()
