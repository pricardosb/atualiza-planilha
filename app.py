import io
import pandas as pd
import numpy as np
import streamlit as st
from openpyxl import load_workbook
from copy import copy

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="SINALE WEB", layout="wide")

# --- FUNÇÕES DE SUPORTE (ORIGINAIS) ---
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
        if isinstance(val, pd.Series): val = val.iloc[0]
        if pd.isna(val): return None
        if hasattr(val, 'item'): val = val.item()
        return val
    except Exception: return None

def titulo_estilizado(subtitulo=""):
    st.markdown(
        f"<div style='text-align: center; padding: 1.5rem; background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%); color: white; border-radius: 12px; margin-bottom: 1.5rem; box-shadow: 0 4px 6px rgba(0,0,0,0.1);'>"
        f"<h1 style='margin:0;'>⚡ SINALE WEB</h1>"
        f"<p style='margin:0; font-size: 1.1rem; opacity: 0.9;'>{subtitulo}</p>"
        f"</div>", unsafe_allow_html=True
    )

# --- MENU ---
menu_opcao = st.sidebar.radio("Selecione a rotina:", [
    "ATUALIZAÇÃO DE DADOS - INCLUSÃO DE TRABALHO",
    "ATUALIZAÇÕES GERAIS",
    "LIMPAR ARQUIVO",
    "SOMENTE TRABALHADORES ATIVOS",
    "SAIR DO SISTEMA"
])

# --- OPÇÃO 1: INCLUSÃO DE TRABALHO (CÓDIGO ORIGINAL COMPLETO) ---
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
                    
                    if col_idx == 1 or mapping.get(col_idx) == "⚠️ Auto-incrementar (Seq)":
                        target_cell.value = seq_val
                    elif col_idx in mapping:
                        target_cell.value = extrair_valor_limpo(df_origem, idx, mapping[col_idx])
            
            buffer = io.BytesIO(); wb.save(buffer)
            st.success("Processado!"); st.download_button("📥 Baixar", buffer.getvalue(), "sinale_atualizado.xlsx")

# --- OPÇÃO 2: ATUALIZAÇÕES GERAIS ---
elif menu_opcao == "ATUALIZAÇÕES GERAIS":
    titulo_estilizado("Atualizações Gerais de Dados")
    sinale_file = st.file_uploader("Selecione o arquivo do SINALE (.xlsx)", type=["xlsx"])
    header = st.number_input("Linha do cabeçalho:", value=11, min_value=1)
    
    if sinale_file:
        wb = load_workbook(sinale_file)
        aba = st.selectbox("Escolha a aba:", wb.sheetnames)
        ws = wb[aba]
        df = pd.read_excel(sinale_file, sheet_name=aba, header=header-1)
        
        tab1, tab2 = st.tabs(["📊 Visualizar/Pesquisar", "✏️ Atualizar Dados"])
        with tab1:
            st.dataframe(df, use_container_width=True)
        with tab2:
            cabecalhos = {str(ws.cell(row=header, column=c).value).strip(): c for c in range(1, ws.max_column + 1)}
            dados_tabela = []
            for r in range(header + 1, ws.max_row + 1):
                row = {"Selecionar": False}
                for nome, c_idx in cabecalhos.items(): row[nome] = ws.cell(row=r, column=c_idx).value
                row["_linha"] = r
                dados_tabela.append(row)
            
            df_edit = pd.DataFrame(dados_tabela)
            df_selecionado = st.data_editor(df_edit, column_config={"Selecionar": st.column_config.CheckboxColumn()}, hide_index=True)
            
            linhas_alvo = df_selecionado[df_selecionado["Selecionar"] == True]["_linha"].tolist()
            if linhas_alvo:
                col_alvo = st.selectbox("Coluna para alterar:", list(cabecalhos.keys()))
                novo_valor = st.text_input("Novo Valor:")
                if st.button("🚀 Processar Atualização"):
                    for r in linhas_alvo:
                        ws.cell(row=r, column=cabecalhos[col_alvo], value=novo_valor)
                    buffer = io.BytesIO(); wb.save(buffer)
                    st.download_button("📥 Baixar Arquivo Atualizado", buffer.getvalue(), "sinale_atualizado.xlsx")

# --- OUTRAS OPÇÕES ---
elif menu_opcao == "LIMPAR ARQUIVO":
    st.write("Funcionalidade de Limpar Arquivo...")
elif menu_opcao == "SOMENTE TRABALHADORES ATIVOS":
    st.write("Funcionalidade de Filtro de Ativos...")
elif menu_opcao == "SAIR DO SISTEMA":
    st.stop()
