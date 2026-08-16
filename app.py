import io
import pandas as pd
import numpy as np
import streamlit as st
from openpyxl import load_workbook
from copy import copy

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="SINALE WEB - Em Busca de Agilidade", layout="wide")

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

def titulo_estilizado(subtitulo=""):
    html_content = (
        "<div style='text-align: center; padding: 1.8rem; background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%); border-radius: 12px; margin-bottom: 1.5rem; box-shadow: 0 6px 12px rgba(0,0,0,0.15);'>"
        "<h1 style='color: white; font-family: \"Segoe UI\", Tahoma, Geneva, Verdana, sans-serif; font-weight: 800; font-size: 2.2rem; margin: 0; letter-spacing: 1.5px; text-transform: uppercase;'>"
        "⚡ SINALE WEB"
        "</h1>"
        "<p style='color: #e0e6ed; font-family: \"Segoe UI\", Tahoma, Geneva, Verdana, sans-serif; font-weight: 300; font-size: 1.15rem; margin-top: 6px; letter-spacing: 2px; text-transform: uppercase;'>"
        "— Em Busca de Agilidade —"
        "</p>"
        "</div>"
    )
    if subtitulo:
        html_content += (
            f"<div style='margin-bottom: 1.5rem;'>"
            f"<h3 style='color: #2a5298; border-bottom: 2px solid #2a5298; padding-bottom: 5px; font-weight: 600;'>{subtitulo}</h3>"
            f"</div>"
        )
    st.markdown(html_content, unsafe_allow_html=True)

def aviso_sinale():
    html_aviso = (
        "<div style='background-color: #f8f9fa; border-left: 4px solid #2a5298; padding: 12px 16px; border-radius: 6px; margin-bottom: 1.5rem; color: #333; font-family: \"Segoe UI\", sans-serif; box-shadow: 0 2px 4px rgba(0,0,0,0.05);'>"
        "⚠️ <b>Atenção:</b> É preciso ter acesso ao sistema <b>SINALE</b> para obter este arquivo."
        "</div>"
    )
    st.markdown(html_aviso, unsafe_allow_html=True)

# --- MENU DE BARRA LATERAL ---
st.sidebar.title("📌 Menu de Opções")
menu_opcao = st.sidebar.radio(
    "Selecione a rotina:",
    [
        "ATUALIZAÇÃO DE DADOS - INCLUSÃO DE TRABALHO",
        "INFORMAÇÕES GERAIS",
        "ATUALIZAR DADOS",
        "LIMPAR ARQUIVO",
        "SOMENTE TRABALHADORES ATIVOS",
        "SAIR DO SISTEMA"
    ]
)

if menu_opcao == "SAIR DO SISTEMA":
    st.warning("Sessão encerrada. Atualize a página caso deseje retornar.")
    st.stop()

# --- ROTEAMENTO DAS OPÇÕES ---
if menu_opcao == "ATUALIZAÇÃO DE DADOS - INCLUSÃO DE TRABALHO":
    titulo_estilizado("Rotina: Inclusão de Trabalho")

    # --- CARREGAMENTO ---
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
        st.subheader("4. Correlação dos dados dos Arquivos ORIGEM X DESTINO")
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

        # --- 5. LOCAL DA ATUALIZAÇÃO ---
        st.subheader("5. Local da Atualização")
        modo_insercao = st.radio("Local de inserção:", ["Final da planilha", "A partir de uma linha específica"])
        target_row = st.number_input("Linha:", min_value=header_dest+1, value=header_dest+1) if modo_insercao == "A partir de uma linha específica" else ws.max_row + 1

        st.write("---")

        # --- PROCESSAMENTO ---
        if st.button("🚀 Processar e Atualizar"):
            if not selected_indices: st.error("Selecione itens!"); st.stop()
            
            ref_row_idx = (target_row - 1) if modo_insercao == "A partir de uma linha específica" else (ws.max_row)
            
            base_seq = 0
            if ref_row_idx >= header_dest:
                val_acima = ws.cell(row=ref_row_idx, column=1).value
                try:
                    base_seq = int(val_acima)
                except:
                    base_seq = 0
            
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
                        origem_col = mapping[col_idx]
                        target_cell.value = extrair_valor_limpo(df_origem, idx, origem_col)
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
            buffer.seek(0)
            st.success("✅ Processamento concluído com sucesso!")
            st.download_button("📥 Baixar Arquivo Atualizado", buffer.getvalue(), "sinale_atualizado.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif menu_opcao == "INFORMAÇÕES GERAIS":
    titulo_estilizado("Informações Gerais do Sistema SINALE")
    aviso_sinale()
    
    sinale_file = st.file_uploader("Selecione o arquivo exportado do SINALE (.xlsx)", type=["xlsx"], key="sinale_info_upload")
    header_sinale = st.number_input("Linha do cabeçalho no arquivo SINALE:", value=11, min_value=1, key="hdr_sinale_info")
    
    if sinale_file:
        try:
            wb_sinale = load_workbook(sinale_file)
            sheet_sinale = st.selectbox("Escolha a aba do arquivo SINALE:", wb_sinale.sheetnames, key="sheet_sinale_info")
            
            sinale_file.seek(0)
            df_info = pd.read_excel(sinale_file, sheet_name=sheet_sinale, header=header_sinale-1)
            ws_s = wb_sinale[sheet_sinale]
            
            st.write("---")
            st.markdown("### 📊 Resumo do Arquivo e Aba Selecionada")
            col_i1, col_i2, col_i3 = st.columns(3)
            col_i1.metric("Total de Linhas (Excel)", ws_s.max_row)
            col_i2.metric("Total de Colunas (Excel)", ws_s.max_column)
            col_i3.metric("Registros Carregados (DF)", len(df_info))
            
            st.write(f"- **Aba em análise:** `{sheet_sinale}`")
            st.write(f"- **Linha de cabeçalho:** {header_sinale}")
            
            st.write("---")
            st.subheader("🔍 Pesquisa e Seleção de Registros")
            col_busca_info = st.selectbox("Coluna identificadora para busca:", df_info.columns, key="col_busca_info")
            
            opcoes_selecao_info = [f"{val} (Linha {idx})" for idx, val in df_info[col_busca_info].items()]
            selected_options_info = st.multiselect("🔍 Escolha/Pesquise os registros:", opcoes_selecao_info, key="sel_opts_info")
            selected_indices_info = [int(item.split("(Linha ")[1].replace(")", "")) for item in selected_options_info]

            if selected_indices_info:
                st.info(f"📊 **{len(selected_indices_info)}** registro(s) selecionado(s).")
                
                st.subheader("👁️ Seleção de Colunas para Exibição")
                colunas_selecionadas = st.multiselect(
                    "Escolha quais colunas deseja visualizar nos registros selecionados:",
                    options=list(df_info.columns),
                    default=list(df_info.columns),
                    key="cols_sel_info"
                )
                
                if colunas_selecionadas:
                    df_filtrado = df_info.loc[selected_indices_info, colunas_selecionadas]
                    st.dataframe(df_filtrado, use_container_width=True)
                else:
                    st.warning("Selecione pelo menos uma coluna para visualizar os dados.")
            else:
                st.info("💡 Selecione um ou mais registros acima para pesquisar e visualizar os detalhes.")
        except Exception as e:
            st.error(f"Erro ao processar a planilha selecionada: {e}")

elif menu_opcao == "ATUALIZAR DADOS":
    titulo_estilizado("Atualizar Dados do SINALE")
    aviso_sinale()
    
    sinale_file_upd = st.file_uploader("Selecione o arquivo do SINALE para atualizar (.xlsx)", type=["xlsx"], key="sinale_upd_upload")
    header_upd = st.number_input("Linha do cabeçalho no arquivo SINALE:", value=11, min_value=1, key="hdr_sinale_upd")
    
    if sinale_file_upd:
        wb_upd = load_workbook(sinale_file_upd)
        sheet_upd = st.selectbox("Escolha a aba a ser tratada:", wb_upd.sheetnames, key="sheet_sinale_upd")
        ws_u = wb_upd[sheet_upd]
        
        st.write("---")
        st.write(f"Aba selecionada: **{sheet_upd}**")
        st.write("Insira as correções ou ajustes necessários nos dados abaixo:")
        
        if st.button("🚀 Processar e Baixar Atualização"):
            buffer = io.BytesIO()
            wb_upd.save(buffer)
            buffer.seek(0)
            st.success("✅ Arquivo atualizado com sucesso!")
            st.download_button("📥 Baixar Arquivo SINALE Atualizado", buffer.getvalue(), "sinale_atualizado_dados.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif menu_opcao == "LIMPAR ARQUIVO":
    titulo_estilizado("Limpar Arquivo do SINALE")
    aviso_sinale()
    
    sinale_file_clean = st.file_uploader("Selecione o arquivo do SINALE para limpeza (.xlsx)", type=["xlsx"], key="sinale_clean_upload")
    header_clean = st.number_input("Linha do cabeçalho no arquivo SINALE:", value=11, min_value=1, key="hdr_sinale_clean")
    
    if sinale_file_clean:
        wb_clean = load_workbook(sinale_file_clean)
        sheet_clean = st.selectbox("Escolha a aba para limpeza:", wb_clean.sheetnames, key="sheet_sinale_clean")
        ws_c = wb_clean[sheet_clean]
        
        if st.button("🗑️ Executar Limpeza e Baixar"):
            if ws_c.max_row > header_clean:
                ws_c.delete_rows(header_clean + 1, ws_c.max_row - header_clean)
            
            buffer = io.BytesIO()
            wb_clean.save(buffer)
            buffer.seek(0)
            st.success(f"✅ Arquivo limpo com sucesso na aba **{sheet_clean}**, mantendo a estrutura!")
            st.download_button("📥 Baixar Arquivo Limpo", buffer.getvalue(), "sinale_limpo.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif menu_opcao == "SOMENTE TRABALHADORES ATIVOS":
    titulo_estilizado("Somente Trabalhadores Ativos")
    aviso_sinale()
    
    sinale_file_active = st.file_uploader("Selecione o arquivo do SINALE para filtragem (.xlsx)", type=["xlsx"], key="sinale_active_upload")
    header_active = st.number_input("Linha do cabeçalho no arquivo SINALE:", value=11, min_value=1, key="hdr_sinale_active")
    
    if sinale_file_active:
        try:
            sinale_file_active.seek(0)
            wb_active = load_workbook(sinale_file_active)
            sheet_active = st.selectbox("Escolha a aba:", wb_active.sheetnames, key="sheet_sinale_active")
            
            sinale_file_active.seek(0)
            df_preview = pd.read_excel(sinale_file_active, sheet_name=sheet_active, header=header_active-1)
            ws_a = wb_active[sheet_active]
            
            st.write("---")
            col_status = st.selectbox("Selecione a coluna que indica o status/situação:", df_preview.columns, key="col_status_active")
            val_ativo = st.text_input("Informe o texto que define o trabalhador ATIVO:", value="ATIVO", key="val_ativo_input")
            
            if st.button("🚀 Filtrar Apenas Ativos e Baixar"):
                linhas_para_remover = []
                col_idx_excel = None
                for c_idx in range(1, ws_a.max_column + 1):
                    if ws_a.cell(row=header_active, column=c_idx).value == col_status:
                        col_idx_excel = c_idx
                        break
                if not col_idx_excel:
                    col_idx_excel = list(df_preview.columns).index(col_status) + 1

                for r in range(header_active + 1, ws_a.max_row + 1):
                    val_celula = ws_a.cell(row=r, column=col_idx_excel).value
                    if val_celula is None or str(val_ativo).strip().upper() not in str(val_celula).strip().upper():
                        linhas_para_remover.append(r)
                
                for r in sorted(linhas_para_remover, reverse=True):
                    ws_a.delete_rows(r, 1)
                
                buffer = io.BytesIO()
                wb_active.save(buffer)
                buffer.seek(0)
                st.success(f"✅ Arquivo filtrado com sucesso na aba **{sheet_active}**, mantendo apenas trabalhadores ativos!")
                st.download_button("📥 Baixar Arquivo de Trabalhadores Ativos", buffer.getvalue(), "sinale_trabalhadores_ativos.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Erro ao processar o arquivo: {e}")
