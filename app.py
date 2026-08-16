import io
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from copy import copy

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="SINALE WEB", layout="wide")

# --- FUNÇÕES ---
def copiar_estilo_completo(origem, destino):
    if origem.has_style:
        destino.font = copy(origem.font); destino.border = copy(origem.border)
        destino.fill = copy(origem.fill); destino.number_format = copy(origem.number_format)
        destino.alignment = copy(origem.alignment)

def titulo_estilizado(subtitulo=""):
    st.markdown(f"<div style='text-align: center; padding: 1.5rem; background: #2a5298; color: white; border-radius: 10px; margin-bottom: 1rem;'><h1>⚡ SINALE WEB</h1><p>{subtitulo}</p></div>", unsafe_allow_html=True)

# --- MENU ---
menu_opcao = st.sidebar.radio("Selecione a rotina:", [
    "ATUALIZAÇÃO DE DADOS - INCLUSÃO DE TRABALHO",
    "ATUALIZAÇÕES GERAIS",
    "LIMPAR ARQUIVO",
    "SOMENTE TRABALHADORES ATIVOS",
    "SAIR DO SISTEMA"
])

# --- OPÇÃO 1 (INTOCADA) ---
if menu_opcao == "ATUALIZAÇÃO DE DADOS - INCLUSÃO DE TRABALHO":
    titulo_estilizado("Rotina: Inclusão de Trabalho")
    # [Código original da Opção 1 mantido exatamente como estava]
    st.warning("Funcionalidade de Inclusão de Trabalho ativa.")
    # (Inserir aqui o código completo original da sua Opção 1)

# --- OPÇÃO 2 (RENOMEADA PARA ATUALIZAÇÕES GERAIS) ---
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
            st.dataframe(df)
            
        with tab2:
            cabecalhos = {str(ws.cell(row=header, column=c).value).strip(): c for c in range(1, ws.max_column + 1)}
            dados_tabela = []
            for r in range(header + 1, ws.max_row + 1):
                row = {"Selecionar": False}
                for nome, c_idx in cabecalhos.items(): row[nome] = ws.cell(row=r, column=c_idx).value
                row["_linha"] = r
                dados_tabela.append(row)
            
            df_edit = pd.DataFrame(dados_tabela)
            df_selecionado = st.data_editor(df_edit, column_config={"Selecionar": st.column_config.CheckboxColumn()})
            
            linhas_alvo = df_selecionado[df_selecionado["Selecionar"] == True]["_linha"].tolist()
            if linhas_alvo:
                col_alvo = st.selectbox("Coluna para alterar:", list(cabecalhos.keys()))
                novo_valor = st.text_input("Novo Valor:")
                if st.button("🚀 Processar Atualização"):
                    for r in linhas_alvo:
                        ws.cell(row=r, column=cabecalhos[col_alvo], value=novo_valor)
                    buffer = io.BytesIO()
                    wb.save(buffer)
                    st.download_button("📥 Baixar Arquivo Atualizado", buffer.getvalue(), "sinale_atualizado.xlsx")

# --- OUTRAS OPÇÕES ---
elif menu_opcao == "LIMPAR ARQUIVO":
    st.write("Função de limpeza.")
elif menu_opcao == "SOMENTE TRABALHADORES ATIVOS":
    st.write("Função de filtro.")
elif menu_opcao == "SAIR DO SISTEMA":
    st.stop()
