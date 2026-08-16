import io
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from copy import copy

st.set_page_config(page_title="Integrador Profissional", layout="wide")
st.title("⚡ Integrador: Seleção e Mapeamento")

# --- 1. CARREGAMENTO DOS ARQUIVOS ---
col1, col2 = st.columns(2)
with col1:
    source_file = st.file_uploader("1. Arquivo de ORIGEM", type=["xlsx", "xls", "csv", "txt"])
    origem_tem_cabecalho = st.checkbox("Origem tem cabeçalho?", value=True)
with col2:
    dest_file = st.file_uploader("2. Arquivo de DESTINO (.xlsx)", type=["xlsx"])
    header_dest = st.number_input("Linha do cabeçalho no Destino:", value=11, min_value=1)

if source_file and dest_file:
    # --- LEITURA DA ORIGEM ---
    if "source_df" not in st.session_state:
        hdr = 0 if origem_tem_cabecalho else None
        df_origem = pd.read_excel(source_file) if source_file.name.endswith('.xlsx') else pd.read_csv(source_file, header=hdr)
        st.session_state["source_df"] = df_origem
    df_origem = st.session_state["source_df"]

    # --- 2. SELEÇÃO DOS DADOS (Onde você escolhe o que transferir) ---
    st.subheader("3. Seleção de Dados da Origem")
    col_busca = st.selectbox("Coluna para identificar os registros:", df_origem.columns)
    
    # Cria uma lista formatada para o multiselect
    opcoes = [f"{val} (Linha {idx})" for idx, val in df_origem[col_busca].items()]
    selected_options = st.multiselect("🔍 Escolha os registros que deseja enviar:", opcoes)
    
    # Extrai os índices das linhas selecionadas
    selected_indices = [int(item.split("(Linha ")[1].replace(")", "")) for item in selected_options]

    # --- LEITURA DO DESTINO ---
    dest_file.seek(0)
    wb = load_workbook(io.BytesIO(dest_file.getvalue()), data_only=True)
    target_sheet = st.selectbox("4. Escolha a ABA de Destino:", wb.sheetnames)
    ws = wb[target_sheet]

    # --- 5. MAPEAMENTO POR COLUNA (Onde você diz para onde vai cada dado) ---
    st.subheader("5. Mapeamento de Colunas")
    max_col = ws.max_column
    mapping = {}
    
    cols_ui = st.columns(4)
    for i in range(1, max_col + 1):
        header_val = ws.cell(row=header_dest, column=i).value
        label = f"Col {i} ({header_val})" if header_val else f"Col {i} (Em branco)"
        
        with cols_ui[(i-1) % 4]:
            options = ["--- Não mapear ---", "⚠️ Auto-incrementar (Seq)"] + list(df_origem.columns)
            map_val = st.selectbox(label, options, key=f"map_{i}")
            if map_val != "--- Não mapear ---":
                mapping[i] = map_val

    # --- 6. PROCESSAMENTO ---
    if st.button("🚀 Processar e Atualizar Destino"):
        if not selected_indices:
            st.error("Selecione pelo menos um registro!")
        else:
            try:
                wb_write = load_workbook(io.BytesIO(dest_file.getvalue()))
                ws_write = wb_write[target_sheet]
                start_row = ws_write.max_row + 1
                
                # Pega a última sequência se for o caso
                try: seq_atual = int(ws_write.cell(row=start_row-1, column=1).value or 0)
                except: seq_atual = 0

                for idx in selected_indices:
                    for col_idx, origem_col in mapping.items():
                        if origem_col == "⚠️ Auto-incrementar (Seq)":
                            seq_atual += 1
                            ws_write.cell(row=start_row, column=col_idx, value=seq_atual)
                        else:
                            ws_write.cell(row=start_row, column=col_idx, value=df_origem.iloc[idx][origem_col])
                    start_row += 1

                buffer = io.BytesIO()
                wb_write.save(buffer)
                st.download_button("📥 Baixar Arquivo Atualizado", buffer.getvalue(), "destino_final.xlsx")
                st.success("✅ Tudo pronto!")
            except Exception as e:
                st.error(f"Erro: {e}")
