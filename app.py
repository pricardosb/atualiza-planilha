import io
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from copy import copy

st.set_page_config(page_title="Integrador Profissional", layout="wide")
st.title("⚡ Integrador: Mapeamento Preciso por Coluna")

# --- 1. CARREGAMENTO ---
col1, col2 = st.columns(2)
with col1:
    source_file = st.file_uploader("1. Origem (Excel, CSV, TXT)", type=["xlsx", "xls", "csv", "txt"])
    origem_tem_cabecalho = st.checkbox("Origem tem cabeçalho?", value=True)
with col2:
    dest_file = st.file_uploader("2. Destino (.xlsx)", type=["xlsx"])
    header_dest = st.number_input("Linha do cabeçalho no Destino:", value=11, min_value=1)

if source_file and dest_file:
    # --- LEITURA ORIGEM ---
    # ... (mesma lógica cacheada de antes para performance) ...
    if "source_df" not in st.session_state:
        hdr = 0 if origem_tem_cabecalho else None
        # (Lógica de leitura simplificada para foco na escrita)
        df_origem = pd.read_excel(source_file) if source_file.name.endswith('.xlsx') else pd.read_csv(source_file, header=hdr)
        st.session_state["source_df"] = df_origem
    df_origem = st.session_state["source_df"]

    # --- LEITURA DESTINO ---
    dest_file.seek(0)
    wb = load_workbook(io.BytesIO(dest_file.getvalue()), data_only=True)
    target_sheet = st.selectbox("3. Escolha a ABA:", wb.sheetnames)
    ws = wb[target_sheet]

    # --- MAPEAMENTO POR ÍNDICE (A CHAVE DE TUDO) ---
    st.subheader("4. Mapeamento (Coluna a Coluna)")
    max_col = ws.max_column
    mapping = {}
    
    # Criamos colunas na UI para organizar melhor
    cols_ui = st.columns(4)
    for i in range(1, max_col + 1):
        # Pega o valor real na linha 11, se estiver em branco, chama de "Coluna N"
        header_val = ws.cell(row=header_dest, column=i).value
        label = f"Coluna {i} ({header_val})" if header_val else f"Coluna {i} (Em branco)"
        
        with cols_ui[(i-1) % 4]:
            options = ["--- Não mapear ---", "⚠️ Auto-incrementar (Seq)"] + list(df_origem.columns)
            map_val = st.selectbox(label, options, key=f"map_{i}")
            if map_val != "--- Não mapear ---":
                mapping[i] = map_val # O índice 'i' é a coluna exata do Excel

    # --- SELEÇÃO DE REGISTROS ---
    st.subheader("5. Seleção de Dados")
    # (Adicione aqui a lógica de multiselect que você já usava)
    indices_selecionados = [0, 1, 2] # Exemplo, assumindo que você já selecionou

    if st.button("🚀 Processar e Atualizar"):
        # --- ESCRITA FINAL ---
        wb_write = load_workbook(io.BytesIO(dest_file.getvalue()))
        ws_write = wb_write[target_sheet]
        
        # Define a linha de início (ajuste conforme seu 'modo_insercao')
        start_row = ws_write.max_row + 1
        
        # Encontra a sequência atual da coluna 1 (se for numérica)
        try:
            seq_atual = int(ws_write.cell(row=start_row-1, column=1).value or 0)
        except:
            seq_atual = 0

        for idx_origem in indices_selecionados:
            # Escreve baseado no mapeamento
            for col_idx, origem_col in mapping.items():
                if origem_col == "⚠️ Auto-incrementar (Seq)":
                    seq_atual += 1
                    ws_write.cell(row=start_row, column=col_idx, value=seq_atual)
                else:
                    valor = df_origem.iloc[idx_origem][origem_col]
                    ws_write.cell(row=start_row, column=col_idx, value=valor)
            
            start_row += 1

        # --- DOWNLOAD ---
        buffer = io.BytesIO()
        wb_write.save(buffer)
        st.download_button("📥 Baixar Arquivo Atualizado", buffer.getvalue(), "destino_final.xlsx")
        st.success("✅ Atualização feita respeitando exatamente suas colunas!")
