import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Sistema de Gestão de Dados", layout="wide")
st.title("📊 Sistema de Integração de Dados")

col1, col2 = st.columns(2)
with col1:
    source_file = st.file_uploader("1. Enviar Arquivo de Origem (Relatório Bruto)", type=["xlsx", "xls", "csv"])
with col2:
    dest_file = st.file_uploader("2. Enviar Arquivo de Destino (Seu Controle)", type=["xlsx", "xls", "csv"])

if source_file and dest_file:
    try:
        source_df = pd.read_excel(source_file) if source_file.name.endswith('.xlsx') else pd.read_csv(source_file)
        dest_df = pd.read_excel(dest_file) if dest_file.name.endswith('.xlsx') else pd.read_csv(dest_file)
        st.success("Arquivos carregados!")

        st.subheader("3. Mapeamento de Colunas (De/Para)")
        mapping = {}
        for dest_col in dest_df.columns:
            mapping[dest_col] = st.selectbox(f"Destino: '{dest_col}' corresponde a:", options=source_df.columns, key=dest_col)

        st.subheader("4. Seleção de Dados")
        search = st.text_input("Filtrar por nome na Origem:", "")
        filtered_df = source_df.copy()
        if search:
            name_cols = [c for c in source_df.columns if 'nome' in c.lower()]
            if name_cols: filtered_df = filtered_df[filtered_df[name_cols[0]].str.contains(search, case=False, na=False)]

        selected_df = st.data_editor(filtered_df, use_container_width=True)

        st.subheader("5. Inserção no Destino")
        mode = st.radio("Onde inserir?", ["Final do arquivo", "Linha específica"])
        target_line = 0
        if mode == "Linha específica":
            target_line = st.number_input("Número da linha (0 é o topo):", min_value=0, max_value=len(dest_df))

        if st.button("Executar Integração"):
            new_data = selected_df[list(mapping.values())].copy()
            new_data.columns = list(mapping.keys())
            if mode == "Final do arquivo":
                final_df = pd.concat([dest_df, new_data], ignore_index=True)
            else:
                final_df = pd.concat([dest_df.iloc[:target_line], new_data, dest_df.iloc[target_line:]], ignore_index=True)

            buffer = io.BytesIO()
            final_df.to_excel(buffer, index=False)
            st.download_button("📥 Baixar Arquivo Atualizado", data=buffer.getvalue(), file_name="arquivo_atualizado.xlsx")
    except Exception as e:
        st.error(f"Erro ao processar: {e}")