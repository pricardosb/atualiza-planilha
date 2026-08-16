import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Sistema de Gestão de Dados", layout="wide")
st.title("📊 Painel de Integração de Dados")

col1, col2 = st.columns(2)
with col1:
    source_file = st.file_uploader("1. Enviar Arquivo de Origem", type=["xlsx", "xls", "csv"])
with col2:
    dest_file = st.file_uploader("2. Enviar Arquivo de Destino (Com cabeçalho na linha 11)", type=["xlsx", "xls"])

def ler_origem(arquivo):
    if arquivo.name.lower().endswith(".csv"):
        return pd.read_csv(arquivo)
    return pd.read_excel(arquivo)

if source_file and dest_file:
    try:
        # Carrega Origem
        source_df = ler_origem(source_file)
        if "Selecionar" not in source_df.columns:
            source_df.insert(0, "Selecionar", False)
        
        # Carrega Destino (Ajustado para ler cabeçalho na linha 11 -> header=10)
        dest_xls = pd.ExcelFile(dest_file)
        st.success("Arquivos carregados!")
        
        # --- Seleção da Aba ---
        selected_sheet = st.selectbox("3. Escolha a Aba (Planilha) de Destino:", dest_xls.sheet_names)
        
        # Carrega a aba escolhida COM O CABEÇALHO NA LINHA 11
        dest_df = pd.read_excel(dest_file, sheet_name=selected_sheet, header=10)
        
        # Carrega o conjunto para salvar depois (Mantendo o cabeçalho na linha 11)
        all_dest_dfs = pd.read_excel(dest_file, sheet_name=None, header=10)
        
        # --- Mapeamento ---
        st.subheader("4. Mapeamento (Qual coluna da Origem vai para o Destino?)")
        mapping = {}
        for dest_col in dest_df.columns:
            # Pula colunas vazias que o pandas pode ter criado
            if "Unnamed" in str(dest_col): continue
            
            opcoes = source_df.columns.tolist()
            mapping[dest_col] = st.selectbox(
                f"O campo '{dest_col}' do Destino recebe da Origem:", 
                options=opcoes, 
                key=f"map_{dest_col}"
            )
            
        # --- Seleção ---
        st.subheader("5. Selecionar Dados da Origem")
        search = st.text_input("🔍 Pesquisar na origem:", "")
        df_display = source_df.copy()
        if search:
            mask = df_display.apply(lambda row: row.astype(str).str.contains(search, case=False).any(), axis=1)
            df_display = df_display[mask]
            
        selected_df = st.data_editor(df_display, use_container_width=True, hide_index=True)
        final_selected = selected_df[selected_df["Selecionar"] == True]
        st.write(f"Você selecionou **{len(final_selected)}** registros.")
        
        # --- Inserção ---
        st.subheader("6. Finalizar")
        mode = st.radio("Onde salvar?", ["Final do arquivo", "Em uma linha específica"])
        target_line = 0
        if mode == "Em uma linha específica":
            target_line = st.number_input("Número da linha (relativo ao início dos dados após o cabeçalho):", min_value=0, max_value=len(dest_df))
            
        if st.button("🚀 Processar e Gerar Novo Arquivo"):
            if len(final_selected) == 0:
                st.error("Marque a caixinha 'Selecionar' na tabela de origem!")
            else:
                new_data = final_selected[list(mapping.values())].copy()
                new_data.columns = list(mapping.keys())
                
                current_dest = all_dest_dfs[selected_sheet]
                
                if mode == "Final do arquivo":
                    updated_sheet = pd.concat([current_dest, new_data], ignore_index=True)
                else:
                    updated_sheet = pd.concat([current_dest.iloc[:target_line], new_data, current_dest.iloc[target_line:]], ignore_index=True)
                
                all_dest_dfs[selected_sheet] = updated_sheet
                
                # Salva
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                    for sh_name, sh_df in all_dest_dfs.items():
                        # O startrow=10 garante que os dados comecem na linha 11 ao salvar
                        sh_df.to_excel(writer, sheet_name=sh_name, index=False, startrow=10)
                
                st.success("Tudo pronto!")
                st.download_button("📥 Baixar Planilha Atualizada", data=buffer.getvalue(), file_name="arquivo_final.xlsx")

    except Exception as e:
        st.error(f"Erro: {e}")
