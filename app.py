import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Sistema de Gestão de Dados", layout="wide")
st.title("📊 Painel de Integração de Dados")

col1, col2 = st.columns(2)
with col1:
    source_file = st.file_uploader("1. Enviar Arquivo de Origem", type=["xlsx", "xls", "csv"])
with col2:
    dest_file = st.file_uploader("2. Enviar Arquivo de Destino (Planilha com várias abas)", type=["xlsx", "xls"])

def ler_arquivo(arquivo):
    if arquivo.name.lower().endswith(".csv"):
        return pd.read_csv(arquivo)
    return pd.read_excel(arquivo)

if source_file and dest_file:
    try:
        # Carrega Origem
        source_df = ler_arquivo(source_file)
        # Adiciona coluna de seleção se não existir
        if "Selecionar" not in source_df.columns:
            source_df.insert(0, "Selecionar", False)
        
        # Carrega Destino
        dest_xls = pd.ExcelFile(dest_file)
        
        st.success("Arquivos carregados!")
        
        # --- Seleção da Aba ---
        selected_sheet = st.selectbox("3. Escolha a Aba (Planilha) de Destino:", dest_xls.sheet_names)
        dest_df = pd.read_excel(dest_file, sheet_name=selected_sheet)
        all_dest_dfs = pd.read_excel(dest_file, sheet_name=None)
        
        # --- Mapeamento Inteligente ---
        st.subheader("4. Mapeamento (Qual coluna da Origem vai para o Destino?)")
        mapping = {}
        for dest_col in dest_df.columns:
            # Mostra o nome da coluna e um exemplo do dado abaixo
            opcoes = source_df.columns.tolist()
            mapping[dest_col] = st.selectbox(
                f"O campo '{dest_col}' do Destino recebe da Origem:", 
                options=opcoes, 
                key=f"map_{dest_col}"
            )
            
        # --- Busca e Seleção Automática ---
        st.subheader("5. Selecionar Dados")
        search = st.text_input("🔍 Pesquisar na origem (filtra automaticamente):", "")
        
        # Filtrar o df para exibição
        df_display = source_df.copy()
        if search:
            # Procura em todas as colunas de texto
            mask = df_display.apply(lambda row: row.astype(str).str.contains(search, case=False).any(), axis=1)
            df_display = df_display[mask]
            
        # Edição dos dados selecionados
        selected_df = st.data_editor(df_display, use_container_width=True, hide_index=True)
        
        # Filtrar apenas o que foi marcado com o checkbox "Selecionar"
        final_selected = selected_df[selected_df["Selecionar"] == True]
        st.write(f"Você selecionou **{len(final_selected)}** registros para transferir.")
        
        # --- Inserção ---
        st.subheader("6. Finalizar")
        mode = st.radio("Onde salvar?", ["Final do arquivo", "Em uma linha específica"])
        target_line = 0
        if mode == "Em uma linha específica":
            target_line = st.number_input("Número da linha (0 é o topo):", min_value=0, max_value=len(dest_df))
            
        if st.button("🚀 Processar e Gerar Novo Arquivo"):
            if len(final_selected) == 0:
                st.error("Nenhum dado foi selecionado! Marque a caixinha 'Selecionar' na tabela.")
            else:
                # Prepara os dados (limpa a coluna Selecionar)
                new_data = final_selected[list(mapping.values())].copy()
                new_data.columns = list(mapping.keys())
                
                # Insere
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
                        sh_df.to_excel(writer, sheet_name=sh_name, index=False)
                
                st.success("Tudo pronto!")
                st.download_button("📥 Baixar Planilha Atualizada", data=buffer.getvalue(), file_name="arquivo_final.xlsx")

    except Exception as e:
        st.error(f"Ocorreu um erro: {e}")
