import io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Sistema de Gestão de Dados", layout="wide"
)
st.title("📊 Sistema de Integração de Dados")

col1, col2 = st.columns(2)
with col1:
  source_file = st.file_uploader(
      "1. Enviar Arquivo de Origem (Relatório Bruto)",
      type=["xlsx", "xls", "csv"],
  )
with col2:
  dest_file = st.file_uploader(
      "2. Enviar Arquivo de Destino (Seu Controle com Várias Abas)",
      type=["xlsx", "xls"],
  )


def ler_origem(arquivo):
  nome = arquivo.name.lower()
  if nome.endswith(".csv"):
    return pd.read_csv(arquivo)
  else:
    return pd.read_excel(arquivo)


if source_file and dest_file:
  try:
    source_df = ler_origem(source_file)

    # Lê todas as abas disponíveis no arquivo de destino
    dest_xls = pd.ExcelFile(dest_file)
    dest_sheets = dest_xls.sheet_names

    st.success("Arquivos carregados com sucesso!")

    # --- Seleção da Aba de Destino ---
    st.subheader("3. Escolha a Planilha (Aba) de Destino")
    selected_sheet = st.selectbox(
        "Em qual aba do arquivo de destino você deseja inserir os dados?",
        options=dest_sheets,
    )

    # Carrega a aba escolhida para o mapeamento
    dest_df = pd.read_excel(dest_file, sheet_name=selected_sheet)

    # Carrega TODAS as abas para um dicionário (para não perder as outras abas ao salvar)
    all_dest_dfs = pd.read_excel(dest_file, sheet_name=None)

    # --- Mapeamento (De/Para) ---
    st.subheader(
        f"4. Mapeamento de Colunas (Aba: '{selected_sheet}'- Destino x Origem)"
    )
    mapping = {}
    for dest_col in dest_df.columns:
      mapping[dest_col] = st.selectbox(
          f"Destino '{dest_col}' corresponde a:",
          options=source_df.columns,
          key=dest_col,
      )

    # --- Busca e Seleção ---
    st.subheader("5. Seleção de Dados")
    search = st.text_input("Filtrar por nome na Origem:", "")
    filtered_df = source_df.copy()
    if search:
      name_cols = [c for c in source_df.columns if "nome" in c.lower()]
      if name_cols:
        filtered_df = filtered_df[
            filtered_df[name_cols[0]].str.contains(search, case=False, na=False)
        ]

    selected_df = st.data_editor(filtered_df, use_container_width=True)

    # --- Inserção ---
    st.subheader("6. Inserção no Destino")
    mode = st.radio("Onde inserir?", ["Final do arquivo", "Linha específica"])
    target_line = 0
    if mode == "Linha específica":
      target_line = st.number_input(
          "Número da linha (0 é o topo):",
          min_value=0,
          max_value=len(dest_df),
      )

    if st.button("Executar Integração"):
      new_data = selected_df[list(mapping.values())].copy()
      new_data.columns = list(mapping.keys())

      # Pega os dados atuais da aba escolhida
      current_dest_df = all_dest_dfs[selected_sheet]

      # Insere no final ou na linha específica
      if mode == "Final do arquivo":
        final_sheet_df = pd.concat(
            [current_dest_df, new_data], ignore_index=True
        )
      else:
        final_sheet_df = pd.concat(
            [
                current_dest_df.iloc[:target_line],
                new_data,
                current_dest_df.iloc[target_line:],
            ],
            ignore_index=True,
        )

      # Atualiza apenas a aba alterada dentro do conjunto de abas
      all_dest_dfs[selected_sheet] = final_sheet_df

      # Salva o arquivo de volta preservando TODAS as abas originais do seu controle
      buffer = io.BytesIO()
      with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sh_name, sh_df in all_dest_dfs.items():
          sh_df.to_excel(writer, sheet_name=sh_name, index=False)

      st.success(f"Dados integrados com sucesso na aba '{selected_sheet}'!")
      st.download_button(
          "📥 Baixar Arquivo Atualizado (Com Todas as Abas)",
          data=buffer.getvalue(),
          file_name="arquivo_atualizado.xlsx",
      )
  except Exception as e:
    st.error(f"Erro ao processar: {e}")
