import io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Sistema de Gestão de Dados", layout="wide"
)
st.title("⚡ Painel de Integração Instantânea")

col1, col2 = st.columns(2)
with col1:
  source_file = st.file_uploader(
      "1. Enviar Arquivo de Origem", type=["xlsx", "xls", "csv"]
  )
with col2:
  dest_file = st.file_uploader(
      "2. Enviar Arquivo de Destino (Com cabeçalho na linha 11)",
      type=["xlsx", "xls"],
  )


# --- CACHE: Guarda os arquivos na memória para a busca ser ULTRA RÁPIDA ---
@st.cache_data(show_spinner=False)
def carregar_origem(file):
  if file.name.lower().endswith(".csv"):
    df = pd.read_csv(file)
  else:
    df = pd.read_excel(file)
  if "Selecionar" not in df.columns:
    df.insert(0, "Selecionar", False)
  return df


@st.cache_data(show_spinner=False)
def carregar_destino(file):
  xls = pd.ExcelFile(file)
  all_sheets = {
      sheet: pd.read_excel(file, sheet_name=sheet, header=10)
      for sheet in xls.sheet_names
  }
  return xls.sheet_names, all_sheets


if source_file and dest_file:
  try:
    # Carrega direto da memória (instantâneo)
    source_df = carregar_origem(source_file).copy()
    sheet_names, all_dest_dfs = carregar_destino(dest_file)

    st.success("⚡ Arquivos prontos em memória!")

    # --- Seleção da Aba ---
    selected_sheet = st.selectbox(
        "3. Escolha a Aba (Planilha) de Destino:", sheet_names
    )
    dest_df = all_dest_dfs[selected_sheet]

    # --- Mapeamento ---
    st.subheader("4. Mapeamento (Qual coluna da Origem vai para o Destino?)")
    mapping = {}
    for dest_col in dest_df.columns:
      if "Unnamed" in str(dest_col):
        continue
      opcoes = [c for c in source_df.columns if c != "Selecionar"]
      mapping[dest_col] = st.selectbox(
          f"O campo '{dest_col}' do Destino recebe da Origem:",
          options=opcoes,
          key=f"map_{dest_col}",
      )

    # --- Busca Instantânea ---
    st.subheader("5. Selecionar Dados da Origem")
    search = st.text_input("⚡ Pesquisa Instantânea (digite para filtrar):", "")

    # Filtro super rápido diretamente na memória RAM
    if search:
      mask = source_df.astype(str).apply(
          lambda row: row.str.contains(search, case=False, regex=False).any(),
          axis=1,
      )
      df_display = source_df[mask]
    else:
      df_display = source_df

    selected_df = st.data_editor(
        df_display, use_container_width=True, hide_index=True
    )

    final_selected = selected_df[selected_df["Selecionar"] == True]
    st.write(f"📌 Registros selecionados: **{len(final_selected)}**")

    # --- Inserção ---
    st.subheader("6. Finalizar")
    mode = st.radio(
        "Onde salvar?", ["Final do arquivo", "Em uma linha específica"]
    )
    target_line = 0
    if mode == "Em uma linha específica":
      target_line = st.number_input(
          "Número da linha (após o cabeçalho):",
          min_value=0,
          max_value=len(dest_df),
      )

    if st.button("🚀 Processar e Gerar Novo Arquivo"):
      if len(final_selected) == 0:
        st.error("Marque a caixinha 'Selecionar' na tabela de origem!")
      else:
        cols_origem = list(mapping.values())
        new_data = final_selected[cols_origem].copy()
        new_data.columns = list(mapping.keys())

        current_dest = all_dest_dfs[selected_sheet].copy()

        if mode == "Final do arquivo":
          updated_sheet = pd.concat(
              [current_dest, new_data], ignore_index=True
          )
        else:
          updated_sheet = pd.concat(
              [
                  current_dest.iloc[:target_line],
                  new_data,
                  current_dest.iloc[target_line:],
              ],
              ignore_index=True,
          )

        output_dfs = all_dest_dfs.copy()
        output_dfs[selected_sheet] = updated_sheet

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
          for sh_name, sh_df in output_dfs.items():
            sh_df.to_excel(writer, sheet_name=sh_name, index=False, startrow=10)

        st.success("Tudo pronto!")
        st.download_button(
            "📥 Baixar Planilha Atualizada",
            data=buffer.getvalue(),
            file_name="arquivo_final.xlsx",
        )

  except Exception as e:
    st.error(f"Erro: {e}")
