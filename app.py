# --- OPÇÃO 3: PESQUISA PARA REMIÇÃO ---
elif menu_opcao == "PESQUISA PARA REMIÇÃO":
    titulo_estilizado("Pesquisa para Remição")
    
    st.subheader("1. Configuração de Arquivos, Abas e Campos")
    uploaded_files = st.file_uploader("Selecione um ou mais arquivos (.xlsx, .xls, .ods)", type=["xlsx", "xls", "ods"], accept_multiple_files=True, key="search_upload")
    
    if uploaded_files:
        settings = {}
        for f in uploaded_files:
            f.seek(0)
            xl = pd.ExcelFile(f)
            sheets_available = xl.sheet_names
            
            with st.expander(f"📁 Configurações para: {f.name}", expanded=True):
                # Regra de seleção: Procura abas de remuneração. Se não achar nenhuma, pega estritamente APENAS A PRIMEIRA ABA, seja qual for o nome.
                pref_sheets = [s for s in sheets_available if any(p in s.strip().upper() for p in ["COM REMUNER", "SEM REMUNER"])]
                if pref_sheets:
                    default_sheets = pref_sheets
                else:
                    default_sheets = [sheets_available[0]] if sheets_available else []
                
                selected_sheets = st.multiselect(f"Selecione aba(s) para {f.name}", sheets_available, default=default_sheets, max_selections=2, key=f"sheets_{f.name}")
                
                sheet_config = {}
                for i, sheet in enumerate(selected_sheets):
                    st.markdown(f"**Aba: `{sheet}`**")
                    default_header = 11 if i == 0 else 10
                    header_row = st.number_input(f"Linha do cabeçalho para aba '{sheet}'", value=default_header, min_value=1, key=f"head_{f.name}_{sheet}")
                    
                    try:
                        f.seek(0)
                        df_preview = pd.read_excel(f, sheet_name=sheet, header=header_row-1, nrows=0)
                        cols_aba = list(df_preview.columns)
                    except:
                        cols_aba = []
                    
                    sheet_upper = sheet.strip().upper()
                    default_col = None
                    
                    if "COM REMUNER" in sheet_upper:
                        for c in cols_aba:
                            if str(c).strip().upper() == "NOME":
                                default_col = c
                                break
                        if not default_col and len(cols_aba) > 8:
                            default_col = cols_aba[8]
                    elif "SEM REMUNER" in sheet_upper:
                        for c in cols_aba:
                            if str(c).strip().upper() == "NOME DO INTERNO":
                                default_col = c
                                break
                        if not default_col and len(cols_aba) > 8:
                            default_col = cols_aba[8]
                    else:
                        for c in cols_aba:
                            if str(c).strip().upper() == "NOME":
                                default_col = c
                                break
                        if not default_col:
                            for c in cols_aba:
                                if "NOME" in str(c).strip().upper():
                                    default_col = c
                                    break
                        if not default_col and len(cols_aba) > 8:
                            default_col = cols_aba[8]
                        elif not default_col and cols_aba:
                            default_col = cols_aba[0]
                    
                    opcoes_colunas = ["--- Não pesquisar nesta aba ---"] + cols_aba
                    default_idx = opcoes_colunas.index(default_col) if default_col in opcoes_colunas else 0
                    
                    col_escolhida = st.selectbox(f"Selecione o campo (coluna) para a pesquisa na aba '{sheet}':", opcoes_colunas, index=default_idx, key=f"col_search_{f.name}_{sheet}")
                    
                    sheet_config[sheet] = {
                        "header_idx": header_row - 1,
                        "col_busca": col_escolhida if col_escolhida != "--- Não pesquisar nesta aba ---" else None
                    }
                    st.markdown("---")
                
                settings[f.name] = sheet_config
        
        if st.button("🔍 Carregar e Consolidar Dados para Pesquisa", key="btn_consolidar_op3"):
            all_results = []
            for f in uploaded_files:
                f.seek(0)
                xl = pd.ExcelFile(f)
                mes_ano_m9 = extrair_mes_ano_m9(f, xl.sheet_names)
                
                file_cfg = settings.get(f.name, {})
                for sheet, cfg in file_cfg.items():
                    try:
                        f.seek(0)
                        df_tmp = pd.read_excel(f, sheet_name=sheet, header=cfg["header_idx"])
                        target_col = cfg["col_busca"]
                        if target_col and target_col in df_tmp.columns:
                            df_tmp['MÊS/ANO - ABA'] = f"{mes_ano_m9} - {sheet}"
                            df_tmp['Aba Original'] = sheet  # Identificação exata da aba de origem
                            df_tmp['Campo Pesquisado'] = target_col
                            df_tmp['Valor_Busca'] = df_tmp[target_col].astype(str)
                            df_tmp['Nome (Visualização)'] = df_tmp[target_col].astype(str) + f" - {sheet}"
                            all_results.append(df_tmp)
                    except Exception as e:
                        st.error(f"Erro ao ler {f.name} - Aba {sheet}: {e}")
            
            if all_results:
                st.session_state['pesquisa_df'] = pd.concat(all_results, ignore_index=True)
                st.success(f"Dados consolidados com sucesso! **{len(st.session_state['pesquisa_df'])}** registros carregados.")
            else:
                st.warning("Nenhum dado encontrado com as configurações informadas.")
                st.session_state['pesquisa_df'] = None
    else:
        st.session_state['pesquisa_df'] = None

    if st.session_state.get('pesquisa_df') is not None:
        df_pesq = st.session_state['pesquisa_df']
        st.markdown("---")
        st.subheader("🔍 Filtros de Visualização e Busca")
        
        # 1. PRIMEIRO: Efetuar a pesquisa e seleção do nome
        nomes_disponiveis = sorted(df_pesq['Nome (Visualização)'].dropna().unique())
        nomes_selecionados = st.multiselect(
            "🔍 Digite para pesquisar e selecione o(s) nome(s):",
            options=nomes_disponiveis,
            key="busca_nomes_op3"
        )
        
        df_view = df_pesq.copy()
        if nomes_selecionados:
            df_view = df_view[df_view['Nome (Visualização)'].isin(nomes_selecionados)]
            
        st.metric("Total de Registros Encontrados", len(df_view))
        
        # 2. SEGUNDO: Exibir APENAS os campos da aba de origem onde o item selecionado foi encontrado
        if not df_view.empty:
            abas_presentes = df_view['Aba Original'].unique()
            
            colunas_disponiveis_aba = []
            cols_controle = ['MÊS/ANO - ABA', 'Nome (Visualização)', 'Campo Pesquisado']
            
            for aba in abas_presentes:
                df_aba_temp = df_view[df_view['Aba Original'] == aba]
                outras_cols = [c for c in df_aba_temp.columns if c not in cols_controle and c not in ['Valor_Busca', 'Aba Original']]
                for c in outras_cols:
                    if c not in colunas_disponiveis_aba:
                        colunas_disponiveis_aba.append(c)
            
            lista_colunas_full = cols_controle + colunas_disponiveis_aba
            
            st.info(f"💡 Aba(s) de origem identificada(s) para o(s) registro(s) selecionado(s): **{', '.join(abas_presentes)}**. Os campos abaixo pertencem estritamente a esta(s) aba(s).")
            cols_para_ver = st.multiselect("Selecione os campos para visualizar:", options=lista_colunas_full, default=cols_controle, key="cols_ver_op3")
            
            if cols_para_ver: 
                st.dataframe(df_view[cols_para_ver], use_container_width=True)
            else: 
                st.info("ℹ️ Selecione ao menos um campo acima para exibir a tabela de visualização.")
        else:
            st.info("ℹ️ Nenhum registro selecionado ou encontrado na pesquisa.")
