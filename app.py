elif menu_opcao == "ATUALIZAR DADOS":
    titulo_estilizado("Atualizar Dados do SINALE")
    aviso_sinale()
    
    sinale_file_upd = st.file_uploader("Selecione o arquivo do SINALE para atualizar (.xlsx)", type=["xlsx"], key="sinale_upd_upload")
    header_upd = st.number_input("Linha do cabeçalho no arquivo SINALE:", value=11, min_value=1, key="hdr_sinale_upd")
    
    if sinale_file_upd:
        try:
            wb_upd = load_workbook(sinale_file_upd)
            sheet_upd = st.selectbox("Escolha a aba a ser tratada:", wb_upd.sheetnames, key="sheet_sinale_upd")
            ws_u = wb_upd[sheet_upd]
            
            cabecalhos = {}
            for c_idx in range(1, ws_u.max_column + 1):
                val = ws_u.cell(row=header_upd, column=c_idx).value
                if val is not None:
                    cabecalhos[str(val).strip()] = c_idx
            
            lista_colunas = list(cabecalhos.keys())
            
            st.write("---")
            st.markdown(f"### ✏️ Configuração de Atualização na Aba: **{sheet_upd}**")
            
            if not lista_colunas:
                st.error(f"❌ Nenhuma coluna encontrada na linha de cabeçalho **{header_upd}**!")
            else:
                # --- BOTÕES DE SELECIONAR / DESMARCAR TODOS ---
                if "selectAllState" not in st.session_state:
                    st.session_state["selectAllState"] = True

                st.subheader("1. Seleção dos Registros a Atualizar")
                col_btn1, col_btn2, _ = st.columns([1, 1, 2])
                with col_btn1:
                    if st.button("✅ Selecionar Todos", key="btn_sel_all"):
                        st.session_state["selectAllState"] = True
                        st.rerun()
                with col_btn2:
                    if st.button("❌ Desmarcar Todos", key="btn_des_all"):
                        st.session_state["selectAllState"] = False
                        st.rerun()

                # --- CARREGAMENTO DE TODOS OS DADOS (INCLUINDO VAZIOS) ---
                dados_tabela = []
                for r in range(header_upd + 1, ws_u.max_row + 1):
                    row_data = {"Selecionar": st.session_state["selectAllState"]}
                    for col_name, c_idx in cabecalhos.items():
                        cell_val = ws_u.cell(row=r, column=c_idx).value
                        row_data[col_name] = cell_val
                    
                    row_data["_excel_row"] = r
                    dados_tabela.append(row_data)
                
                if not dados_tabela:
                    st.warning("Nenhum dado encontrado abaixo da linha de cabeçalho.")
                else:
                    df_upd_view = pd.DataFrame(dados_tabela)
                    cols_ordenadas = ["Selecionar"] + [c for c in lista_colunas if c in df_upd_view.columns] + ["_excel_row"]
                    df_upd_view = df_upd_view[[c for c in cols_ordenadas if c in df_upd_view.columns]]
                    
                    df_editado = st.data_editor(
                        df_upd_view,
                        column_config={
                            "Selecionar": st.column_config.CheckboxColumn("Atualizar?"),
                            "_excel_row": st.column_config.NumberColumn("Linha Excel", disabled=True)
                        },
                        disabled=[c for c in df_upd_view.columns if c not in ["Selecionar"]],
                        hide_index=True,
                        key="data_editor_upd"
                    )
                    
                    linhas_excel_alvo = df_editado[df_editado["Selecionar"] == True]["_excel_row"].tolist()
                    st.info(f"📊 **{len(linhas_excel_alvo)}** registro(s) selecionado(s) para atualização.")
                    
                    st.write("---")
                    
                    # --- 2. DEFINIÇÃO DO CAMPO E VALOR ---
                    st.subheader("2. Definir Campo e Novo Valor")
                    col_alvo_upd = st.selectbox("Selecione o campo (coluna) que deseja atualizar:", lista_colunas, key="col_alvo_upd")
                    novo_valor = st.text_input("Informe o novo valor para este campo:", key="novo_valor_input")
                    
                    st.write("---")
                    
                    if st.button("🚀 Processar Atualização", key="btn_proc_upd"):
                        if not linhas_excel_alvo:
                            st.error("⚠️ Selecione ao menos um registro na tabela acima antes de processar!")
                        else:
                            col_idx_excel = cabecalhos.get(col_alvo_upd)
                            if not col_idx_excel:
                                st.error(f"Coluna '{col_alvo_upd}' não encontrada no arquivo.")
                            else:
                                count_atualizados = 0
                                for excel_row in linhas_excel_alvo:
                                    ws_u.cell(row=excel_row, column=col_idx_excel, value=novo_valor)
                                    count_atualizados += 1
                                
                                buffer = io.BytesIO()
                                wb_upd.save(buffer)
                                st.session_state["upd_file_bytes"] = buffer.getvalue()
                                st.session_state["upd_success_msg"] = f"✅ {count_atualizados} registro(s) atualizado(s) com sucesso na coluna **{col_alvo_upd}**!"
                    
                    if "upd_file_bytes" in st.session_state:
                        st.success(st.session_state["upd_success_msg"])
                        st.download_button(
                            "📥 Baixar Arquivo SINALE Atualizado", 
                            st.session_state["upd_file_bytes"], 
                            "sinale_atualizado_dados.xlsx", 
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="dl_btn_upd"
                        )
        except Exception as e:
            st.error(f"Erro ao processar o arquivo para atualização: {e}")
