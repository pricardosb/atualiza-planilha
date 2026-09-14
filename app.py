import streamlit as st
import pandas as pd
import numpy as np
import re
import io
import datetime
import calendar
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font
from copy import copy
import streamlit.components.v1 as components

# --- IMPORTAÇÕES DO GOOGLE DRIVE (EMBUTIDAS) ---
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

# =============================================================================
# --- CONEXÃO E FUNÇÕES DO GOOGLE DRIVE ---
# =============================================================================
SCOPES = ['https://www.googleapis.com/auth/drive']

def obter_servico_drive():
    """Autentica e retorna o serviço da API do Google Drive usando st.secrets."""
    try:
        credentials_info = st.secrets["gcp_service_account"]
        creds = Credentials.from_service_account_info(credentials_info, scopes=SCOPES)
        service = build('drive', 'v3', credentials=creds)
        return service
    except Exception as e:
        st.error(f"❌ Erro ao autenticar no Google Drive. Verifique o st.secrets. Detalhes: {e}")
        return None

def listar_arquivos_da_pasta(folder_id):
    """Lista todos os arquivos dentro de uma pasta específica no Drive."""
    service = obter_servico_drive()
    if not service:
        return []
    
    try:
        query = f"'{folder_id}' in parents and trashed = false"
        results = service.files().list(
            q=query,
            pageSize=1000,
            fields="nextPageToken, files(id, name, mimeType)"
        ).execute()
        
        items = results.get('files', [])
        return items
    except Exception as e:
        st.error(f"❌ Erro ao listar arquivos da pasta {folder_id}: {e}")
        return []

def baixar_arquivo_bytes(file_id: str) -> bytes:
    """Baixa o conteúdo de um arquivo do Drive para a memória (bytes)."""
    service = obter_servico_drive()
    if not service:
        return b""
        
    try:
        request = service.files().get_media(fileId=file_id)
        file_stream = io.BytesIO()
        downloader = MediaIoBaseDownload(file_stream, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()
        return file_stream.getvalue()
    except Exception as e:
        st.error(f"❌ Erro ao baixar o arquivo {file_id}: {e}")
        return b""

def salvar_arquivo_no_drive(nome_arquivo, bytes_conteudo, mime_type, folder_id):
    """Salva um novo arquivo no Google Drive e retorna o link de visualização."""
    service = obter_servico_drive()
    if not service:
        return ""
        
    try:
        file_metadata = {
            'name': nome_arquivo,
            'parents': [folder_id]
        }
        media = MediaIoBaseUpload(io.BytesIO(bytes_conteudo), mimetype=mime_type, resumable=True)
        
        file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id, webViewLink'
        ).execute()
        
        return file.get('webViewLink', f"https://drive.google.com/file/d/{file.get('id')}/view")
    except Exception as e:
        st.error(f"❌ Erro ao salvar o arquivo no Drive: {e}")
        return ""

def eh_arquivo_valido(item):
    """Filtra pastas válidas e ignora arquivos temporários, ocultos e de sistema."""
    nome = item.get("name", "")
    if not nome:
        return False
    if nome.lower() == 'credentials.json':
        return False
    if item.get('mimeType') == 'application/vnd.google-apps.folder':
        return True
    if nome.lower() in ('thumbs.db', 'desktop.ini'):
        return False
    if nome.startswith(('.~', '~$', '.')):
        return False
    if '.' not in nome:
        return False
    return True


# =============================================================================
# --- FUNÇÕES GERAIS E DE SUPORTE ---
# =============================================================================

def tentar_converter_numero(val):
    """Converte texto numérico em int/float nativo para o Excel reconhecer como número."""
    if pd.isna(val) or val == "" or val is None:
        return ""
    if isinstance(val, (int, float)):
        return val
    val_str = str(val).strip().replace(',', '.')
    try:
        num = float(val_str)
        return int(num) if num.is_integer() else num
    except (ValueError, TypeError):
        return str(val)

def limpar_texto_xml(texto):
    """Remove caracteres inválidos de controle ASCII que corrompem documentos Word (.docx)."""
    if pd.isna(texto) or texto is None:
        return ""
    texto_str = str(texto)
    return re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', texto_str)

def gerar_excel_bytes(dados_exportacao):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Salvamento Remição"

    for item in dados_exportacao:
        ws.append([f"NOME: {item['nome']}"])
        ws.append([f"ORGANIZAÇÃO: {item['organiz']} | FUNÇÃO: {item['funcao']} | REMUNERAÇÃO: {item['remuneracao']} | SAÍDA: {item['saida']}"])
        ws.append([])

        pivot_df = item['pivot_df']
        if not pivot_df.empty:
            headers = ["ANO"] + list(pivot_df.columns)
            ws.append(headers)
            for idx_row, row_data in pivot_df.iterrows():
                row_vals = [tentar_converter_numero(idx_row)] + [tentar_converter_numero(v) for v in row_data.values]
                ws.append(row_vals)

        ws.append(["Total de Dias:", tentar_converter_numero(item['total_dias'])])
        ws.append([])
        ws.append([])

    wb.save(output)
    return output.getvalue()

def gerar_docx_bytes(dados_exportacao):
    output = io.BytesIO()
    from docx import Document
    
    doc = Document()
    doc.add_heading("Espaço de Dados para Salvamento", level=1)

    for item in dados_exportacao:
        doc.add_heading(limpar_texto_xml(f"NOME: {item['nome']}"), level=2)
        p_meta = doc.add_paragraph()
        p_meta.add_run(
            limpar_texto_xml(
                f"ORGANIZAÇÃO: {item['organiz']} | "
                f"FUNÇÃO: {item['funcao']} | "
                f"REMUNERAÇÃO: {item['remuneracao']} | "
                f"SAÍDA: {item['saida']}"
            )
        )

        pivot_df = item['pivot_df']
        if not pivot_df.empty:
            headers = ["ANO"] + list(pivot_df.columns)
            table = doc.add_table(rows=1, cols=len(headers))
            table.style = 'Table Grid'
            
            hdr_cells = table.rows[0].cells
            for i, h in enumerate(headers):
                hdr_cells[i].text = limpar_texto_xml(h)

            for idx_row, row_data in pivot_df.iterrows():
                row_cells = table.add_row().cells
                row_cells[0].text = limpar_texto_xml(idx_row)
                for i, val in enumerate(row_data.values):
                    row_cells[i+1].text = limpar_texto_xml(val)

        p_tot = doc.add_paragraph()
        p_tot.add_run(limpar_texto_xml(f"Total de Dias: {item['total_dias']}")).bold = True
        doc.add_paragraph()

    doc.save(output)
    return output.getvalue()

def extrair_mes_ano_do_nome(nome_arquivo):
    import re
    
    meses = {
        "JANEIRO": "01", "FEVEREIRO": "02", "MARÇO": "03", "MARCO": "03",
        "ABRIL": "04", "MAIO": "05", "JUNHO": "06", "JULHO": "07",
        "AGOSTO": "08", "SETEMBRO": "09", "OUTUBRO": "10",
        "NOVEMBRO": "11", "DEZEMBRO": "12"
    }
    
    nome_upper = str(nome_arquivo).upper()
    ano_match = re.search(r'\b(20\d{2})\b', nome_upper)
    ano = ano_match.group(1) if ano_match else None
    
    mes = None
    for nome_mes, num_mes in meses.items():
        if nome_mes in nome_upper:
            mes = num_mes
            break
            
    if mes and ano:
        return f"{mes}/{ano}"
    
    return "SEM MÊS/ANO"

def copiar_estilo_completo(origem, destino):
    if origem.has_style:
        destino.font = copy(origem.font)
        destino.border = copy(origem.border)
        destino.fill = copy(origem.fill)
        destino.number_format = copy(origem.number_format)
        destino.protection = copy(origem.protection)
        destino.alignment = copy(origem.alignment)

def deduplicar_colunas(colunas):
    vistos = {}
    novas_colunas = []
    for col in colunas:
        col_str = str(col).strip()
        if col_str in vistos:
            vistos[col_str] += 1
            novas_colunas.append(f"{col_str} ({vistos[col_str]})")
        else:
            vistos[col_str] = 1
            novas_colunas.append(col_str)
    return novas_colunas

def extrair_valor_limpo(df, idx, col_name):
    try:
        val = df.iloc[idx][col_name]
        if isinstance(val, pd.Series):
            val = val.iloc[0]
        if pd.isna(val):
            return None
        return val.item() if hasattr(val, 'item') else val
    except:
        return None

def converter_valor_inteligente(val_str, dtype_original):
    if val_str is None or str(val_str).strip() == "":
        return None
    val_str = str(val_str).strip()
    if pd.api.types.is_integer_dtype(dtype_original):
        try:
            return int(val_str)
        except ValueError:
            pass
    elif pd.api.types.is_float_dtype(dtype_original):
        try:
            return float(val_str.replace(',', '.'))
        except ValueError:
            pass
    try:
        return float(val_str.replace(',', '.'))
    except ValueError:
        return val_str

def formatar_datas_dataframe(df_input):
    df_out = df_input.copy()
    for col in df_out.columns:
        if pd.api.types.is_datetime64_any_dtype(df_out[col]):
            df_out[col] = df_out[col].dt.strftime('%d/%m/%Y').fillna('')
        else:
            df_out[col] = df_out[col].apply(
                lambda v: "" if pd.isna(v) else (
                    v.strftime('%d/%m/%Y') if isinstance(v, (datetime.datetime, datetime.date, pd.Timestamp))
                    else (str(v).split(' ')[0] if isinstance(v, str) and (' 00:00:00' in str(v) or 'T00:00:00' in str(v)) else v)
                )
            )
    return df_out

def calcular_pascoa(ano):
    a = ano % 19
    b = ano // 100
    c = ano % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    L = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * L) // 451
    mes = (h + L - 7 * m + 114) // 31
    dia = ((h + L - 7 * m + 114) % 31) + 1
    return datetime.date(ano, mes, dia)

def obter_estatisticas_mes(ano, mes):
    cal = calendar.monthcalendar(ano, mes)
    pascoa = calcular_pascoa(ano)
    feriados = [
        datetime.date(ano, 1, 1),
        pascoa - datetime.timedelta(days=47),
        pascoa - datetime.timedelta(days=2),
        datetime.date(ano, 4, 21),
        datetime.date(ano, 5, 1),
        pascoa + datetime.timedelta(days=60),
        datetime.date(ano, 9, 7),
        datetime.date(ano, 10, 12),
        datetime.date(ano, 11, 2),
        datetime.date(ano, 11, 15),
        datetime.date(ano, 11, 20),
        datetime.date(ano, 12, 25),
    ]
    feriados_mes = [f for f in feriados if f.month == mes and f.year == ano]

    dias_seg_sex_total = 0
    dias_seg_sab_total = 0
    feriados_seg_sex = 0
    feriados_seg_sab = 0
    lista_feriados_detalhes = []

    for semana in cal:
        for i in range(7):
            dia = semana[i]
            if dia != 0:
                data_atual = datetime.date(ano, mes, dia)
                wd = data_atual.weekday()
                if wd < 5:
                    dias_seg_sex_total += 1
                    dias_seg_sab_total += 1
                elif wd == 5:
                    dias_seg_sab_total += 1

                if data_atual in feriados_mes:
                    if wd < 5:
                        feriados_seg_sex += 1
                        feriados_seg_sab += 1
                        lista_feriados_detalhes.append((data_atual, "Seg a Sex"))
                    elif wd == 5:
                        feriados_seg_sab += 1
                        lista_feriados_detalhes.append((data_atual, "Sábado"))

    return {
        "seg_sex_total": dias_seg_sex_total,
        "seg_sex_feriados": feriados_seg_sex,
        "seg_sex_uteis": dias_seg_sex_total - feriados_seg_sex,
        "seg_sab_total": dias_seg_sab_total,
        "seg_sab_feriados": feriados_seg_sab,
        "seg_sab_uteis": dias_seg_sab_total - feriados_seg_sab,
        "feriados_detalhes": lista_feriados_detalhes
    }

def gerar_arquivo_atualizado_bytes(source_input, header, fila, df_original, sheet_name=None):
    wb = load_workbook(io.BytesIO(source_input) if isinstance(source_input, bytes) else source_input)
    ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb[wb.sheetnames[0]]
    for mod in fila:
        col_target = mod['coluna']
        valor_convertido = converter_valor_inteligente(mod['novo_valor'], df_original[col_target].dtype)
        for idx in mod['indices']:
            excel_row = idx + header + 1
            ws.cell(row=excel_row, column=df_original.columns.get_loc(col_target) + 1, value=valor_convertido)

            if col_target.strip().upper() in ["SAIDA", "SAÍDA"]:
                for col_idx in range(1, ws.max_column + 1):
                    cell = ws.cell(row=excel_row, column=col_idx)
                    current_font = cell.font
                    if current_font:
                        cell.font = Font(
                            name=current_font.name,
                            size=current_font.size,
                            bold=current_font.bold,
                            italic=current_font.italic,
                            strike=current_font.strike,
                            underline=current_font.underline,
                            color="FF0000"
                        )
                    else:
                        cell.font = Font(color="FF0000")

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()

def titulo_estilizado(subtitulo=""):
    st.markdown(
        f"<div style='text-align: center; padding: 1.5rem; background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%); color: white; border-radius: 12px; margin-bottom: 1.5rem;'><h1>⚡ SINALE WEB</h1><p>{subtitulo}</p></div>",
        unsafe_allow_html=True
    )

def obter_nome_coluna_por_letra(df, colunas_disponiveis, letra):
    mapa_letras = {
        'A': 0, 'B': 1, 'C': 2, 'D': 3, 'E': 4, 'F': 5, 'G': 6, 'H': 7,
        'I': 8, 'J': 9, 'K': 10, 'L': 11, 'M': 12, 'N': 13, 'O': 14,
        'P': 15, 'Q': 16, 'R': 17, 'S': 18, 'T': 19, 'U': 20, 'V': 21,
        'W': 22, 'X': 23, 'Y': 24, 'Z': 25
    }
    idx = mapa_letras.get(letra.upper())
    if idx is not None and idx < len(colunas_disponiveis):
        return colunas_disponiveis[idx]
    return None

def gerar_config_largura_colunas(df_subset, colunas):
    config = {}
    for col in colunas:
        if col in df_subset.columns:
            nome_coluna_upper = str(col).strip().upper()
            
            if nome_coluna_upper == "NOME":
                tamanho_conteudo = df_subset[col].astype(str).str.len().max() if not df_subset[col].empty else 10
                if pd.isna(tamanho_conteudo):
                    tamanho_conteudo = 10
                
                largura_pixels = int(tamanho_conteudo * 8) + 20
                largura_pixels = max(150, min(largura_pixels, 450))
            else:
                tamanho_titulo = len(str(col))
                largura_pixels = int(tamanho_titulo * 9) + 20
                largura_pixels = max(50, largura_pixels)
            
            config[col] = st.column_config.Column(width=largura_pixels)
            
    return config

def limpar_resultados_downstream():
    """Limpa a visualização e caches para manter a tela limpa ao alterar seleções."""
    st.session_state["executar_config"] = False
    st.session_state["pesquisa_df"] = None
    st.session_state["arquivos_drive_alvo"] = []
    st.session_state["bytes_cache"] = {}

def alternar_marcar_desmarcar_pasta(folder_id, folder_name="Pasta"):
    """Alterna a seleção de todos os itens da pasta."""
    try:
        raw_itens = listar_arquivos_da_pasta(folder_id)
        itens = [f for f in raw_itens if eh_arquivo_valido(f)]
    except Exception:
        itens = []

    itens_ids = [item["id"] for item in itens] + [folder_id]
    todos_selecionados = all(i_id in st.session_state.get("itens_selecionados_map", {}) for i_id in itens_ids)

    novo_estado = not todos_selecionados

    if not novo_estado:
        st.session_state["itens_selecionados_map"].pop(folder_id, None)
        st.session_state[f"chk_folder_{folder_id}__in__{folder_id}"] = False
        for item in itens:
            i_id = item["id"]
            st.session_state["itens_selecionados_map"].pop(i_id, None)
            if item.get('mimeType') == 'application/vnd.google-apps.folder':
                st.session_state[f"chk_folder_{i_id}__in__{folder_id}"] = False
            else:
                st.session_state[f"chk_file_{i_id}__in__{folder_id}"] = False
    else:
        if "pastas_abertas" not in st.session_state: st.session_state["pastas_abertas"] = set()
        st.session_state["pastas_abertas"].add(folder_id)
        st.session_state["itens_selecionados_map"][folder_id] = {"name": folder_name, "type": "folder", "id": folder_id}
        
        for item in itens:
            i_id = item["id"]
            if item.get('mimeType') == 'application/vnd.google-apps.folder':
                st.session_state["pastas_abertas"].add(i_id)
                st.session_state["itens_selecionados_map"][i_id] = {"name": item["name"], "type": "folder", "id": i_id}
                st.session_state[f"chk_folder_{i_id}__in__{folder_id}"] = True
            else:
                st.session_state["itens_selecionados_map"][i_id] = {"name": item["name"], "type": "file", "id": i_id}
                st.session_state[f"chk_file_{i_id}__in__{folder_id}"] = True

    limpar_resultados_downstream()

# =============================================================================
# --- CONFIGURAÇÃO DA PÁGINA E ESTADOS ---
# =============================================================================

st.set_page_config(page_title="SINALE WEB", layout="wide")

if "source_df" not in st.session_state:
    st.session_state["source_df"] = None
if "wb_data" not in st.session_state:
    st.session_state["wb_data"] = None
if "last_dest_name" not in st.session_state:
    st.session_state["last_dest_name"] = None
if "fila_modificacoes" not in st.session_state:
    st.session_state["fila_modificacoes"] = []
if "select_all" not in st.session_state:
    st.session_state["select_all"] = False
if "file_settings" not in st.session_state:
    st.session_state["file_settings"] = {}
if "pesquisa_df" not in st.session_state:
    st.session_state["pesquisa_df"] = None
if "executar_config" not in st.session_state:
    st.session_state["executar_config"] = False


# =============================================================================
# --- MENU PRINCIPAL ---
# =============================================================================

menu_opcao = st.sidebar.radio(
    "Selecione a rotina:",
    [
        "INCLUSÃO DE TRABALHO",
        "ATUALIZAÇÕES GERAIS",
        "PESQUISA PARA REMIÇÃO",
        "LIMPAR ARQUIVO",
        "SOMENTE TRABALHADORES ATIVOS",
        "SAIR DO SISTEMA"
    ]
)


# =============================================================================
# --- OPÇÃO 1: INCLUSÃO DE TRABALHO ---
# =============================================================================
if menu_opcao == "INCLUSÃO DE TRABALHO":
    titulo_estilizado("INTEGRADOR ==> DADOS GERAIS DO INTERNO >>> SINALE")
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("1. Arquivo de ORIGEM")
        source_file = st.file_uploader("Selecione o arquivo de ORIGEM", type=["xlsx", "xls", "csv", "txt"], key="src_upload")
        origem_tem_cabecalho = st.checkbox("Arquivo de Origem tem cabeçalho?", value=True)
    with col2:
        st.subheader("2. Arquivo de DESTINO")
        dest_file = st.file_uploader("Selecione o arquivo de DESTINO (.xlsx)", type=["xlsx"], key="dest_upload")
        header_dest = st.number_input("Linha do cabeçalho no Arquivo de Destino:", value=11, min_value=1)

    if source_file:
        cache_key_src = f"{source_file.name}_{origem_tem_cabecalho}"
        if "source_df" not in st.session_state or st.session_state.get("last_cache_key_src") != cache_key_src:
            hdr = 0 if origem_tem_cabecalho else None
            try:
                source_file.seek(0)
                ext = source_file.name.split('.')[-1].lower()
                engine_util = 'xlrd' if ext == 'xls' else ('openpyxl' if ext == 'xlsx' else None)
                raw = pd.read_excel(source_file, header=hdr, engine=engine_util)
                raw.columns = deduplicar_colunas(raw.columns) if origem_tem_cabecalho else [f"Col {i+1}" for i in range(len(raw.columns))]
                st.session_state["source_df"] = raw
                st.session_state["last_cache_key_src"] = cache_key_src
            except Exception as e:
                st.error(f"Erro ao ler arquivo: {e}")

    if dest_file:
        if "wb_data" not in st.session_state or st.session_state.get("last_dest_name") != dest_file.name:
            dest_file.seek(0)
            st.session_state["wb_data"] = dest_file.getvalue()
            st.session_state["last_dest_name"] = dest_file.name

    df_origem = st.session_state.get("source_df")
    wb_data = st.session_state.get("wb_data")

    if df_origem is not None and wb_data is not None:
        wb = load_workbook(io.BytesIO(wb_data))
        target_sheet = st.selectbox("Escolha a ABA na Planilha de Destino a ser Atualizada:", wb.sheetnames)
        ws = wb[target_sheet]

        st.subheader("3. Seleção de Registros")
        col_busca = st.selectbox("Coluna identificadora (para seleção):", df_origem.columns)
        opcoes_selecao = [f"{val} (Linha {idx})" for idx, val in df_origem[col_busca].items()]
        selected_options = st.multiselect("🔍 Escolha os registros:", opcoes_selecao)
        selected_indices = [int(item.split("(Linha ")[1].replace(")", "")) for item in selected_options]

        if selected_indices:
            st.info(f"📊 **{len(selected_indices)}** registro(s) selecionado(s) para atualização.")

        st.write("---")
        st.subheader("4. Correlação dos dados dos Arquivos ORIGEM X DESTINO")
        mapping = {}
        cols_ui = st.columns(4)
        opcoes_mapeamento = ["--- Não mapear ---", "⚠️ Auto-incrementar (Seq)"] + list(df_origem.columns)
        for i in range(1, ws.max_column + 1):
            header_val = ws.cell(row=header_dest, column=i).value
            with cols_ui[(i - 1) % 4]:
                map_val = st.selectbox(f"Col {i} ({header_val or 'S/ Título'})", opcoes_mapeamento, key=f"map_{i}")
                if map_val != "--- Não mapear ---":
                    mapping[i] = map_val

        st.write("---")
        st.subheader("5. Local da Atualização")
        modo_insercao = st.radio("Local de inserção:", ["Final da planilha", "A partir de uma linha específica"])
        target_row = st.number_input("Linha:", min_value=header_dest + 1, value=header_dest + 1) if modo_insercao == "A partir de uma linha específica" else ws.max_row + 1

        st.write("---")
        if st.button("🚀 Processar e Atualizar"):
            if not selected_indices:
                st.error("Selecione itens!")
                st.stop()
            ref_row_idx = (target_row - 1) if modo_insercao == "A partir de uma linha específica" else ws.max_row
            base_seq = 0
            if ref_row_idx >= header_dest:
                val_acima = ws.cell(row=ref_row_idx, column=1).value
                try:
                    base_seq = int(val_acima)
                except:
                    base_seq = 0
            if modo_insercao == "A partir de uma linha específica":
                ws.insert_rows(target_row, amount=len(selected_indices))
            current_row = target_row
            seq_val = base_seq
            for idx in selected_indices:
                seq_val += 1
                ref_row_idx = current_row - 1
                for col_idx in range(1, ws.max_column + 1):
                    target_cell = ws.cell(row=current_row, column=col_idx)
                    ref_cell = ws.cell(row=ref_row_idx, column=col_idx)
                    copiar_estilo_completo(ref_cell, target_cell)
                    if col_idx == 1 or mapping.get(col_idx) == "⚠️ Auto-incrementar (Seq)":
                        target_cell.value = seq_val
                    elif col_idx in mapping:
                        target_cell.value = extrair_valor_limpo(df_origem, idx, mapping[col_idx])
                    else:
                        target_cell.value = ref_cell.value
                current_row += 1
            buffer = io.BytesIO()
            wb.save(buffer)
            st.session_state["wb_data"] = buffer.getvalue()
            st.success("✅ Processamento concluído com sucesso!")
            st.download_button(
                "📥 Baixar Versão Atualizada",
                st.session_state["wb_data"],
                "sinale_atualizado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# =============================================================================
# --- OPÇÃO 2: ATUALIZAÇÕES GERAIS ---
# =============================================================================
elif menu_opcao == "ATUALIZAÇÕES GERAIS":
    titulo_estilizado("Atualizações Gerais")

    if st.session_state.get("wb_data") is not None:
        st.info("📁 Arquivo carregado automaticamente da memória.")
        if st.checkbox("🗑️ Descartar dados da memória e carregar novo arquivo", value=False, key="desc_op2"):
            st.session_state["wb_data"] = None
            st.session_state["fila_modificacoes"] = []
            st.success("Memória limpa com sucesso!")
            st.rerun()
    else:
        st.warning("⚠️ Nenhum arquivo de destino encontrado na memória. Faça o upload abaixo.")
        sinale_file = st.file_uploader("Selecione o arquivo do SINALE (.xlsx)", type=["xlsx"], key="upload_op2")
        if sinale_file:
            st.session_state["wb_data"] = sinale_file.getvalue()
            st.session_state["last_sinale_name"] = sinale_file.name
            st.rerun()

    if st.session_state.get("wb_data") is not None:
        wb_temp = load_workbook(io.BytesIO(st.session_state["wb_data"]), data_only=True)
        target_sheet = st.selectbox("Escolha a ABA do arquivo para trabalhar:", wb_temp.sheetnames, key="aba_op2")
        header = st.number_input("Linha do cabeçalho:", value=11, min_value=1, key="header_op2")
        df = pd.read_excel(io.BytesIO(st.session_state["wb_data"]), sheet_name=target_sheet, header=header - 1)

        st.subheader("🔍 Filtros de Visualização")
        cols_para_ver = st.multiselect("Quais campos deseja visualizar?", df.columns.tolist(), default=df.columns.tolist())
        col_filtro, val_filtro = st.columns(2)
        with col_filtro:
            filtro_col = st.selectbox("Coluna para buscar:", df.columns, key="filtro_col_op2")
        valores_existentes = sorted([str(v) for v in df[filtro_col].dropna().unique()])
        with val_filtro:
            filtro_vals = st.multiselect("Selecione o(s) valor(es) para filtrar:", valores_existentes, key="filtro_vals_op2")

        df_view = df.copy()
        if filtro_vals:
            df_view = df_view[df_view[filtro_col].astype(str).isin(filtro_vals)]
        st.metric("Total de Registros Encontrados", len(df_view))

        df_view_fmt = formatar_datas_dataframe(df_view[cols_para_ver])
        st.dataframe(df_view_fmt, use_container_width=True, hide_index=True)

        st.subheader("✏️ Seleção para Atualizar")
        if "select_all" not in st.session_state:
            st.session_state["select_all"] = False
        cols_btns = st.columns([1, 1, 4])
        with cols_btns[0]:
            if st.button("✅ Marcar Todos", key="btn_marcar_t"):
                st.session_state["select_all"] = True
                st.rerun()
        with cols_btns[1]:
            if st.button("❌ Desmarcar Todos", key="btn_desmarcar_t"):
                st.session_state["select_all"] = False
                st.rerun()

        df_for_edit = df_view.copy()
        df_for_edit.insert(0, "Atualizar?", st.session_state["select_all"])
        df_editado = st.data_editor(
            df_for_edit,
            column_config={"Atualizar?": st.column_config.CheckboxColumn()},
            use_container_width=True,
            key="editor_op2"
        )

        selecionados = df_editado[df_editado["Atualizar?"] == True]
        st.metric("Total de Registros Marcados", len(selecionados))

        if not selecionados.empty:
            col_target = st.selectbox("Selecione a coluna que deseja alterar:", df.columns, key="col_target_op2")
            if col_target.strip().upper() == "DIAS":
                st.markdown("---")
                st.subheader("📅 Cálculo Automático de Dias Úteis (Seg a Sáb / Seg a Sex)")
                c_mes, c_ano = st.columns(2)
                meses_dict = {
                    "Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4, "Maio": 5, "Junho": 6,
                    "Julho": 7, "Agosto": 8, "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12
                }
                with c_mes:
                    mes_escolhido_nome = st.selectbox("Selecione o Mês:", list(meses_dict.keys()), key="sel_mes_dias")
                    mes_num = meses_dict[mes_escolhido_nome]
                with c_ano:
                    ano_escolhido = st.number_input("Digite o Ano:", min_value=2020, max_value=2035, value=datetime.date.today().year, key="sel_ano_dias")
                stats = obter_estatisticas_mes(ano_escolhido, mes_num)
                st.info(f"**Resumo para {mes_escolhido_nome}/{ano_escolhido}:**\n* **Segunda a Sábado:** {stats['seg_sab_total']} brutos | **Úteis:** **{stats['seg_sab_uteis']}**\n* **Segunda a Sexta:** {stats['seg_sex_total']} brutos | **Úteis:** **{stats['seg_sex_uteis']}**")

            valores_antigos_str = ", ".join([str(v) for v in selecionados[col_target].dropna().unique()])
            st.info(f"📌 **Valor(es) atual(is) / antigo(s)** no campo **'{col_target}'**: **{valores_antigos_str if valores_antigos_str else 'Vazio'}**")
            novo_val = st.text_input("Digite o novo valor:", key="novo_val_op2")

            if st.button("➕ Adicionar à Fila de Modificações", key="btn_add_fila"):
                st.session_state["fila_modificacoes"].append({
                    "indices": selecionados.index.tolist(),
                    "coluna": col_target,
                    "novo_valor": novo_val,
                    "valor_antigo": valores_antigos_str,
                    "vl_busca": ", ".join(filtro_vals) if filtro_vals else "Todos",
                    "aba": target_sheet
                })
                st.success("Modificação adicionada à fila!")
                st.rerun()

        if st.session_state["fila_modificacoes"]:
            st.markdown("---")
            st.subheader("📋 Fila de Modificações Pendentes")
            df_fila_resumo = pd.DataFrame([
                {
                    "Remover?": False,
                    "ID_ITEM": i,
                    "ABA": item.get("aba", "Geral"),
                    "CAMPO": item.get("coluna", ""),
                    "NOVO VALOR": item.get("novo_valor", "")
                }
                for i, item in enumerate(st.session_state["fila_modificacoes"])
            ])
            df_fila_editado = st.data_editor(
                df_fila_resumo,
                column_config={"Remover?": st.column_config.CheckboxColumn("Remover?"), "ID_ITEM": None},
                disabled=["ABA", "CAMPO", "NOVO VALOR"],
                use_container_width=True,
                key="editor_fila"
            )
            col_f1, col_f2, col_f3 = st.columns(3)
            with col_f1:
                if st.button("🗑️ Remover Selecionados"):
                    indices = df_fila_editado[df_fila_editado["Remover?"] == True]["ID_ITEM"].tolist()
                    st.session_state["fila_modificacoes"] = [
                        item for i, item in enumerate(st.session_state["fila_modificacoes"]) if i not in indices
                    ]
                    st.rerun()
            with col_f3:
                file_bytes = gerar_arquivo_atualizado_bytes(
                    io.BytesIO(st.session_state["wb_data"]),
                    header,
                    st.session_state["fila_modificacoes"],
                    df,
                    sheet_name=target_sheet
                )
                st.download_button(
                    "📥 Baixar Arquivo Atualizado",
                    file_bytes,
                    "sinale_atualizado_final.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

# =============================================================================
# --- OPÇÃO 3: PESQUISA PARA REMIÇÃO (VIA GOOGLE DRIVE) ---
# =============================================================================
elif menu_opcao == "PESQUISA PARA REMIÇÃO":
    titulo_estilizado("Pesquisa para Remição")

    if "uploader_key" not in st.session_state:
        st.session_state["uploader_key"] = 0

    # Defina o ID da pasta raiz do seu Drive aqui
    ROOT_FOLDER_ID = "1ZeCu40Bzt1hb1BsgNArG_zKR54GPcuOY"

    if "itens_selecionados_map" not in st.session_state:
        st.session_state["itens_selecionados_map"] = {}

    if "pastas_abertas" not in st.session_state:
        st.session_state["pastas_abertas"] = set()

    st.subheader("1. Seleção de Pastas e Arquivos (Google Drive)")

    if not ROOT_FOLDER_ID.strip():
        st.error("ID da pasta do Google Drive não configurado.")
        st.stop()

    ITEMS_PER_ROW = 3
    pastas_para_processar = [{"id": ROOT_FOLDER_ID, "name": "Pasta Raiz"}]
    pastas_processadas = set()

    while pastas_para_processar:
        atual = pastas_para_processar.pop(0)
        p_id = atual["id"]
        p_name = atual["name"]

        if p_id in pastas_processadas:
            continue
        pastas_processadas.add(p_id)

        with st.container():
            st.markdown(f"📂 **Pasta Atual: `{p_name}`**")

            try:
                raw_itens = listar_arquivos_da_pasta(p_id)
                itens = [f for f in raw_itens if eh_arquivo_valido(f)]
            except Exception as e:
                st.error(f"Erro ao listar arquivos de '{p_name}': {e}")
                itens = []

            if not itens:
                st.caption("*(Pasta vazia ou sem arquivos válidos)*")
                st.markdown("---")
                continue

            pastas = sorted([f for f in itens if f.get('mimeType') == 'application/vnd.google-apps.folder'], key=lambda x: x['name'].lower())
            arquivos = sorted([f for f in itens if f.get('mimeType') != 'application/vnd.google-apps.folder'], key=lambda x: x['name'].lower())

            col_b1, _ = st.columns([2, 4])
            with col_b1:
                if st.button("Marcar / Desmarcar Tudo", key=f"btn_toggle_all_{p_id}"):
                    alternar_marcar_desmarcar_pasta(p_id, p_name)
                    st.rerun()

            if pastas:
                st.markdown("**📁 Subpastas:**")
                for i in range(0, len(pastas), ITEMS_PER_ROW):
                    cols = st.columns(ITEMS_PER_ROW)
                    for idx, p in enumerate(pastas[i:i + ITEMS_PER_ROW]):
                        with cols[idx]:
                            sub_id = p['id']
                            sub_name = p['name']
                            chk_key = f"chk_folder_{sub_id}__in__{p_id}"
                            
                            if chk_key not in st.session_state:
                                st.session_state[chk_key] = (sub_id in st.session_state["itens_selecionados_map"])

                            checked = st.checkbox(f"📁 {sub_name}", key=chk_key)

                            is_selected = (sub_id in st.session_state["itens_selecionados_map"])
                            if checked != is_selected:
                                if checked:
                                    st.session_state["itens_selecionados_map"][sub_id] = {"name": sub_name, "type": "folder", "id": sub_id}
                                    st.session_state["pastas_abertas"].add(sub_id)
                                else:
                                    st.session_state["itens_selecionados_map"].pop(sub_id, None)
                                limpar_resultados_downstream()
                                st.rerun()

                            if checked or sub_id in st.session_state["pastas_abertas"]:
                                pastas_para_processar.append({"id": sub_id, "name": sub_name})

            if arquivos:
                st.markdown("**📄 Arquivos:**")
                for i in range(0, len(arquivos), ITEMS_PER_ROW):
                    cols = st.columns(ITEMS_PER_ROW)
                    for idx, f in enumerate(arquivos[i:i + ITEMS_PER_ROW]):
                        with cols[idx]:
                            f_id = f['id']
                            f_name = f['name']
                            chk_key = f"chk_file_{f_id}__in__{p_id}"
                            
                            if chk_key not in st.session_state:
                                st.session_state[chk_key] = (f_id in st.session_state["itens_selecionados_map"])

                            checked = st.checkbox(f"📄 {f_name}", key=chk_key)

                            is_selected = (f_id in st.session_state["itens_selecionados_map"])
                            if checked != is_selected:
                                if checked:
                                    st.session_state["itens_selecionados_map"][f_id] = {"name": f_name, "type": "file", "id": f_id}
                                else:
                                    st.session_state["itens_selecionados_map"].pop(f_id, None)
                                limpar_resultados_downstream()
                                st.rerun()
            st.markdown("---")

    st.markdown("### 📋 Área de Itens Selecionados (Unificação Geral)")

    arquivos_finais_map = {k: v["name"] for k, v in st.session_state["itens_selecionados_map"].items() if v["type"] == "file"}
    arquivos_selecionados_lista = sorted([{"id": fid, "name": fname} for fid, fname in arquivos_finais_map.items()], key=lambda x: x["name"].lower())

    with st.expander(f"📦 Resumo Consolidado de Arquivos Selecionados ({len(arquivos_selecionados_lista)} arquivo(s) mapeado(s))", expanded=True):
        if arquivos_selecionados_lista:
            st.dataframe(pd.DataFrame(arquivos_selecionados_lista)[["name"]].rename(columns={"name": "Nome do Arquivo"}), use_container_width=True, hide_index=True)
            if st.button("❌ Limpar Toda a Seleção"):
                st.session_state["itens_selecionados_map"] = {}
                st.session_state["pastas_abertas"] = set()
                for k in list(st.session_state.keys()):
                    if k.startswith("chk_folder_") or k.startswith("chk_file_"):
                        st.session_state[k] = False
                limpar_resultados_downstream()
                st.rerun()
        else:
            st.info("Nenhum arquivo selecionado até o momento.")

    fazer_upload_btn = st.button("2. Carregar do Drive e Configurar Abas", key="btn_fazer_upload_op3", type="primary")

    if fazer_upload_btn:
        if arquivos_selecionados_lista:
            st.session_state["executar_config"] = True
            st.session_state["rolar_apos_upload"] = True
            st.session_state["arquivos_drive_alvo"] = arquivos_selecionados_lista
            st.success("Arquivos prontos para processamento!")
        else:
            st.error("Selecione pelo menos um arquivo para continuar.")
            limpar_resultados_downstream()

    arquivos_alvo = st.session_state.get("arquivos_drive_alvo", [])
    if arquivos_alvo and st.session_state.get("executar_config"):
        settings = {}
        if "bytes_cache" not in st.session_state:
            st.session_state["bytes_cache"] = {}

        for f_idx, f_info in enumerate(arquivos_alvo):
            f_name = f_info["name"]
            f_id = f_info["id"]
            file_key = f"{f_idx}_{f_name}"

            if f_id not in st.session_state["bytes_cache"]:
                with st.spinner(f"Baixando {f_name} do Drive..."):
                    st.session_state["bytes_cache"][f_id] = baixar_arquivo_bytes(f_id)

            f_bytes = st.session_state["bytes_cache"][f_id]
            file_ext = f_name.split('.')[-1].lower()

            try:
                engine_val = 'odf' if file_ext == 'ods' else None
                xl = pd.ExcelFile(io.BytesIO(f_bytes), engine=engine_val)
                sheets_available = xl.sheet_names
            except Exception as e:
                st.error(f"Erro ao ler o arquivo {f_name}: {e}.")
                continue

            pref_sheets = [s for s in sheets_available if any(p in s.strip().upper() for p in ["COM REMUNER", "SEM REMUNER", "DEM_COM", "DEM_SEM"])]
            default_sheets = pref_sheets if pref_sheets else ([sheets_available[0]] if sheets_available else [])
            is_fallback = not bool(pref_sheets)

            with st.expander(f"📁 Configurações para: Arquivo {f_idx+1} - **{f_name}**", expanded=True):
                selected_sheets = st.multiselect(f"Selecione aba(s) para {f_name}", sheets_available, default=default_sheets, key=f"sheets_{file_key}_{st.session_state['uploader_key']}")

                sheet_config = {}
                for i, sheet in enumerate(selected_sheets):
                    st.markdown(f"**Aba: `{sheet}`**")
                    sheet_upper = sheet.strip().upper()
                    if "DEM_COM" in sheet_upper: default_header = 17
                    elif "DEM_SEM" in sheet_upper: default_header = 19
                    elif any(p in sheet_upper for p in ["COM REMUNER", "SEM REMUNER"]): default_header = 11
                    else: default_header = 10 if is_fallback else 11

                    header_row = st.number_input(f"Linha do cabeçalho para aba '{sheet}'", value=default_header, min_value=1, key=f"head_{file_key}_{sheet}_{st.session_state['uploader_key']}")

                    try:
                        df_preview = pd.read_excel(io.BytesIO(f_bytes), sheet_name=sheet, header=header_row - 1, nrows=0, engine=engine_val)
                        cols_aba = [str(c).strip() for c in df_preview.columns]
                    except:
                        cols_aba = []

                    default_col = None
                    for c in cols_aba:
                        if str(c).strip().upper() in ["NOME DO INTERNO", "NOME DO INTERNO "]: default_col = c; break
                    if not default_col:
                        for c in cols_aba:
                            if str(c).strip().upper() == "NOME": default_col = c; break
                    if not default_col:
                        for c in cols_aba:
                            if str(c).strip().upper().startswith("NOME"): default_col = c; break
                    if not default_col:
                        for c in cols_aba:
                            if "NOME" in str(c).strip().upper(): default_col = c; break
                    if not default_col and len(cols_aba) > 8: default_col = cols_aba[8]
                    elif not default_col and cols_aba: default_col = cols_aba[0]

                    opcoes_colunas = ["--- Não pesquisar nesta aba ---"] + cols_aba
                    default_idx = opcoes_colunas.index(default_col) if default_col in opcoes_colunas else 0

                    col_escolhida = st.selectbox(f"Selecione o campo para pesquisa '{sheet}':", opcoes_colunas, index=default_idx, key=f"col_search_{file_key}_{sheet}_{st.session_state['uploader_key']}")

                    sheet_config[sheet] = {"header_idx": header_row - 1, "col_busca": col_escolhida if col_escolhida != "--- Não pesquisar nesta aba ---" else None}
                    st.markdown("---")
                settings[file_key] = sheet_config

        btn_consolidar = st.button("🔍 Carregar e Consolidar Dados", key="btn_consolidar_op3", type="primary")

        if st.session_state.get("rolar_apos_upload"):
            components.html("<script>function rolar() { window.scrollTo({ top: document.body.scrollHeight, behavior: 'smooth' }); } setTimeout(rolar, 400);</script>", height=0)
            st.session_state["rolar_apos_upload"] = False

        if btn_consolidar:
            all_results = []
            meses_sigla_map = {"1": "JAN", "01": "JAN", "2": "FEV", "02": "FEV", "3": "MAR", "03": "MAR", "4": "ABR", "04": "ABR", "5": "MAI", "05": "MAI", "6": "JUN", "06": "JUN", "7": "JUL", "07": "JUL", "8": "AGO", "08": "AGO", "9": "SET", "09": "SET", "10": "OUT", "11": "NOV", "12": "DEZ"}

            for f_idx, f_info in enumerate(arquivos_alvo):
                f_name = f_info["name"]
                f_id = f_info["id"]
                file_key = f"{f_idx}_{f_name}"
                f_bytes = st.session_state["bytes_cache"].get(f_id, b"")
                file_ext = f_name.split('.')[-1].lower()
                engine_val = 'odf' if file_ext == 'ods' else None

                try: mes_ano_arquivo = extrair_mes_ano_do_nome(f_name)
                except: mes_ano_arquivo = "SEM MÊS/ANO"

                mes_ano_formatado = mes_ano_arquivo
                if "/" in mes_ano_arquivo and mes_ano_arquivo != "SEM MÊS/ANO":
                    parts = mes_ano_arquivo.split("/")
                    mes_ano_formatado = f"{meses_sigla_map.get(parts[0].strip(), parts[0].strip())}/{parts[1].strip()}"

                for sheet, cfg in settings.get(file_key, {}).items():
                    try:
                        df_tmp = pd.read_excel(io.BytesIO(f_bytes), sheet_name=sheet, header=cfg["header_idx"], engine=engine_val)
                        df_tmp.columns = deduplicar_colunas([str(c).strip() for c in df_tmp.columns])

                        col_pedida = cfg.get("col_busca")
                        target_col = next((c for c in df_tmp.columns if str(c).strip().upper() == str(col_pedida).strip().upper()), None) if col_pedida else None
                        target_col = target_col or next((c for c in df_tmp.columns if "NOME DO INTERNO" in str(c).strip().upper()), None)
                        target_col = target_col or next((c for c in df_tmp.columns if "NOME" in str(c).strip().upper()), None)
                        target_col = target_col or (df_tmp.columns[8] if len(df_tmp.columns) > 8 else (df_tmp.columns[0] if len(df_tmp.columns) > 0 else None))

                        if target_col and target_col in df_tmp.columns:
                            colunas_originais = list(df_tmp.columns)
                            df_tmp["Aba Original"] = sheet
                            df_tmp["Campo Pesquisado"] = target_col
                            df_tmp["Nome (Visualização)"] = df_tmp[target_col].astype(str).str.strip()
                            df_tmp["NOME_LIMPO"] = df_tmp["Nome (Visualização)"].str.upper()
                            df_tmp = df_tmp[~df_tmp["NOME_LIMPO"].isin(['', 'NAN', 'NONE', '0', 'NAT', 'NC', 'N/C'])].copy()

                            aba_upper = sheet.strip().upper()
                            is_dem_com = "DEM_COM" in aba_upper
                            is_dem_sem = "DEM_SEM" in aba_upper
                            is_com_remuner = "COM REMUNER" in aba_upper
                            is_sem_remuner = "SEM REMUNER" in aba_upper
                            col_f = obter_nome_coluna_por_letra(df_tmp, colunas_originais, 'F')

                            usar_padrao_antigo = False; usar_dem_sem_antigo = False
                            is_03_a_05_2023 = False; is_06_a_07_2023 = False; is_08_2023 = False

                            if mes_ano_arquivo != "SEM MÊS/ANO":
                                try:
                                    mes_val, ano_val = int(mes_ano_arquivo.split('/')[0]), int(mes_ano_arquivo.split('/')[1])
                                    if ano_val == 2023 and mes_val in [3, 4, 5]: is_03_a_05_2023 = True
                                    elif ano_val == 2023 and mes_val in [6, 7]: is_06_a_07_2023 = True
                                    elif ano_val == 2023 and mes_val == 8: is_08_2023 = True
                                    if ano_val < 2025 or (ano_val == 2025 and mes_val < 9): usar_padrao_antigo = True
                                    if ano_val < 2019 or (ano_val == 2019 and mes_val < 11): usar_dem_sem_antigo = True
                                except Exception: pass

                            def extrair_dados_e_categoria(row):
                                if is_03_a_05_2023:
                                    if is_dem_com or is_com_remuner: cat = "COM REMUNERAÇÃO"; letras = ["I", "B", "T", "V", "W", "X", "Y"]
                                    elif is_dem_sem or is_sem_remuner: cat = "SEM REMUNERAÇÃO"; letras = ["J", "B", "S", "U", "V", "W", "X"]
                                    else:
                                        is_sim = str(row.get(col_f, "")).strip().upper() == "SIM"
                                        cat = "COM REMUNERAÇÃO" if is_sim else "SEM REMUNERAÇÃO"
                                        letras = ["I", "B", "T", "V", "W", "X", "Y"] if is_sim else ["J", "B", "S", "U", "V", "W", "X"]
                                elif is_06_a_07_2023:
                                    if is_dem_com or is_com_remuner: cat = "COM REMUNERAÇÃO"; letras = ["I", "B", "U", "W", "X", "Y", "Z"]
                                    elif is_dem_sem or is_sem_remuner: cat = "SEM REMUNERAÇÃO"; letras = ["J", "B", "S", "V", "W", "X", "Y"]
                                    else:
                                        is_sim = str(row.get(col_f, "")).strip().upper() == "SIM"
                                        cat = "COM REMUNERAÇÃO" if is_sim else "SEM REMUNERAÇÃO"
                                        letras = ["I", "B", "U", "W", "X", "Y", "Z"] if is_sim else ["J", "B", "S", "V", "W", "X", "Y"]
                                elif is_08_2023:
                                    if is_dem_com or is_com_remuner: cat = "COM REMUNERAÇÃO"; letras = ["I", "B", "R", "T", "U", "V", "W"]
                                    elif is_dem_sem or is_sem_remuner: cat = "SEM REMUNERAÇÃO"; letras = ["I", "B", "Q", "S", "T", "U", "V"]
                                    else:
                                        is_sim = str(row.get(col_f, "")).strip().upper() == "SIM"
                                        cat = "COM REMUNERAÇÃO" if is_sim else "SEM REMUNERAÇÃO"
                                        letras = ["I", "B", "R", "T", "U", "V", "W"] if is_sim else ["I", "B", "Q", "S", "T", "U", "V"]
                                else:
                                    if is_dem_com: cat = "COM REMUNERAÇÃO"; letras = ["I", "B", None, "S", "T", "U", "V"]
                                    elif is_dem_sem: cat = "SEM REMUNERAÇÃO"; letras = ["I", "B", "Y", "R", "S", "T", "U"] if usar_dem_sem_antigo else ["I", "B", "Y", "S", "T", "U", "V"]
                                    elif is_com_remuner: cat = "COM REMUNERAÇÃO"; letras = ["I", "B", "Q", "S", "T", "U", "V"] if usar_padrao_antigo else ["B", "I", "J", "T", "U", "V", "W"]
                                    elif is_sem_remuner: cat = "SEM REMUNERAÇÃO"; letras = ["I", "B", "W", "R", "S", "T", "U"]
                                    else:
                                        is_sim = str(row.get(col_f, "")).strip().upper() == "SIM"
                                        cat = "COM REMUNERAÇÃO" if is_sim else "SEM REMUNERAÇÃO"
                                        letras = ["J", "C", "X", "S", "T", "U", "V"]

                                row_vals = {"Categoria_Aba": cat, "LABEL_EXIBICAO": f"{mes_ano_formatado} - {cat}"}
                                for idx_p, let in enumerate(letras):
                                    if let is None: row_vals[f"POS_{idx_p}"] = ""; row_vals[f"HEADER_{idx_p}"] = ""
                                    else:
                                        col_n = obter_nome_coluna_por_letra(df_tmp, colunas_originais, let)
                                        row_vals[f"POS_{idx_p}"] = row.get(col_n, None)
                                        row_vals[f"HEADER_{idx_p}"] = str(col_n) if col_n else f"Campo {idx_p+1}"
                                return pd.Series(row_vals)

                            res_df = df_tmp.apply(extrair_dados_e_categoria, axis=1)
                            df_tmp["MÊS/ANO - ABA"] = res_df["LABEL_EXIBICAO"]
                            all_results.append(pd.concat([df_tmp[["MÊS/ANO - ABA", "Aba Original", "Campo Pesquisado", "Nome (Visualização)", "NOME_LIMPO"]], res_df], axis=1))
                    except Exception as e: st.error(f"Erro ao ler {f_name} - Aba {sheet}: {e}")

            if all_results:
                st.session_state["pesquisa_df"] = pd.concat(all_results, ignore_index=True)
                st.success(f"Dados consolidados! **{len(st.session_state['pesquisa_df'])}** registros carregados.")
            else:
                st.warning("Nenhum dado encontrado.")
                st.session_state["pesquisa_df"] = None

    if st.session_state.get("pesquisa_df") is not None:
        df_pesq = st.session_state["pesquisa_df"]
        st.markdown("---")
        st.subheader("🔍 Filtros de Visualização e Busca")

        col_ord1, _ = st.columns([2, 2])
        with col_ord1:
            is_ascending = "Crescente" in st.radio("📅 Ordenação:", ["Crescente (Antigo ➔ Recente)", "Decrescente (Recente ➔ Antigo)"], horizontal=True, key=f"ord_{st.session_state['uploader_key']}")

        nomes_selecionados = st.multiselect("🔍 Pesquisar nome(s):", sorted(df_pesq["Nome (Visualização)"].dropna().unique()), key=f"busca_{st.session_state['uploader_key']}")
        df_view = df_pesq[df_pesq["Nome (Visualização)"].isin(nomes_selecionados)] if nomes_selecionados else df_pesq.copy()
        st.metric("Total de Registros Encontrados", len(df_view))

        if not df_view.empty:
            def extrair_chave_data(val):
                try:
                    ds = str(val).split(' - ')[0].strip()
                    if ds == "SEM MÊS/ANO": return 999999 if is_ascending else -1
                    m_inv = {"JAN": "01", "FEV": "02", "MAR": "03", "ABR": "04", "MAI": "05", "JUN": "06", "JUL": "07", "AGO": "08", "SET": "09", "OUT": "10", "NOV": "11", "DEZ": "12"}
                    m, y = ds.split('/')
                    return int(y) * 100 + int(m_inv.get(m.upper(), m))
                except: return 999999 if is_ascending else -1

            df_view['chave_ordenacao'] = df_view['MÊS/ANO - ABA'].apply(extrair_chave_data)
            df_display_all = formatar_datas_dataframe(df_view.sort_values(by=['chave_ordenacao'], ascending=is_ascending).drop(columns=['chave_ordenacao']))

            def fmt_num(val): return "" if pd.isna(val) else str(int(round(float(val)))) if isinstance(val, (int, float)) or (isinstance(val, str) and val.replace('.','',1).isdigit()) else str(val).strip()
            def conv_num(val):
                try: return float(str(val).replace(',', '.').strip())
                except: return 0.0

            mapa_meses = {"01":"JAN", "1":"JAN", "02":"FEV", "2":"FEV", "03":"MAR", "3":"MAR", "04":"ABR", "4":"ABR", "05":"MAI", "5":"MAI", "06":"JUN", "6":"JUN", "07":"JUL", "7":"JUL", "08":"AGO", "8":"AGO", "09":"SET", "9":"SET", "10":"OUT", "11":"NOV", "12":"DEZ"}
            
            todos_dados_exportacao = []

            for titulo, cat_key, prefixo in [("🟢 COM REMUNERAÇÃO", "COM REMUNERAÇÃO", "com"), ("🟡 SEM REMUNERAÇÃO", "SEM REMUNERAÇÃO", "sem")]:
                df_grupo = df_display_all[df_display_all["Categoria_Aba"] == cat_key]
                if not df_grupo.empty:
                    pos_cols = sorted([c for c in df_grupo.columns if str(c).startswith("POS_")], key=lambda x: int(x.split("_")[1]))
                    cabs = ["NOME", "ORGANIZ", "FUNÇÃO", "ENTRADA", "SAIDA", "PREV", "REAL"]
                    rename_map = {pc: cabs[i] if i < len(cabs) else f"Campo {i+1}" for i, pc in enumerate(pos_cols)}

                    df_render = df_grupo[["MÊS/ANO - ABA"] + pos_cols].rename(columns=rename_map).rename(columns={"MÊS/ANO - ABA": "MES/ANO - ABA"})
                    if "REAL" in df_render.columns: df_render["REAL"] = df_render["REAL"].apply(fmt_num)

                    st.markdown(f"### {titulo} ({len(df_render)} registro(s))")
                    k_sel = f"sel_all_{prefixo}"
                    if k_sel not in st.session_state: st.session_state[k_sel] = False
                    if st.columns([2, 4])[0].button("Marcar / Desmarcar Todos", key=f"btn_tgl_{prefixo}_{st.session_state['uploader_key']}"):
                        st.session_state[k_sel] = not st.session_state[k_sel]; st.rerun()

                    df_render.insert(0, "SELECIONAR?", st.session_state[k_sel])
                    col_cfg = gerar_config_largura_colunas(df_render, df_render.columns.tolist())
                    col_cfg["SELECIONAR?"] = st.column_config.CheckboxColumn("SELECIONAR?", default=False)

                    df_editado_res = st.data_editor(df_render, column_config=col_cfg, use_container_width=True, hide_index=True, key=f"ed_{prefixo}_{st.session_state['uploader_key']}")
                    selecionados_grupo = df_editado_res[df_editado_res["SELECIONAR?"] == True]

                    if not selecionados_grupo.empty:
                        st.markdown("---")
                        st.markdown(f"### 📋 Espaço de Visualização — {titulo}")

                        for nome_interno in selecionados_grupo["NOME"].dropna().unique():
                            df_nome_sel = selecionados_grupo[selecionados_grupo["NOME"] == nome_interno]
                            org = ", ".join([str(v) for v in df_nome_sel["ORGANIZ"].dropna().unique() if str(v).strip()])
                            fun = ", ".join([str(v) for v in df_nome_sel["FUNÇÃO"].dropna().unique() if str(v).strip()])
                            sai = ", ".join([str(v) for v in df_nome_sel["SAIDA"].dropna().unique() if str(v).strip()])
                            t_dias = int(round(sum(conv_num(v) for v in df_nome_sel["REAL"])))

                            st.markdown(f"**NOME:** {nome_interno} &nbsp;|&nbsp; **ORGANIZAÇÃO:** {org or 'N/A'} &nbsp;|&nbsp; **FUNÇÃO:** {fun or 'N/A'} &nbsp;|&nbsp; **REMUNERAÇÃO:** {cat_key} &nbsp;|&nbsp; **SAÍDA:** {sai or 'N/A'}")

                            m_data = []
                            for _, r_row in df_nome_sel.iterrows():
                                d_ma = str(r_row.get("MES/ANO - ABA", "")).split(" - ")[0]
                                ms, an = "N/A", "N/A"
                                if "/" in d_ma: ms, an = mapa_meses.get(d_ma.split("/")[0].strip().upper(), d_ma.split("/")[0].strip()), d_ma.split("/")[1].strip()
                                m_data.append({"ANO": an, "MÊS": ms, "REAL": fmt_num(r_row.get("REAL", ""))})

                            df_pivot = pd.DataFrame()
                            if m_data:
                                df_pivot = pd.DataFrame(m_data).pivot_table(index="ANO", columns="MÊS", values="REAL", aggfunc=lambda x: " / ".join([str(v) for v in x if str(v).strip()])).fillna("")
                                cols_meses = sorted(df_pivot.columns, key=lambda m: ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"].index(m) if m in ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"] else 99)
                                st.dataframe(df_pivot[cols_meses], use_container_width=True)
                                df_pivot = df_pivot[cols_meses]

                            st.markdown(f"**Total de Dias:** {t_dias}\n<br>", unsafe_allow_html=True)
                            todos_dados_exportacao.append({"nome": nome_interno, "organiz": org or "N/A", "funcao": fun or "N/A", "remuneracao": cat_key, "saida": sai or "N/A", "pivot_df": df_pivot, "total_dias": t_dias})

                        st.caption(f"📌 **{len(selecionados_grupo)}** item(ns) selecionado(s) nesta tabela.")
                        st.markdown("---")

            if todos_dados_exportacao:
                st.markdown("### 📥 Baixar ou Salvar Relatório Unificado")
                ex_bytes, doc_bytes = gerar_excel_bytes(todos_dados_exportacao), gerar_docx_bytes(todos_dados_exportacao)

                c1, c2 = st.columns(2)
                with c1: st.download_button("📊 Baixar Excel", ex_bytes, "salvamento_remicao_consolidado.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
                with c2: st.download_button("📄 Baixar Word", doc_bytes, "salvamento_remicao_consolidado.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary")

                st.markdown("#### ☁️ Salvar Diretamente no Google Drive")
                c3, c4 = st.columns(2)
                with c3:
                    if st.button("💾 Salvar Excel no Drive"):
                        with st.spinner("Enviando..."): link = salvar_arquivo_no_drive("salvamento_remicao_consolidado.xlsx", ex_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ROOT_FOLDER_ID.strip())
                        st.success(f"Salvo! [Abrir no Drive]({link})")
                with c4:
                    if st.button("💾 Salvar Word no Drive"):
                        with st.spinner("Enviando..."): link = salvar_arquivo_no_drive("salvamento_remicao_consolidado.docx", doc_bytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", ROOT_FOLDER_ID.strip())
                        st.success(f"Salvo! [Abrir no Drive]({link})")
                st.markdown("---")
        else:
            st.info("ℹ️ Nenhum registro selecionado.")

    if st.button("🗑️ Limpar Tudo"):
        c_at = st.session_state.get("uploader_key", 0) + 1
        st.session_state.clear(); st.session_state["uploader_key"] = c_at; st.rerun()
