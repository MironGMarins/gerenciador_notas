# ==============================================================================
# IMPORTS
# ==============================================================================
import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import time
from datetime import datetime, timedelta
import numpy as np 

# ==============================================================================
# CONFIGURAÇÃO DA PÁGINA
# ==============================================================================
st.set_page_config(
    layout="wide",
    page_title="Gerenciador Administrativo"
)

# ==============================================================================
# CONSTANTES
# ==============================================================================
PLANILHA_ORIGEM_NOME = "Total BaseCamp Consolidado" # Origem dos dados brutos
PLANILHA_CONSOLIDADA_NOME = "Total BaseCamp para Notas" # Destino da consolidação (Dashboard)
PLANILHA_BACKLOG_NOME = "Backlog"
PLANILHA_EQUIPES_NOME = "Equipes"
PLANILHA_SENHAS_NOME = "Senhas"
PLANILHA_HISTORICO_NOME = "HistoricoDiario"

MESES_PT_NUM = {
    'janeiro': 1, 'fevereiro': 2, 'março': 3, 'abril': 4, 'maio': 5, 'junho': 6,
    'julho': 7, 'agosto': 8, 'setembro': 9, 'outubro': 10, 'novembro': 11, 'dezembro': 12
}
MESES_NUM_PT = {v: k.capitalize() for k, v in MESES_PT_NUM.items()}

# ==============================================================================
# FUNÇÕES AUXILIARES
# ==============================================================================
def obter_nome_aba_mes_atual():
    hoje = datetime.now()
    return f"{MESES_NUM_PT[hoje.month]} {hoje.year}"

def extrair_mes_ano_da_aba(nome_aba):
    try:
        partes = nome_aba.split()
        if len(partes) == 2:
            mes_nome = partes[0].lower()
            ano_str = partes[1]
            if mes_nome in MESES_PT_NUM and ano_str.isdigit() and len(ano_str) == 4:
                return MESES_PT_NUM[mes_nome], int(ano_str)
    except: pass
    return None

def obter_lista_colunas_para_remover(spreadsheet):
    cols_to_drop = ['Peso'] 
    try:
        ws = spreadsheet.worksheet(PLANILHA_EQUIPES_NOME)
        df = pd.DataFrame(ws.get_all_records())
        if 'Posição' in df.columns and 'Nome' in df.columns:
            lideres = df[df['Posição'] == 'Lider']['Nome'].tolist()
            cols_to_drop.extend(lideres)
    except: pass
    return cols_to_drop

def regenerar_id_pelo_link(df):
    if df.empty: return df
    df.columns = df.columns.astype(str).str.strip()
    if 'Link' in df.columns:
        df['ID'] = df['Link'].astype(str).str.split('/').str[-1].str.strip()
        df = df[df['ID'] != '']
    return df

def converter_data_robusta(series):
    series = series.astype(str).str.strip()
    series = series.replace(['nan', 'None', '', 'NaT', '0', '#N/A'], np.nan)
    return pd.to_datetime(series, dayfirst=True, errors='coerce')

# ==============================================================================
# CONEXÃO E CACHE
# ==============================================================================
@st.cache_resource(ttl=600)
def autorizar_cliente():
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    try:
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
    except: 
        try:
             creds = Credentials.from_service_account_file("google_credentials.json", scopes=scopes)
        except: return None
    return gspread.authorize(creds)

@st.cache_resource(ttl=600)
def obter_spreadsheet_cacheada():
    client = autorizar_cliente()
    if not client: return None
    try: return client.open_by_url(st.secrets.get("SHEET_URL"))
    except Exception as e: st.error(f"Erro planilha: {e}"); return None

def carregar_aba_robusta(worksheet):
    for tentativa in range(3):
        try:
            all_values = worksheet.get_all_values()
            if not all_values: return pd.DataFrame()
            headers = all_values[0]; data = all_values[1:]
            cols = pd.Series(headers)
            for dup in cols[cols.duplicated()].unique():
                cols[cols[cols == dup].index.values.tolist()] = [dup + '.' + str(i) if i != 0 else dup for i in range(sum(cols == dup))]
            df = pd.DataFrame(data, columns=cols)
            df = regenerar_id_pelo_link(df)
            return df
        except gspread.exceptions.APIError as e:
            if "429" in str(e): time.sleep(2 * (tentativa + 1)); continue
            else: return pd.DataFrame()
        except: return pd.DataFrame()
    return pd.DataFrame()

# ==============================================================================
# LOGIN
# ==============================================================================
@st.cache_data(ttl=60)
def check_credentials(username, password):
    try:
        spreadsheet = obter_spreadsheet_cacheada()
        ws_senhas = spreadsheet.worksheet(PLANILHA_SENHAS_NOME)
        df_senhas = pd.DataFrame(ws_senhas.get_all_records())
        u = str(username).strip(); p = str(password).strip()
        df_senhas['Usuario'] = df_senhas['Usuario'].astype(str).str.strip()
        df_senhas['Senha'] = df_senhas['Senha'].astype(str).str.strip()
        match = df_senhas[(df_senhas['Usuario'] == u) & (df_senhas['Senha'] == p)]
        if not match.empty: return match.iloc[0].get('Status', 'Visualizador')
        return None
    except: return None

# ==============================================================================
# AÇÕES DO SISTEMA
# ==============================================================================
def sincronizar_basecamp_com_mes_especifico(nome_aba_destino):
    """
    Copia dados da Planilha Base (Total BaseCamp Consolidado) para uma aba de mês específica.
    IGNORA tarefas [ARCHIVED].
    Mantém filtro de datas para popular a aba do mês corretamente (Mês Atual vs Histórico).
    """
    spreadsheet = obter_spreadsheet_cacheada()
    
    mes_ano = extrair_mes_ano_da_aba(nome_aba_destino)
    if not mes_ano: return f"Nome da aba '{nome_aba_destino}' inválido."
    mes_alvo, ano_alvo = mes_ano

    try:
        ws_origem = spreadsheet.worksheet(PLANILHA_ORIGEM_NOME)
        df_origem = carregar_aba_robusta(ws_origem)
    except: return f"Aba Origem '{PLANILHA_ORIGEM_NOME}' não encontrada."
    if df_origem.empty: return "Origem vazia."

    # --- FILTRO DE EXCLUSÃO DE ARQUIVADAS ---
    df_origem.columns = df_origem.columns.astype(str).str.strip()
    if 'Lista' in df_origem.columns:
        # Remove linhas que contém "[ARCHIVED]"
        df_origem = df_origem[~df_origem['Lista'].astype(str).str.contains("\[ARCHIVED\]", case=False, regex=True, na=False)]
    # ----------------------------------------

    try:
        ws_destino = spreadsheet.worksheet(nome_aba_destino)
        ws_destino.clear()
    except gspread.exceptions.WorksheetNotFound:
        try: ws_destino = spreadsheet.add_worksheet(title=nome_aba_destino, rows=1000, cols=30)
        except Exception as e: return f"Erro criar aba: {e}"

    # Lógica de Datas para popular a aba do mês
    if 'Data Final' in df_origem.columns:
        df_origem['Data_Obj'] = converter_data_robusta(df_origem['Data Final'])
        hoje = datetime.now()
        data_aba = datetime(ano_alvo, mes_alvo, 1)
        data_ref_servidor = datetime(hoje.year, hoje.month, 1)
        
        eh_mes_relevante = data_aba >= data_ref_servidor
        
        if eh_mes_relevante:
            # Mês Atual/Futuro: Pega o mês + Backlog (sem data)
            condicao = (
                ((df_origem['Data_Obj'].dt.month == mes_alvo) & (df_origem['Data_Obj'].dt.year == ano_alvo)) |
                (df_origem['Data_Obj'].isna())
            )
        else:
            # Mês Passado: Apenas o mês exato
            condicao = (df_origem['Data_Obj'].dt.month == mes_alvo) & (df_origem['Data_Obj'].dt.year == ano_alvo)
        
        df_final = df_origem[condicao].copy()
        df_final = df_final.drop(columns=['Data_Obj'])
    else:
        df_final = df_origem.copy()

    cols_drop = obter_lista_colunas_para_remover(spreadsheet)
    df_final = df_final.drop(columns=[c for c in cols_drop if c in df_final.columns], errors='ignore').fillna('')
    
    try:
        ws_destino.update([df_final.columns.values.tolist()] + df_final.astype(str).values.tolist(), value_input_option='USER_ENTERED')
        return "Sucesso"
    except Exception as e: return f"Erro salvar: {e}"

def atualizar_aba_backlog():
    """
    Lê a origem, IGNORA ARQUIVADAS, filtra 'Backlog' na coluna Lista e salva.
    """
    spreadsheet = obter_spreadsheet_cacheada()
    
    try:
        ws_origem = spreadsheet.worksheet(PLANILHA_ORIGEM_NOME)
        df_origem = carregar_aba_robusta(ws_origem)
    except: return f"Aba Origem '{PLANILHA_ORIGEM_NOME}' não encontrada."
    
    if df_origem.empty: return "Origem vazia."
    
    df_origem.columns = df_origem.columns.astype(str).str.strip()
    
    if 'Lista' not in df_origem.columns: return "Coluna 'Lista' não encontrada na origem."

    # --- FILTRO DE EXCLUSÃO DE ARQUIVADAS ---
    df_origem = df_origem[~df_origem['Lista'].astype(str).str.contains("\[ARCHIVED\]", case=False, regex=True, na=False)]
    # ----------------------------------------
        
    # Filtra Backlog
    mask_backlog = df_origem['Lista'].astype(str).str.contains("Backlog", case=False, na=False)
    df_backlog = df_origem[mask_backlog].copy()
    
    cols_drop = obter_lista_colunas_para_remover(spreadsheet)
    df_backlog = df_backlog.drop(columns=[c for c in cols_drop if c in df_backlog.columns], errors='ignore').fillna('')
    
    try:
        try: 
            ws_backlog = spreadsheet.worksheet(PLANILHA_BACKLOG_NOME)
            ws_backlog.clear()
        except gspread.exceptions.WorksheetNotFound:
            ws_backlog = spreadsheet.add_worksheet(title=PLANILHA_BACKLOG_NOME, rows=1000, cols=30)
            
        ws_backlog.update([df_backlog.columns.values.tolist()] + df_backlog.astype(str).values.tolist(), value_input_option='USER_ENTERED')
        return f"Sucesso! {len(df_backlog)} tarefas no Backlog."
    except Exception as e: return f"Erro ao salvar Backlog: {e}"

def consolidar_geral_para_dashboard():
    """
    Consolida TODAS as abas de MESES em um 'Mapa Histórico'.
    Lógica: EMPILHAMENTO SIMPLES (SNAPSHOT).
    - Ignora datas (assume que se está na aba do mês, pertence àquele histórico).
    - Permite repetições (mesma tarefa pode aparecer em Nov e Dez para mostrar evolução).
    - Remove [ARCHIVED] para limpeza.
    """
    spreadsheet = obter_spreadsheet_cacheada()
    all_worksheets = spreadsheet.worksheets()
    dfs = []
    
    for ws in all_worksheets:
        mes_ano = extrair_mes_ano_da_aba(ws.title)
        if mes_ano:
            time.sleep(1.5)
            
            df_mes = carregar_aba_robusta(ws)
            if df_mes.empty: continue
            
            df_mes.columns = df_mes.columns.astype(str).str.strip()
            
            # --- FILTRO DE EXCLUSÃO DE ARQUIVADAS ---
            if 'Lista' in df_mes.columns:
                df_mes = df_mes[~df_mes['Lista'].astype(str).str.contains("\[ARCHIVED\]", case=False, regex=True, na=False)]
            
            # Identifica a fonte do snapshot (ex: "Snapshot: Novembro 2024")
            df_mes['Fonte_Dados'] = f"Snapshot: {ws.title}"
            
            # ADICIONA AO CONSOLIDADO SEM FILTRO DE DATA
            dfs.append(df_mes)

    if not dfs: return "Nenhum dado (aba mensal) encontrado para consolidar."

    df_final = pd.concat(dfs, ignore_index=True)
    cols_drop = obter_lista_colunas_para_remover(spreadsheet)
    df_final = df_final.drop(columns=[c for c in cols_drop if c in df_final.columns], errors='ignore').fillna('')

    try:
        try: ws_final = spreadsheet.worksheet(PLANILHA_CONSOLIDADA_NOME)
        except: ws_final = spreadsheet.add_worksheet(title=PLANILHA_CONSOLIDADA_NOME, rows=2000, cols=30)
        
        ws_final.clear()
        df_save = df_final
        df_list = [df_save.columns.values.tolist()] + df_save.astype(str).values.tolist()
        ws_final.update(df_list, value_input_option='USER_ENTERED')
        return f"Sucesso! {len(df_save)} tarefas consolidadas (snapshots) na aba '{PLANILHA_CONSOLIDADA_NOME}'."
    except Exception as e: return f"Erro salvar: {e}"

def atualizar_historico_diario():
    """
    Lê da aba 'Total BaseCamp Consolidado' e filtra pela semana atual na coluna 'Lista'.
    IGNORA tarefas arquivadas.
    """
    try:
        hoje = pd.Timestamp.now().normalize()
        inicio_sem = hoje - timedelta(days=hoje.dayofweek)
        data_ref_lista_str = inicio_sem.strftime('%d/%m/%Y')
        
        spreadsheet = obter_spreadsheet_cacheada()
        
        try: 
            ws_src = spreadsheet.worksheet(PLANILHA_ORIGEM_NOME)
            df_src = carregar_aba_robusta(ws_src)
        except: return f"Aba '{PLANILHA_ORIGEM_NOME}' não encontrada."
        
        if df_src.empty: return "Aba de origem vazia."
        df_src.columns = df_src.columns.astype(str).str.strip()

        if 'Lista' not in df_src.columns: return "Coluna 'Lista' ausente."

        # Ignora Arquivadas
        df_src = df_src[~df_src['Lista'].astype(str).str.contains("\[ARCHIVED\]", case=False, regex=True, na=False)]
            
        # Filtra pela Lista da Semana
        df_semana = df_src[df_src['Lista'].astype(str).str.contains(data_ref_lista_str, na=False, regex=False)]
        total_sem = len(df_semana)
        
        if 'Data Final' not in df_semana.columns: df_semana['Data Final'] = pd.NaT
        else: df_semana['Data Final'] = converter_data_robusta(df_semana['Data Final'])

        fechadas_sem = df_semana[df_semana['Data Final'].notna()].shape[0]
        
        try: ws_hist = spreadsheet.worksheet(PLANILHA_HISTORICO_NOME)
        except: 
            ws_hist = spreadsheet.add_worksheet(title=PLANILHA_HISTORICO_NOME, rows=1000, cols=3)
            ws_hist.append_row(["Data", "Total_Fechadas", "Total_Tarefas"])
        
        hoje_str = hoje.strftime('%d/%m/%Y')
        linha = [hoje_str, int(fechadas_sem), int(total_sem)]
        
        try: 
            cell = ws_hist.find(hoje_str, in_column=1)
            ws_hist.update(f'A{cell.row}:C{cell.row}', [linha], value_input_option='USER_ENTERED')
        except: 
            ws_hist.append_row(linha, value_input_option='USER_ENTERED')
            
        return f"OK! Semana {data_ref_lista_str}: {fechadas_sem}/{total_sem}"
        
    except Exception as e: return f"Erro: {e}"

def deletar_tarefa_global(id_del):
    spreadsheet = obter_spreadsheet_cacheada()
    id_del = str(id_del).strip()
    try:
        try: 
            ws = spreadsheet.worksheet(obter_nome_aba_mes_atual())
            df = carregar_aba_robusta(ws)
            if 'ID' in df.columns:
                df_n = df[df['ID'] != id_del].fillna('')
                if len(df_n) < len(df): 
                    ws.clear()
                    ws.update([df_n.columns.values.tolist()] + df_n.astype(str).values.tolist(), value_input_option='USER_ENTERED')
        except: pass
        
        ws = spreadsheet.worksheet(PLANILHA_ORIGEM_NOME)
        df = carregar_aba_robusta(ws)
        if 'ID' in df.columns:
            df_n = df[df['ID'] != id_del].fillna('')
            if len(df_n) < len(df): 
                ws.clear()
                ws.update([df_n.columns.values.tolist()] + df_n.astype(str).values.tolist(), value_input_option='USER_ENTERED')
                return True
        return False
    except: return False

# ==============================================================================
# DIAGNÓSTICO (ATUALIZADO PARA MOSTRAR NÃO-VAZIAS)
# ==============================================================================
def diagnostico_datas(nome_aba_destino):
    """Função para depurar dados da Origem."""
    spreadsheet = obter_spreadsheet_cacheada()
    mes_ano = extrair_mes_ano_da_aba(nome_aba_destino)
    
    st.markdown(f"### 🔧 Diagnóstico Avançado para '{nome_aba_destino}'")
    st.write(f"**Horário do Servidor (UTC):** {datetime.now()}")
    
    if not mes_ano:
        st.error("Nome da aba inválido.")
        return
        
    mes_alvo, ano_alvo = mes_ano
    
    ws_origem = spreadsheet.worksheet(PLANILHA_ORIGEM_NOME)
    df_origem = carregar_aba_robusta(ws_origem)
    st.write(f"**Linhas na Origem:** {len(df_origem)}")
    
    # DIAGNÓSTICO DE ARQUIVADAS
    if 'Lista' in df_origem.columns:
        qtd_arq = df_origem['Lista'].astype(str).str.contains("\[ARCHIVED\]", case=False, regex=True, na=False).sum()
        st.warning(f"**Linhas detectadas como [ARCHIVED]:** {qtd_arq} (Estas serão ignoradas)")
        
        # APLICA O FILTRO PARA O DIAGNÓSTICO SER REALISTA
        df_origem = df_origem[~df_origem['Lista'].astype(str).str.contains("\[ARCHIVED\]", case=False, regex=True, na=False)]
        st.success(f"**Linhas ATIVAS (para análise):** {len(df_origem)}")
    
    if 'Data Final' in df_origem.columns:
        # Mostra apenas valores preenchidos
        df_preenchida = df_origem[df_origem['Data Final'].astype(str).str.strip() != '']
        st.write(f"**Total de datas preenchidas (ATIVAS):** {len(df_preenchida)}")
        
        if not df_preenchida.empty:
            st.write("Amostra de 'Data Final' (RAW - Preenchidas):")
            st.dataframe(df_preenchida['Data Final'].head(5))
            
            # Converte
            df_preenchida['Data_Obj'] = converter_data_robusta(df_preenchida['Data Final'])
            
            # FALHAS DE CONVERSÃO
            falhas = df_preenchida[df_preenchida['Data_Obj'].isna()]
            if not falhas.empty:
                st.error(f"**ERRO DE CONVERSÃO:** {len(falhas)} datas não puderam ser lidas.")
                st.write("Amostra de datas inválidas:")
                st.dataframe(falhas['Data Final'].head(5))
            
            st.write("Amostra de 'Data Final' (CONVERTIDA):")
            st.dataframe(df_preenchida['Data_Obj'].head(5))
            
            # Verifica quantas entrariam no filtro do mês
            no_mes = df_preenchida[
                (df_preenchida['Data_Obj'].dt.month == mes_alvo) & 
                (df_preenchida['Data_Obj'].dt.year == ano_alvo)
            ]
            st.write(f"**Linhas identificadas neste mês ({mes_alvo}/{ano_alvo}):** {len(no_mes)}")
            
            if not no_mes.empty:
                st.write("Exemplo de tarefas encontradas:")
                cols_show = ['ID', 'Nome Task', 'Data Final'] if 'Nome Task' in no_mes.columns else ['ID', 'Data Final']
                st.dataframe(no_mes[cols_show].head(5), hide_index=True)
        else:
            st.warning("A coluna 'Data Final' parece estar vazia em todas as linhas.")

        sem_data = df_origem[df_origem['Data Final'].astype(str).str.strip() == '']
        st.write(f"**Linhas SEM data (Backlog):** {len(sem_data)}")

    else:
        st.error("Coluna 'Data Final' não encontrada!")

# ==============================================================================
# INTERFACE
# ==============================================================================
if "authenticated" not in st.session_state: st.session_state.authenticated = False
if "user_role" not in st.session_state: st.session_state.user_role = None
if "id_para_buscar" not in st.session_state: st.session_state.id_para_buscar = ""

if not st.session_state.authenticated:
    st.title(f"🔒 Login - Gestão {obter_nome_aba_mes_atual()}")
    with st.form("login"):
        u = st.text_input("Usuário")
        p = st.text_input("Senha", type="password")
        if st.form_submit_button("Entrar"):
            role = check_credentials(u, p)
            if role: 
                st.session_state.authenticated = True
                st.session_state.user_role = role
                st.rerun()
            else: st.error("Acesso negado.")
else:
    aba_atual = obter_nome_aba_mes_atual()
    spreadsheet = obter_spreadsheet_cacheada()
    
    col1, col2 = st.columns([3,1])
    col1.title("📝 Gerenciador Administrativo")
    if col2.button("Sair"): 
        st.session_state.authenticated = False
        st.rerun()

    with st.sidebar:
        st.image("media portal logo.png", width=200)
        st.header("Ações")
        if st.session_state.user_role == "Editor":
            st.info("⚠️ Ações de Escrita")
            
            meses_opcoes = [f"{MESES_NUM_PT[m]} {datetime.now().year}" for m in range(1, 13)]
            idx_atual = datetime.now().month - 1
            aba_selecionada = st.selectbox("Aba de Mês para Atualizar (Base -> Mês):", meses_opcoes, index=idx_atual)
            
            if st.button(f"1. Atualizar Aba '{aba_selecionada}'"):
                with st.spinner("Sincronizando..."):
                    res = sincronizar_basecamp_com_mes_especifico(aba_selecionada)
                    if res == "Sucesso": st.success(f"Aba '{aba_selecionada}' atualizada com sucesso!")
                    else: st.error(res)
            
            st.markdown("---")
            if st.button("2. Consolidar DashBoard (Meses -> Consolidado)"):
                with st.spinner("Consolidando histórico..."):
                    res = consolidar_geral_para_dashboard()
                    if "Sucesso" in res: st.success(res)
                    elif "Nenhum" in res: st.warning(res)
                    else: st.error(res)
            
            if st.button("3. Snapshot Gráfico (Semana Atual)"):
                with st.spinner("Lendo Origem e salvando histórico..."):
                    res = atualizar_historico_diario()
                    if "OK" in res: st.success(res)
                    else: st.error(res)
                    
            if st.button("4. Atualizar Backlog"):
                with st.spinner("Atualizando Backlog..."):
                    res = atualizar_aba_backlog()
                    if "Sucesso" in res: st.success(res)
                    else: st.error(res)
            
            st.markdown("---")
            with st.expander("🔧 Diagnóstico de Dados (Debug)"):
                 if st.button("Rodar Diagnóstico"):
                     diagnostico_datas(aba_selecionada)

            st.markdown("---")
            st.subheader("Deletar Tarefa")
            id_del = st.text_input("ID para deletar")
            if st.button("Confirmar Deleção"):
                if id_del:
                    if deletar_tarefa_global(id_del): 
                        st.success(f"Tarefa {id_del} deletada!")
                        time.sleep(1)
                        st.rerun()
                    else: st.error("ID não encontrado na origem.")
                else: st.warning("Digite o ID.")

    st.info(f"Visualizando dados da aba: **{aba_atual}**")
    try:
        if spreadsheet:
            try:
                ws_atual = spreadsheet.worksheet(aba_atual)
                df = carregar_aba_robusta(ws_atual)
                
                if not df.empty:
                    col_filt, col_search = st.columns(2)
                    filtro = col_filt.multiselect("Filtrar por Encarregado", ["Todos"] + sorted(df['Encarregado'].astype(str).unique()), default="Todos")
                    busca = col_search.text_input("Buscar ID", value=st.session_state.id_para_buscar)
                    
                    if busca: 
                        st.session_state.id_para_buscar = busca
                        df = df[df['ID'] == busca]
                    elif "Todos" not in filtro: 
                        df = df[df['Encarregado'].isin(filtro)]
                    
                    st.dataframe(df, use_container_width=True, hide_index=True)
                    st.caption(f"Total de linhas visualizadas: {len(df)}")
                else: st.warning("Aba vazia.")
            except gspread.exceptions.WorksheetNotFound:
                st.warning(f"A aba '{aba_atual}' ainda não existe. Clique em 'Atualizar Mês' para criá-la.")
    except Exception as e:
        st.error(f"Erro de conexão: {e}")