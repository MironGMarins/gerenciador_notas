# ==============================================================================
# IMPORTS
# ==============================================================================
import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import io
import requests # Adicionado para chamadas de API
import numpy as np # Adicionado para np.where e outros
from datetime import datetime, timedelta # Necessário para a nova função

# ==============================================================================
# CONFIGURAÇÃO DA PÁGINA
# ==============================================================================
st.set_page_config(
    layout="wide",
    page_title="Gerenciador de Notas de Tarefas"
)

# ==============================================================================
# NOME DAS PLANILHAS (CONFIGURÁVEL)
# ==============================================================================
PLANILHA_ORIGEM_NOME = "Total BaseCamp"
PLANILHA_NOTAS_NOME = "Total BaseCamp para Notas"
PLANILHA_EQUIPES_NOME = "Equipes"
PLANILHA_SENHAS_NOME = "Senhas" # <-- Nova aba para login
PLANILHA_HISTORICO_NOME = "HistoricoDiario" # <-- NOVO PARA DASHBOARD

# ==============================================================================
# AUTENTICAÇÃO E CONEXÃO
# ==============================================================================
@st.cache_resource(ttl=600)
def autorizar_cliente():
    """Autoriza e retorna o cliente gspread."""
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    try:
        creds_json = st.secrets["gcp_service_account"]
        creds = Credentials.from_service_account_info(creds_json, scopes=scopes)
    except (FileNotFoundError, KeyError):
        st.error("Credenciais do Google (gcp_service_account) não encontradas nos segredos do Streamlit.")
        return None
    return gspread.authorize(creds)

# ==============================================================================
# FUNÇÃO AUXILIAR DE CARREGAMENTO ROBUSTO
# ==============================================================================
def carregar_aba_de_forma_robusta(worksheet):
    """Carrega uma aba mesmo que ela tenha cabeçalhos duplicados ou esteja vazia."""
    all_values = worksheet.get_all_values()
    if not all_values:
        # Retorna DataFrame vazio com colunas se a aba tiver só cabeçalho
        try:
            headers = worksheet.row_values(1)
            if headers:
                return pd.DataFrame(columns=headers)
        except Exception:
            pass # Ignora se não conseguir ler nem o cabeçalho
        return pd.DataFrame()
    
    headers = all_values[0]
    data = all_values[1:]
    
    # Tratamento para cabeçalhos duplicados (adiciona sufixo)
    cols = pd.Series(headers)
    for dup in cols[cols.duplicated()].unique():
        cols[cols[cols == dup].index.values.tolist()] = [dup + '.' + str(i) if i != 0 else dup for i in range(sum(cols == dup))]
    
    df = pd.DataFrame(data, columns=cols)
    # Remove colunas que são totalmente vazias (sem nome de cabeçalho)
    df = df.loc[:, df.columns.notna() & (df.columns != '')]
    return df


# ==============================================================================
# --- FUNÇÃO DE LOGIN (ATUALIZADA PARA RETORNAR A FUNÇÃO/STATUS) ---
# ==============================================================================
@st.cache_data(ttl=60) # Cache curto para verificação de senha
def check_credentials(username, password):
    """Verifica o usuário e senha e retorna a função (Status) do usuário."""
    try:
        client = autorizar_cliente()
        if client is None: return None
        
        url_da_planilha = st.secrets.get("SHEET_URL")
        if not url_da_planilha:
            st.error("URL da planilha (SHEET_URL) não encontrada nos segredos.")
            return None
            
        spreadsheet = client.open_by_url(url_da_planilha)
        ws_senhas = spreadsheet.worksheet(PLANILHA_SENHAS_NOME)
        
        # Usa get_all_records para simplicidade, já que a aba Senhas é controlada
        df_senhas = pd.DataFrame(ws_senhas.get_all_records())
        
        username = str(username).strip()
        password = str(password).strip()
        
        df_senhas['Usuario'] = df_senhas['Usuario'].astype(str).str.strip()
        df_senhas['Senha'] = df_senhas['Senha'].astype(str).str.strip()

        match = df_senhas[
            (df_senhas['Usuario'] == username) & 
            (df_senhas['Senha'] == password)
        ]
        
        if not match.empty:
            # Retorna a função (Status) do usuário. Padrão para "Visualizador" se a coluna estiver vazia.
            role = match.iloc[0].get('Status', 'Visualizador')
            if role == "": # Trata caso de célula vazia
                return "Visualizador"
            return role
        else:
            return None # Retorna None se a senha ou usuário estiverem incorretos
    except gspread.exceptions.WorksheetNotFound:
        st.error(f"Aba de '{PLANILHA_SENHAS_NOME}' não encontrada. Verifique sua planilha.")
        return None
    except Exception as e:
        st.error(f"Erro ao verificar credenciais: {e}")
        return None

# ==============================================================================
# --- FUNÇÃO DA API DO BASECAMP ---
# ==============================================================================
def atualizar_basecamp_api():
    """
    Função para buscar dados da API do Basecamp e atualizar a planilha de origem.
    """
    try:
        # ==========================================================================
        ### COLE SEU CÓDIGO DA API DO BASECAMP AQUI ###
        st.info("Simulando chamada da API do Basecamp...")
        nova_tarefa_exemplo = {
            "ID": [str(pd.Timestamp.now().timestamp()).replace('.', '')], "Tarefa": ["Nova Tarefa via API"],
            "Encarregado": ["OCTAVIO"], "Data Inicial": [pd.Timestamp.now().strftime('%Y-%m-%d')],
            "Data Final": [""], "Data Estipulada": [""],
            "Link": ["https://3.basecamp.com/.../todos/NOVOTODOID"],
            "Observação": [""], "Peso": [""], "Pablo": [""], "Leonardo": [""], "Itiel": [""], "Ítalo": [""]
        }
        df_novas_tarefas = pd.DataFrame(nova_tarefa_exemplo)
        # ==========================================================================

        client = autorizar_cliente()
        url_da_planilha = st.secrets.get("SHEET_URL")
        if not url_da_planilha:
            st.error("URL da planilha não configurada nos segredos.")
            return False
            
        spreadsheet = client.open_by_url(url_da_planilha)
        worksheet_origem = spreadsheet.worksheet(PLANILHA_ORIGEM_NOME)
        worksheet_origem.append_rows(df_novas_tarefas.values.tolist(), value_input_option='USER_ENTERED')
        return True

    except Exception as e:
        st.error(f"Erro ao atualizar dados do BaseCamp: {e}")
        return False
        
# ==============================================================================
# --- ATUALIZAR HISTÓRICO DIÁRIO (LÓGICA DA SEMANA ATUAL) ---
# ==============================================================================
def atualizar_historico_diario():
    """
    Calcula o total de tarefas e o total de tarefas fechadas DA SEMANA ATUAL
    da aba 'Total BaseCamp para Notas' e salva/atualiza o registro de HOJE
    na aba 'HistoricoDiario'.
    """
    try:
        # --- 1. Checagem de Dia da Semana (Seg-Sex) ---
        hoje = pd.Timestamp.now().normalize()
        dia_da_semana_iso = hoje.dayofweek # 0=Seg, 6=Dom
        
        if dia_da_semana_iso > 4: # Se for Sábado (5) ou Domingo (6)
            st.warning("O snapshot de histórico só é salvo de Segunda a Sexta-feira.")
            return False

        # --- 2. Carregar Dados Fonte ('Total BaseCamp para Notas') ---
        client = autorizar_cliente()
        if client is None: 
            st.error("Falha na autenticação.")
            return False
        
        url_da_planilha = st.secrets.get("SHEET_URL")
        if not url_da_planilha:
            st.error("URL da planilha não configurada nos segredos.")
            return False
            
        spreadsheet = client.open_by_url(url_da_planilha)
        # --- MUDANÇA: Fonte de dados agora é a 'Notas' ---
        ws_source = spreadsheet.worksheet(PLANILHA_NOTAS_NOME) 
        df_source = carregar_aba_de_forma_robusta(ws_source)

        if df_source.empty:
            st.warning(f"A aba '{PLANILHA_NOTAS_NOME}' está vazia. Nenhum histórico salvo.")
            return False

        # --- 3. Processar Datas e Adicionar Calendário ---
        df_source['Data Inicial'] = pd.to_datetime(df_source['Data Inicial'], dayfirst=True, errors='coerce')
        df_source['Data Final'] = pd.to_datetime(df_source['Data Final'], dayfirst=True, errors='coerce')
        df_source['Status_Tarefa'] = np.where(df_source['Data Final'].isnull(), 'Aberto', 'Executado')
        df_source['Data Final (aberta)'] = df_source['Data Final'].fillna(hoje)

        # Cria a "Tabela Calendário" necessária
        data_inicio_calendario = df_source['Data Inicial'].min() if pd.notna(df_source['Data Inicial'].min()) else hoje
        tabela_calendario = pd.DataFrame({"Date": pd.date_range(start=data_inicio_calendario, end=hoje, freq='D')})
        tabela_calendario['Dia da Semana_ISO'] = tabela_calendario['Date'].dt.dayofweek
        tabela_calendario['Data_Inicio_Semana'] = tabela_calendario['Date'] - pd.to_timedelta(tabela_calendario['Dia da Semana_ISO'], unit='d')
        tabela_calendario['Data_Sexta_Feira'] = tabela_calendario['Data_Inicio_Semana'] + pd.to_timedelta(4, unit='d')
        tabela_calendario['Semana_Ano'] = tabela_calendario['Data_Sexta_Feira'].dt.strftime('%Y-%U')

        # Junta as tarefas com o calendário
        df_analise = pd.merge(df_source, tabela_calendario, how='left', left_on='Data Final (aberta)', right_on='Date').drop(columns=['Date'])

        # --- 4. Fazer os Cálculos da Semana Atual ---
        
        # Encontra a "Semana_Ano" de hoje
        dias_para_sexta = (4 - dia_da_semana_iso + 7) % 7
        sexta_desta_semana = hoje + pd.to_timedelta(dias_para_sexta, unit='d')
        semana_ano_atual = sexta_desta_semana.strftime('%Y-%U')
        
        # Filtra o dataframe para incluir apenas tarefas desta semana
        df_semana_atual = df_analise[df_analise['Semana_Ano'] == semana_ano_atual].copy()

        if df_semana_atual.empty:
            st.warning("Nenhuma tarefa (aberta ou fechada) encontrada para a semana atual.")
            return False

        total_tarefas = len(df_semana_atual)
        total_fechadas = len(df_semana_atual[df_semana_atual['Status_Tarefa'] == 'Executado'])
        
        # --- 5. Preparar a nova linha ---
        hoje_str = hoje.strftime('%d/%m/%Y') # Formato PT-BR
        
        # Colunas: "Data Final", "Total_Fechadas", "Total_Tarefas"
        nova_linha = [hoje_str, total_fechadas, total_tarefas]
        
        # 6. Salvar ou Atualizar na aba "HistoricoDiario"
        ws_historico = spreadsheet.worksheet(PLANILHA_HISTORICO_NOME)
        
        try:
            # Procura pela data de hoje na primeira coluna ("Data Final")
            cell = ws_historico.find(hoje_str, in_column=1) 
        except gspread.exceptions.CellNotFound:
            cell = None
        except Exception as e:
            st.error(f"Erro ao procurar data: {e}")
            return False

        if cell:
            # ==============================================================================
            # --- CORREÇÃO: Usa .update() com A1 notation, não .update_row() ---
            # ==============================================================================
            # Assume 3 colunas: A, B, C
            range_para_atualizar = f'A{cell.row}:C{cell.row}'
            ws_historico.update(range_para_atualizar, [nova_linha], value_input_option='USER_ENTERED')
            # ==============================================================================
            # --- FIM DA CORREÇÃO ---
            # ==============================================================================
            st.success(f"Histórico da semana ATUALIZADO: {hoje_str} - {total_tarefas} Tarefas Totais, {total_fechadas} Fechadas.")
        else:
            # SE NÃO ENCONTROU: Adiciona uma nova linha
            ws_historico.append_row(nova_linha, value_input_option='USER_ENTERED')
            st.success(f"Novo histórico da semana SALVO: {hoje_str} - {total_tarefas} Tarefas Totais, {total_fechadas} Fechadas.")
        
        return True

    except gspread.exceptions.WorksheetNotFound as e:
        st.error(f"Aba não encontrada: {e}. Verifique os nomes: '{PLANILHA_HISTORICO_NOME}' e '{PLANILHA_NOTAS_NOME}'.")
        return False
    except Exception as e:
        st.error(f"Erro ao salvar histórico: {e}")
        return False
# ==============================================================================
# --- FIM DA FUNÇÃO ATUALIZADA ---
# ==============================================================================
        
# ==============================================================================
# FUNÇÕES DE LÓGICA DE NEGÓCIO
# ==============================================================================
@st.cache_data(ttl=60)
def _carregar_dados_brutos():
    """Função em cache APENAS para ler os dados brutos das planilhas."""
    client = autorizar_cliente()
    if client is None: return None, None, None

    url_da_planilha = st.secrets.get("SHEET_URL")
    if not url_da_planilha:
        st.error("A URL da planilha (SHEET_URL) não foi encontrada nos segredos do Streamlit.")
        return None, None, None

    try:
        spreadsheet = client.open_by_url(url_da_planilha)
        ws_origem = spreadsheet.worksheet(PLANILHA_ORIGEM_NOME)
        ws_notas = spreadsheet.worksheet(PLANILHA_NOTAS_NOME)
        ws_equipes = spreadsheet.worksheet(PLANILHA_EQUIPES_NOME)

        df_origem = carregar_aba_de_forma_robusta(ws_origem)
        df_notas = carregar_aba_de_forma_robusta(ws_notas)
        df_equipes = carregar_aba_de_forma_robusta(ws_equipes)
        
        return df_origem, df_notas, df_equipes
        
    except gspread.exceptions.WorksheetNotFound as e:
        st.error(f"Erro Crítico: Aba não encontrada: {e}")
        return None, None, None
    except Exception as e:
        st.error(f"Erro ao carregar dados brutos: {e}")
        return None, None, None


def salvar_df_na_aba(aba_nome, df, show_success=True):
    """Função robusta para salvar: obtém uma nova conexão."""
    try:
        client = autorizar_cliente()
        url_da_planilha = st.secrets.get("SHEET_URL")
        if not url_da_planilha:
            st.error("URL da planilha não configurada nos segredos para salvar.")
            return False
        
        spreadsheet = client.open_by_url(url_da_planilha)
        worksheet = spreadsheet.worksheet(aba_nome)

        df_filled = df.fillna('')
        
        for col in df_filled.select_dtypes(include=['datetime', 'datetimetz']).columns:
             df_filled[col] = df_filled[col].dt.strftime('%d/%m/%Y')
        
        float_cols = df_filled.select_dtypes(include=['float', 'float64']).columns
        for col in float_cols:
            df_filled[col] = df_filled[col].astype(str).str.replace('.', ',')
        
        if 'ID' in df_filled.columns:
            df_filled['ID'] = df_filled['ID'].astype(str)
             
        df_list = [df_filled.columns.values.tolist()] + df_filled.astype(str).values.tolist()
        worksheet.clear()
        worksheet.update(df_list, value_input_option='USER_ENTERED')
        
        if show_success:
            st.toast("✅ Alterações salvas com sucesso na Planilha Google!")
        return True
    except Exception as e:
        st.error(f"Erro ao salvar na planilha: {e}")
        return False

# ==============================================================================
# --- FUNÇÃO 'sincronizar_e_processar_dados' ATUALIZADA ---
# ==============================================================================
def sincronizar_e_processar_dados():
    """Usa dados em cache para processar, sincronizar e retornar os DataFrames finais."""
    try:
        df_origem, df_notas, df_equipes = _carregar_dados_brutos()
        
        if df_notas is None or df_origem is None or df_equipes is None:
             st.error("Falha ao carregar uma ou mais abas. Sincronização cancelada.")
             return None, None, None, None, pd.DataFrame(), pd.DataFrame()

        if 'Encarregado' in df_origem.columns: df_origem['Encarregado'] = df_origem['Encarregado'].astype(str)
        if 'Encarregado' in df_notas.columns: df_notas['Encarregado'] = df_notas['Encarregado'].astype(str).str.strip()
        if 'Nome' in df_equipes.columns: df_equipes['Nome'] = df_equipes['Nome'].astype(str).str.strip()

        colunas_data = ['Data Inicial', 'Data Final', 'Data Estipulada']
        for col in colunas_data:
            if col in df_notas.columns:
                df_notas[col] = df_notas[col].replace(['', 'None', None, 'nan'], pd.NA)
                df_notas[col] = pd.to_datetime(df_notas[col], dayfirst=True, errors='coerce')
            else:
                df_notas[col] = pd.NaT 
            
            if col in df_origem.columns:
                df_origem[col] = df_origem[col].replace(['', 'None', None, 'nan'], pd.NA)
                df_origem[col] = pd.to_datetime(df_origem[col], dayfirst=True, errors='coerce')
            else:
                df_origem[col] = pd.NaT

        if 'Link' in df_origem.columns:
            df_origem['ID'] = df_origem['Link'].str.split('/').str[-1].fillna('').astype(str)
        else:
            st.error("Coluna 'Link' não encontrada na aba 'Origem'.")
            df_origem['ID'] = None

        if 'Link' in df_notas.columns:
            df_notas['ID'] = df_notas['Link'].str.split('/').str[-1].fillna('').astype(str)
        else:
            st.error("Coluna 'Link' não encontrada na aba 'Notas'.")
            df_notas['ID'] = None

        df_notas.dropna(subset=['ID'], inplace=True); df_notas = df_notas[df_notas['ID'] != '']
        if df_notas['ID'].duplicated().any(): df_notas.drop_duplicates(subset=['ID'], keep='first', inplace=True, ignore_index=True)
        
        df_origem.dropna(subset=['ID'], inplace=True); df_origem = df_origem[df_origem['ID'] != '']
        if df_origem['ID'].duplicated().any(): df_origem.drop_duplicates(subset=['ID'], keep='first', inplace=True, ignore_index=True)

        mudancas_detectadas = False

        if 'ID' in df_origem.columns and 'ID' in df_notas.columns and 'Link' in df_origem.columns and 'Link' in df_notas.columns:
            origem_ids = set(df_origem['ID'])
            notas_ids = set(df_notas['ID'])
            
            colunas_para_atualizar = ['ID', 'Link', 'Tarefa'] + colunas_data
            ids_para_atualizar = list(origem_ids.intersection(notas_ids))
            
            links_para_atualizar = df_origem[df_origem['ID'].isin(ids_para_atualizar)]['Link'].tolist()
            
            if links_para_atualizar:
                df_origem_update = df_origem[df_origem['Link'].isin(links_para_atualizar)].set_index('Link')
                df_notas.set_index('Link', inplace=True)
                
                colunas_update_existentes = [c for c in colunas_para_atualizar if c in df_origem_update.columns and c in df_notas.columns and c != 'Link']
                
                df_antes = df_notas.loc[links_para_atualizar, colunas_update_existentes].copy()
                df_notas.update(df_origem_update[colunas_update_existentes])
                df_depois = df_notas.loc[links_para_atualizar, colunas_update_existentes]
                
                if not df_antes.equals(df_depois):
                    mudancas_detectadas = True
                
                df_notas.reset_index(inplace=True)

            novas_tarefas_ids = origem_ids - notas_ids
            if novas_tarefas_ids:
                mudancas_detectadas = True # Houve mudança
                novas_tarefas_df = df_origem[df_origem['ID'].isin(novas_tarefas_ids)].copy()
                st.toast(f"Sincronizando {len(novas_tarefas_df)} nova(s) tarefa(s)...")
                for col in df_notas.columns:
                    if col not in novas_tarefas_df.columns: novas_tarefas_df[col] = ''
                novas_tarefas_df = novas_tarefas_df[df_notas.columns.tolist()]
                
                for col in colunas_data:
                    if col in novas_tarefas_df.columns:
                        novas_tarefas_df[col] = pd.to_datetime(novas_tarefas_df[col], errors='coerce')

                df_notas = pd.concat([df_notas, novas_tarefas_df], ignore_index=True)

            if mudancas_detectadas:
                st.info("Detectamos atualizações de dados da planilha de origem. Salvando e recarregando...")
                if salvar_df_na_aba(PLANILHA_NOTAS_NOME, df_notas, show_success=False):
                    st.cache_data.clear(); st.rerun()

        if 'Posição' in df_equipes.columns and 'Nome' in df_equipes.columns:
            true_leader_names = df_equipes[df_equipes['Posição'] == 'Lider']['Nome'].tolist()
        else: true_leader_names = []
        
        true_leader_names_lower = {str(name).lower() for name in true_leader_names}
        lideres = [col for col in df_notas.columns if str(col).lower() in true_leader_names_lower]
        
        encarregados = sorted(df_notas['Encarregado'].astype(str).unique())
        
        colunas_de_notas = ['Peso'] + lideres
        for col in colunas_de_notas:
            if col in df_notas.columns:
                df_notas[col] = df_notas[col].astype(str).str.replace(',', '.', regex=False)
                df_notas[col] = df_notas[col].replace(['', 'None', None, 'nan'], pd.NA)
                df_notas[col] = pd.to_numeric(df_notas[col], errors='coerce').fillna(0)
        
        
        # --- INÍCIO: LÓGICA DE ALERTA DE DISCREPÂNCIA E ÓRFÃS ---
        df_alertas = pd.DataFrame()
        df_orfas = pd.DataFrame()

        try:
            if 'ID' not in df_origem.columns or 'ID' not in df_notas.columns:
                st.warning("Não foi possível gerar alertas: Coluna 'ID' faltando em uma das abas.")
                raise ValueError("Coluna 'ID' Faltando")

            origem_ids = set(df_origem['ID'])
            notas_ids = set(df_notas['ID'])

            tarefas_orfas_ids = notas_ids - origem_ids
            if tarefas_orfas_ids:
                df_orfas_raw = df_notas[df_notas['ID'].isin(tarefas_orfas_ids)].copy()
                colunas_orfas = ['ID', 'Link', 'Tarefa', 'Encarregado', 'Data Final', 'Peso']
                colunas_orfas_existentes = [c for c in colunas_orfas if c in df_orfas_raw.columns]
                df_orfas = df_orfas_raw[colunas_orfas_existentes]

            
            if 'Peso' in df_origem.columns:
                df_origem['Peso_num'] = df_origem['Peso'].astype(str).str.replace(',', '.', regex=False)
                df_origem['Peso_num'] = df_origem['Peso_num'].replace(['', 'None', None, 'nan'], pd.NA)
                df_origem['Peso_num'] = pd.to_numeric(df_origem['Peso_num'], errors='coerce').fillna(0)
            else:
                df_origem['Peso_num'] = 0 
            
            ids_comuns = list(origem_ids.intersection(notas_ids))
            if ids_comuns:
                colunas_origem = ['ID', 'Encarregado', 'Peso_num']
                colunas_notas = ['ID', 'Encarregado', 'Peso'] 

                colunas_origem = [c for c in colunas_origem if c in df_origem.columns]
                colunas_notas = [c for c in colunas_notas if c in df_notas.columns]

                df_origem_comp = df_origem[df_origem['ID'].isin(ids_comuns)][colunas_origem]
                df_notas_comp = df_notas[df_notas['ID'].isin(ids_comuns)][colunas_notas]

                df_merged = pd.merge(
                    df_origem_comp, 
                    df_notas_comp, 
                    on='ID', 
                    suffixes=('_origem', '_notas')
                )
                
                if 'Encarregado_origem' in df_merged.columns and 'Encarregado_notas' in df_merged.columns:
                    enc_origem = df_merged['Encarregado_origem'].astype(str).str.strip()
                    enc_notas = df_merged['Encarregado_notas'].astype(str).str.strip()
                    primeiro_nome_origem = enc_origem.str.split(n=1).str[0].fillna('')
                    primeiro_nome_notas = enc_notas.str.split(n=1).str[0].fillna('')
                    
                    nome_origem_normalizado = primeiro_nome_origem.str.normalize('NFD').str.encode('ascii', 'ignore').str.decode('utf-8').str.lower()
                    nome_notas_normalizado = primeiro_nome_notas.str.normalize('NFD').str.encode('ascii', 'ignore').str.decode('utf-8').str.lower()
                    
                    df_merged['enc_diff'] = nome_origem_normalizado != nome_notas_normalizado
                else:
                    df_merged['enc_diff'] = False
                
                if 'Peso_num' in df_merged.columns and 'Peso' in df_merged.columns:
                    df_merged['peso_diff'] = df_merged['Peso_num'] != df_merged['Peso']
                else:
                    df_merged['peso_diff'] = False
                    
                discrepant_ids = df_merged[df_merged['enc_diff'] | df_merged['peso_diff']]['ID'].tolist()

                if discrepant_ids:
                    colunas_alerta = ['ID', 'Link', 'Tarefa', 'Encarregado', 'Peso']
                    
                    colunas_alerta_origem = [c for c in colunas_alerta if c in df_origem.columns]
                    colunas_alerta_notas = [c for c in colunas_alerta if c in df_notas.columns]

                    df_origem_alertas = df_origem[df_origem['ID'].isin(discrepant_ids)][colunas_alerta_origem].copy()
                    df_origem_alertas['Aba'] = PLANILHA_ORIGEM_NOME
                    if 'Peso' in df_origem_alertas.columns:
                         df_origem_alertas['Peso'] = df_origem[df_origem['ID'].isin(discrepant_ids)]['Peso']
                    
                    df_notas_alertas = df_notas[df_notas['ID'].isin(discrepant_ids)][colunas_alerta_notas].copy()
                    df_notas_alertas['Aba'] = PLANILHA_NOTAS_NOME
                    if 'Peso' in df_notas_alertas.columns:
                         df_notas_alertas['Peso'] = df_notas[df_notas['ID'].isin(discrepant_ids)]['Peso']


                    df_alertas = pd.concat([df_origem_alertas, df_notas_alertas], ignore_index=True)
                    df_alertas = df_alertas.sort_values(by=['ID', 'Aba'])
                    
                    colunas_finais_alerta = ['Aba', 'ID', 'Link', 'Tarefa', 'Encarregado', 'Peso']
                    colunas_existentes = [c for c in colunas_finais_alerta if c in df_alertas.columns]
                    df_alertas = df_alertas[colunas_existentes]
        
        except Exception as e:
            st.warning(f"Não foi possível gerar alertas de discrepância/órfãs: {e}")
        # --- FIM: LÓGICA DE ALERTA ---

        return df_notas, df_equipes, encarregados, lideres, df_alertas, df_orfas

    except Exception as e:
        st.error(f"Erro no processo de sincronização e carregamento: {e}")
        return None, None, None, None, pd.DataFrame(), pd.DataFrame()
# ==============================================================================

def adicionar_coluna_e_promover_lider(aba_notas_nome, aba_equipes_nome, nome_lider):
    client = autorizar_cliente()
    url_da_planilha = st.secrets.get("SHEET_URL")
    if not url_da_planilha:
        st.error("URL da planilha não configurada nos segredos para promover líder.")
        return False
        
    spreadsheet = client.open_by_url(url_da_planilha)
    ws_notas = spreadsheet.worksheet(aba_notas_nome)
    
    notas_valores = ws_notas.get_all_values()
    headers_notas = notas_valores[0] if notas_valores else []

    if nome_lider not in headers_notas:
        try:
            ws_notas.add_cols(1)
            nova_coluna_index = len(headers_notas) + 1
            ws_notas.update_cell(1, nova_coluna_index, nome_lider)
        except Exception as e:
            st.error(f"Não foi possível adicionar a coluna: {e}"); return False
    
    try:
        ws_equipes = spreadsheet.worksheet(aba_equipes_nome)
        cell = ws_equipes.find(nome_lider, in_column=1) # Procura na coluna 1 (Nome)
        if cell:
            headers_equipe = ws_equipes.row_values(1)
            if "Posição" in headers_equipe:
                posicao_col_index = headers_equipe.index("Posição") + 1
                ws_equipes.update_cell(cell.row, posicao_col_index, "Lider")
                st.toast(f"👑 {nome_lider} promovido a Líder com sucesso!"); return True
            else:
                st.error("Coluna 'Posição' não encontrada na aba 'Equipes'.")
                return False
        else:
            st.warning(f"'{nome_lider}' não encontrado na aba 'Equipes'. Coluna adicionada em 'Notas', mas 'Posição' não atualizada.")
            return True 
            
    except Exception as e:
        st.error(f"Não foi possível atualizar a posição do líder: {e}"); return False

def corrigir_nome_encarregado(aba_nome, df_notas, nome_incorreto, nome_correto):
    try:
        df_corrigido = df_notas.copy()
        df_corrigido['Encarregado'] = df_corrigido['Encarregado'].replace(nome_incorreto, nome_correto)
        if salvar_df_na_aba(aba_nome, df_corrigido):
            st.toast(f"Nome '{nome_incorreto}' corrigido para '{nome_correto}' com sucesso!"); return True
        return False
    except Exception as e:
        st.error(f"Erro ao corrigir o nome: {e}"); return False

def gerar_arquivo_excel(df_geral, df_completo, lideres, df_equipes):
    """Gera um arquivo Excel em memória com 3 abas de relatório."""
    output = io.BytesIO()
    
    if df_equipes is not None and not df_equipes.empty and 'Status' in df_equipes.columns and 'Nome' in df_equipes.columns:
        nomes_ativos = df_equipes[df_equipes['Status'] == 'Ativo']['Nome'].tolist()
        nomes_ativos_lower = {str(nome).lower() for nome in nomes_ativos}
    else:
        nomes_ativos = df_completo['Encarregado'].unique().tolist()
        nomes_ativos_lower = {str(nome).lower() for nome in nomes_ativos}

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        
        # --- Aba 1: Relatório Geral (Já formatado) ---
        df_completo_excel = df_geral.copy() 
        for col in df_completo_excel.select_dtypes(include=['datetime', 'datetimetz']).columns:
            df_completo_excel[col] = df_completo_excel[col].dt.strftime('%d/%m/%Y')
        df_completo_excel.to_excel(writer, sheet_name='Relatório Geral', index=False)

        # --- Aba 2: Soma Geral ---
        df_calculo = df_completo.copy() 

        soma_peso_geral = df_calculo.groupby('Encarregado')['Peso'].sum().reset_index()
        soma_peso_geral_filtrado = soma_peso_geral[soma_peso_geral['Encarregado'].str.lower().isin(nomes_ativos_lower)]
        soma_peso_geral_filtrado.to_excel(writer, sheet_name='Soma Geral', index=False, startrow=0)
        
        df_pontos_lideranca = []
        lideres_lower = {l.lower() for l in lideres}
        colunas_df_lower = {c.lower(): c for c in df_calculo.columns}

        for lider_lower in lideres_lower:
            if lider_lower in colunas_df_lower:
                lider_col_original = colunas_df_lower[lider_lower]
                pontos = df_calculo[lider_col_original].fillna(0).sum()
                df_pontos_lideranca.append({'Líder': lider_col_original, 'Pontos por Liderança': pontos})

        if df_pontos_lideranca:
            df_pontos_lideranca_df = pd.DataFrame(df_pontos_lideranca)
            df_pontos_lideranca_filtrado = df_pontos_lideranca_df[df_pontos_lideranca_df['Líder'].str.lower().isin(nomes_ativos_lower)]
            start_row_pontos = len(soma_peso_geral_filtrado) + 4
            df_pontos_lideranca_filtrado.to_excel(writer, sheet_name='Soma Geral', index=False, startrow=start_row_pontos)
            
            pontuacao_total = soma_peso_geral.copy()
            df_pontos_lideranca_df.rename(columns={'Líder': 'Encarregado'}, inplace=True)
            pontuacao_total['Encarregado_lower'] = pontuacao_total['Encarregado'].str.lower()
            df_pontos_lideranca_df['Encarregado_lower'] = df_pontos_lideranca_df['Encarregado'].str.lower()
            
            pontuacao_total = pd.merge(pontuacao_total, df_pontos_lideranca_df[['Encarregado_lower', 'Pontos por Liderança']], on='Encarregado_lower', how='left').fillna(0)
            pontuacao_total['Pontuação Total'] = pontuacao_total['Peso'] + pontuacao_total['Pontos por Liderança']
            pontuacao_total.drop(columns=['Encarregado_lower', 'Pontos por Liderança'], inplace=True)
            
            pontuacao_total_filtrado = pontuacao_total[pontuacao_total['Encarregado'].str.lower().isin(nomes_ativos_lower)]
            start_row_total = start_row_pontos + len(df_pontos_lideranca_filtrado) + 4
            pontuacao_total_filtrado.to_excel(writer, sheet_name='Soma Geral', index=False, startrow=start_row_total)

        # --- Aba 3: Soma Semanal ---
        df_semanal = df_calculo.copy() 
        df_semanal = df_semanal[pd.notna(df_semanal['Data Final'])]
        
        if not df_semanal.empty:
            df_semanal['Semana_dt'] = df_semanal['Data Final'].apply(lambda x: x + pd.to_timedelta((4 - x.dayofweek + 7) % 7, unit='d'))
            df_semanal['Semana'] = df_semanal['Semana_dt'].dt.strftime('%d/%m/%Y')
            
            sorted_weeks = sorted(df_semanal['Semana'].unique(), key=lambda d: pd.to_datetime(d, format='%d/%m/%Y'))

            pivot_peso = df_semanal.pivot_table(index='Encarregado', columns='Semana', values='Peso', aggfunc='sum').fillna(0)
            pivot_peso_filtrado = pivot_peso[pivot_peso.index.str.lower().isin(nomes_ativos_lower)]
            pivot_peso_filtrado = pivot_peso_filtrado.reindex(columns=sorted_weeks, fill_value=0)
            pivot_peso_filtrado.to_excel(writer, sheet_name='Soma Semanal', startrow=0)
            
            df_semanal['Pontos de Liderança'] = df_semanal[lideres].sum(axis=1)
            pivot_lideranca = df_semanal.pivot_table(index='Encarregado', columns='Semana', values='Pontos de Liderança', aggfunc='sum').fillna(0)
            pivot_lideranca_filtrado = pivot_lideranca[pivot_lideranca.index.str.lower().isin(nomes_ativos_lower)]
            pivot_lideranca_filtrado = pivot_lideranca_filtrado.reindex(columns=sorted_weeks, fill_value=0)
            
            start_row_lideres_semanal = len(pivot_peso_filtrado) + 4
            worksheet = writer.sheets['Soma Semanal']
            worksheet.cell(row=start_row_lideres_semanal-1, column=1, value="Pontos de Liderança (Semanal)")
            pivot_lideranca_filtrado.to_excel(writer, sheet_name='Soma Semanal', startrow=start_row_lideres_semanal)
            
    processed_data = output.getvalue()
    return processed_data

# ==============================================================================
# --- LÓGICA PRINCIPAL DO APP COM LOGIN ---
# ==============================================================================

# Inicializa o estado de autenticação
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "user_role" not in st.session_state:
    st.session_state.user_role = None

# --- TELA DE LOGIN ---
if not st.session_state.authenticated:
    st.title("🔒 Login - Gerenciador de Notas")
    st.write("Por favor, insira suas credenciais para acessar o aplicativo.")
    
    with st.form(key="login_form"):
        username = st.text_input("Usuário")
        password = st.text_input("Senha", type="password")
        login_button = st.form_submit_button(label="Entrar")
        
        if login_button:
            with st.spinner("Verificando..."):
                role = check_credentials(username, password)
                if role:
                    st.session_state.authenticated = True
                    st.session_state.user_role = role # Armazena a função do usuário
                    st.rerun()
                else:
                    st.error("Usuário ou senha inválidos.")

# --- APLICAÇÃO PRINCIPAL (SE LOGADO) ---
else:
    if "id_para_buscar" not in st.session_state:
        st.session_state.id_para_buscar = ""
    
    col_titulo, col_botao = st.columns([3, 1])
    with col_titulo:
        st.title("📝 Gerenciador de Notas de Tarefas")
        st.write(f"Sincronize tarefas, edite notas, promova líderes e exporte relatórios. (Modo: **{st.session_state.user_role}**)")

    df_notas, df_equipes, encarregados, lideres, df_alertas, df_orfas = sincronizar_e_processar_dados()

    if df_alertas is not None and not df_alertas.empty:
        num_discrepancias = len(df_alertas['ID'].unique())
        st.warning(f"🚨 **Alerta de Sincronia!** {num_discrepancias} tarefa(s) têm valores de 'Encarregado' ou 'Peso' diferentes da planilha de origem.")
        
        with st.expander("Clique aqui para ver e corrigir as discrepâncias"):
            st.info("""
            Abaixo estão as tarefas com valores conflitantes.
            - **Linhas Vermelhas (Origem):** Mostram o valor atual na planilha 'Total BaseCamp'.
            - **Linhas Brancas (Notas):** Mostram o valor na sua planilha 'Total BaseCamp para Notas'.
            
            **Para corrigir:**
            - Use a ferramenta "Corrigir Nomes de Encarregados" na barra lateral.
            - Ou edite o 'Peso' ou 'Encarregado' na tabela "Editor de Notas" abaixo e salve.
            """)
            
            def highlight_rows(row):
                if row['Aba'] == PLANILHA_ORIGEM_NOME:
                    return ['background-color: #ffcccc; color: black;'] * len(row)
                else:
                    return ['background-color: white; color: black;'] * len(row)
            
            df_alertas_display = df_alertas.copy()
            if 'Peso' in df_alertas_display.columns:
                 df_alertas_display['Peso'] = df_alertas_display['Peso'].astype(str)

            st.dataframe(
                df_alertas_display.style.apply(highlight_rows, axis=1),
                use_container_width=True,
                column_config={
                    "Link": st.column_config.LinkColumn("Link", display_text="Abrir ↗"),
                    "Peso": st.column_config.TextColumn("Peso"), 
                },
                hide_index=True
            )

    if df_orfas is not None and not df_orfas.empty:
        num_orfas = len(df_orfas)
        st.error(f"🔍 **Tarefas Órfãs Encontradas!** {num_orfas} tarefa(s) estão na sua planilha 'Notas', mas não existem mais na 'Origem' (Total BaseCamp).")
        
        with st.expander("Clique aqui para ver e deletar tarefas órfãs"):
            st.info("""
            Estas tarefas provavelmente foram deletadas ou movidas no BaseCamp.
            Você pode usar a ferramenta "Deletar Tarefa" na barra lateral (copiando o ID abaixo) para removê-las permanentemente da sua planilha de Notas.
            """)
            
            st.dataframe(
                df_orfas,
                use_container_width=True,
                column_config={
                    "Link": st.column_config.LinkColumn("Link", display_text="Abrir ↗"),
                    "Data Final": st.column_config.DateColumn("Data Final", format="DD/MM/YYYY"),
                    "Peso": st.column_config.NumberColumn("Peso", format="%0.2f"),
                },
                hide_index=True
            )


    with st.sidebar:
        st.image("media portal logo.png", width=200)
        st.sidebar.button("Sair / Logout", on_click=lambda: st.session_state.update(authenticated=False, user_role=None))

        if st.session_state.user_role == "Editor":
            st.header("Sincronização Manual")
            if st.button("🔄 Atualizar Dados do BaseCamp"):
                with st.spinner("Buscando novas tarefas do BaseCamp..."):
                    if atualizar_basecamp_api():
                        st.success("Planilha 'Total BaseCamp' atualizada com sucesso!")
                        st.info("Limpando cache e recarregando o aplicativo para sincronizar as notas...")
                        st.cache_data.clear()
                        st.rerun()
                    else:
                        st.error("Falha ao buscar dados do BaseCamp.")
            
            if st.button("📈 Salvar Snapshot Diário (para Dashboard)"):
                with st.spinner("Calculando totais e salvando no histórico..."):
                    if atualizar_historico_diario():
                        st.cache_data.clear() 
                    else:
                        st.error("Falha ao salvar o snapshot diário.")
        
        st.header("Filtros e Ordenação")
        
        status_filtro = st.selectbox("Filtrar por Status do Encarregado:", ["Todos", "Ativos", "Desativados", "Não Listados"])

        nomes_para_filtrar = []
        if df_equipes is not None and df_notas is not None:
            if status_filtro == 'Todos': nomes_para_filtrar = encarregados
            elif status_filtro == 'Ativos':
                nomes_ativos = df_equipes[df_equipes['Status'] == 'Ativo']['Nome'].tolist()
                nomes_para_filtrar = [nome for nome in encarregados if nome in nomes_ativos]
            elif status_filtro == 'Desativados':
                nomes_desativados = df_equipes[df_equipes['Status'] == 'Desativado']['Nome'].tolist()
                nomes_para_filtrar = [nome for nome in encarregados if nome in nomes_desativados]
            elif status_filtro == 'Não Listados':
                nomes_listados = df_equipes['Nome'].tolist()
                nomes_para_filtrar = [nome for nome in encarregados if nome not in nomes_listados]
        
        filtro_encarregado = st.multiselect("Filtrar por Encarregado:", ["Todos"] + sorted(nomes_para_filtrar), default="Todos")
        
        st.write("Ordenar por:")
        opcoes_ordem = ['Encarregado', 'Data Inicial', 'Data Final', 'Tarefa']
        ordem_primaria = st.selectbox("Fator Primário:", opcoes_ordem, index=0)
        ordem_secundaria = st.selectbox("Fator Secundário:", opcoes_ordem, index=1)
        
        if st.session_state.user_role == "Editor":
            st.header("Gerenciar Líderes")
            if df_notas is not None and df_equipes is not None:
                encarregados_nao_lideres = sorted(list(set(encarregados) - set(lideres)))
                if encarregados_nao_lideres:
                    novo_lider = st.selectbox("Promover encarregado a líder:", encarregados_nao_lideres)
                    if st.button(f"Promover {novo_lider}"):
                        if adicionar_coluna_e_promover_lider(PLANILHA_NOTAS_NOME, PLANILHA_EQUIPES_NOME, novo_lider):
                            st.cache_data.clear(); st.success(f"{novo_lider} promovido! A página será recarregada."); st.rerun()
                else: st.info("Todos os encarregados já são líderes.")

            st.header("Corrigir Nomes de Encarregados")
            if df_notas is not None and df_equipes is not None:
                nomes_unicos_notas = sorted(df_notas['Encarregado'].astype(str).unique())
                nomes_corretos_equipe = sorted(df_equipes['Nome'].astype(str).unique())
                nome_a_corrigir = st.selectbox("Selecione o nome com erro:", options=nomes_unicos_notas)
                nome_correto = st.selectbox("Selecione o nome correto:", options=nomes_corretos_equipe)
                if st.button(f"Corrigir '{nome_a_corrigir}' para '{nome_correto}'"):
                    if nome_a_corrigir and nome_correto and nome_a_corrigir != nome_correto:
                        if corrigir_nome_encarregado(PLANILHA_NOTAS_NOME, df_notas, nome_a_corrigir, nome_correto):
                            st.cache_data.clear(); st.success("Nome corrigido! A página será recarregada."); st.rerun()
                    else: st.warning("Selecione nomes diferentes para a correção.")

            st.header("Deletar Tarefa")
            st.warning("Cuidado: Deletar uma tarefa é permanente.")
            
            id_para_deletar = st.text_input("Digite ou cole o ID da tarefa a deletar:")
            
            if id_para_deletar and df_notas is not None:
                id_para_deletar = id_para_deletar.strip()
                
                tarefa_existe_df = df_notas[df_notas['ID'] == id_para_deletar]
                
                if not tarefa_existe_df.empty:
                    tarefa_nome = tarefa_existe_df.iloc[0].get('Tarefa', 'Nome não encontrado')
                    st.write(f"**Tarefa a deletar:** {tarefa_nome[:50]}...")

                    if st.button(f"Confirmar Deleção do ID {id_para_deletar}", type="primary"):
                        with st.spinner(f"Deletando tarefa {id_para_deletar}..."):
                            
                            df_apos_delecao = df_notas[df_notas['ID'] != id_para_deletar].copy()
                            
                            if salvar_df_na_aba(PLANILHA_NOTAS_NOME, df_apos_delecao):
                                st.cache_data.clear()
                                st.success("Tarefa deletada! A página será recarregada.")
                                st.rerun()
                            else:
                                st.error("Erro ao salvar a deleção na planilha.")
                elif id_para_deletar: 
                    st.error(f"ID '{id_para_deletar}' não encontrado na planilha de Notas.")
            
            st.markdown("---") 

    if df_notas is not None and df_equipes is not None: 

        st.markdown("---")
        st.subheader("Buscar Tarefa por ID")
        st.info("Use esta ferramenta para encontrar uma tarefa específica na tabela abaixo para facilitar a edição.")
        
        search_col1, search_col2, search_col3 = st.columns([3, 1, 1])
        with search_col1:
            id_search_input = st.text_input("Digite o ID da tarefa para buscar:", 
                                            value=st.session_state.id_para_buscar,
                                            key="search_id_input")
        with search_col2:
            st.write("") # for vertical alignment
            if st.button("Buscar Tarefa", key="search_button"):
                st.session_state.id_para_buscar = id_search_input.strip()
                st.rerun()
        with search_col3:
            st.write("") # for vertical alignment
            if st.button("Limpar Busca", key="clear_search_button"):
                st.session_state.id_para_buscar = ""
                st.rerun()

        df_para_exibir = df_notas.copy()

        if "Todos" not in filtro_encarregado:
            df_para_exibir = df_para_exibir[df_para_exibir['Encarregado'].isin(filtro_encarregado)]
        
        if ordem_primaria and ordem_secundaria and ordem_primaria != ordem_secundaria:
            try:
                df_para_exibir = df_para_exibir.sort_values(by=[ordem_primaria, ordem_secundaria], na_position='last')
            except TypeError as e:
                st.warning(f"Não foi possível ordenar por {ordem_primaria} ou {ordem_secundaria}. Pode haver tipos de dados mistos. {e}")
                df_para_exibir = df_para_exibir.astype(str).sort_values(by=[ordem_primaria, ordem_secundaria])
        
        if st.session_state.id_para_buscar:
            df_busca_resultado = df_notas[df_notas['ID'] == st.session_state.id_para_buscar]
            
            if df_busca_resultado.empty:
                st.error(f"ID '{st.session_state.id_para_buscar}' não encontrado na planilha 'Notas'.")
                st.session_state.id_para_buscar = "" # Clear bad search
            else:
                st.success(f"Exibindo apenas o ID '{st.session_state.id_para_buscar}'. Clique em 'Limpar Busca' para ver a lista filtrada.")
                df_para_exibir = df_busca_resultado # Substitui a exibição
        

        with col_botao:
            st.write("") 
            st.write("")
            excel_data = gerar_arquivo_excel(df_para_exibir, df_notas, lideres, df_equipes)
            st.download_button(
                label="🖨️ Imprimir Relatório",
                data=excel_data,
                file_name=f"relatorio_tarefas_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        st.markdown("---") # O <hr> original
        st.header("Editor de Notas")
        
        is_read_only = st.session_state.user_role == "Visualizador"
        
        if is_read_only:
            st.info("Você está no modo 'Visualizador'. A edição está desabilitada.")
            colunas_desabilitadas = df_para_exibir.columns.tolist() # Desabilita todas
        else:
            if not st.session_state.id_para_buscar:
                st.info("Você está no modo 'Editor'. Clique nas células para editar 'Encarregado', 'Peso' e as colunas de líderes.")
            else:
                st.info("Você está no modo 'Editor'. Edite a tarefa abaixo e clique em 'Salvar Alterações'. A busca será limpa após salvar.")
            
            colunas_editaveis = ['Encarregado', 'Peso'] + lideres
            colunas_desabilitadas = [col for col in df_para_exibir.columns if col not in colunas_editaveis]
        
        nomes_corretos_equipe = sorted(df_equipes['Nome'].astype(str).unique())
        
        df_editado = st.data_editor(
            df_para_exibir, 
            disabled=colunas_desabilitadas,
            key="editor_notas",
            column_config={
                "Link": st.column_config.LinkColumn(
                    "Link da Tarefa",
                    display_text="Abrir ↗"
                ),
                "Encarregado": st.column_config.SelectboxColumn(
                    "Encarregado",
                    options=nomes_corretos_equipe,
                ),
                "Data Inicial": st.column_config.DateColumn("Data Inicial", format="DD/MM/YYYY"),
                "Data Final": st.column_config.DateColumn("Data Final", format="DD/MM/YYYY"),
                "Data Estipulada": st.column_config.DateColumn("Data Estipulada", format="DD/MM/YYYY"),
                "Peso": st.column_config.NumberColumn("Peso", format="%0.2f"),
                # ==============================================================================
                # --- CORREÇÃO: st_column_config -> st.column_config ---
                # ==============================================================================
                **{lider: st.column_config.NumberColumn(lider, format="%0.2f") for lider in lideres}
            }
        )
        
        st.markdown("---")
        
        if not is_read_only:
            if st.button("Salvar Alterações na Planilha Google", type="primary"):
                with st.spinner("Salvando..."):
                    df_atualizado = df_notas.copy()
                    
                    df_atualizado['Link'] = df_atualizado['Link'].astype(str)
                    df_editado['Link'] = df_editado['Link'].astype(str)
                    
                    df_atualizado.set_index('Link', inplace=True)
                    df_editado_idx = df_editado.set_index('Link')
                    
                    df_atualizado.update(df_editado_idx[colunas_editaveis])
                    
                    colunas_numericas_para_limpar = ['Peso'] + lideres 
                    for col in colunas_numericas_para_limpar:
                        if col in df_atualizado.columns:
                            df_atualizado[col] = df_atualizado[col].astype(str).str.replace(',', '.', regex=False)
                            df_atualizado[col] = pd.to_numeric(df_atualizado[col], errors='coerce').fillna(0)
                    
                    df_atualizado.reset_index(inplace=True)
                    df_atualizado = df_atualizado[df_atualizado['ID'] != '']
                    
                    if salvar_df_na_aba(PLANILHA_NOTAS_NOME, df_atualizado):
                        st.cache_data.clear()
                        st.session_state.id_para_buscar = "" 
                        st.success("Alterações salvas! A página será recarregada."); 
                        st.rerun()
                    else: 
                        st.error("Houve um erro ao salvar.")
    else:
        st.error("Não foi possível carregar os dados das planilhas.")
        # --- CORREÇÃO: Removida a aspa extra no final ---
        st.info("Verifique as permissões, nomes das abas e se a URL da planilha está correta nos segredos.")