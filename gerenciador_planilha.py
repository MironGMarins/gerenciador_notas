# ==============================================================================
# IMPORTS
# ==============================================================================
import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import io
import requests # Adicionado para chamadas de API

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
    """Carrega uma aba mesmo que ela tenha cabeçalhos duplicados."""
    all_values = worksheet.get_all_values()
    if not all_values:
        return pd.DataFrame()
    
    headers = all_values[0]
    data = all_values[1:]
    
    return pd.DataFrame(data, columns=headers)

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

    spreadsheet = client.open_by_url(url_da_planilha)
    ws_origem = spreadsheet.worksheet(PLANILHA_ORIGEM_NOME)
    ws_notas = spreadsheet.worksheet(PLANILHA_NOTAS_NOME)
    ws_equipes = spreadsheet.worksheet(PLANILHA_EQUIPES_NOME)

    df_origem = carregar_aba_de_forma_robusta(ws_origem)
    df_notas = carregar_aba_de_forma_robusta(ws_notas)
    df_equipes = carregar_aba_de_forma_robusta(ws_equipes)
    
    return df_origem, df_notas, df_equipes

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
        df_list = [df_filled.columns.values.tolist()] + df_filled.astype(str).values.tolist()
        worksheet.clear()
        worksheet.update(df_list, value_input_option='USER_ENTERED')
        
        if show_success:
            st.toast("✅ Alterações salvas com sucesso na Planilha Google!")
        return True
    except Exception as e:
        st.error(f"Erro ao salvar na planilha: {e}")
        return False

def sincronizar_e_processar_dados():
    """Usa dados em cache para processar, sincronizar e retornar os DataFrames finais."""
    try:
        df_origem, df_notas, df_equipes = _carregar_dados_brutos()
        if df_notas is None: return None, None, None, None

        if 'Encarregado' in df_origem.columns: df_origem['Encarregado'] = df_origem['Encarregado'].astype(str)
        if 'Encarregado' in df_notas.columns: df_notas['Encarregado'] = df_notas['Encarregado'].astype(str).str.strip()
        if 'Nome' in df_equipes.columns: df_equipes['Nome'] = df_equipes['Nome'].astype(str).str.strip()

        for col in ['Data Inicial', 'Data Final', 'Data Estipulada']:
            if col in df_notas.columns: df_notas[col] = pd.to_datetime(df_notas[col], errors='coerce')

        if 'Link' in df_origem.columns and 'Link' in df_notas.columns:
            df_origem['ID'] = df_origem['Link'].str.split('/').str[-1].fillna('').astype(str)
            df_notas['ID'] = df_notas['Link'].str.split('/').str[-1].fillna('').astype(str)
        
        df_notas.dropna(subset=['ID'], inplace=True); df_notas = df_notas[df_notas['ID'] != '']
        if df_notas['ID'].duplicated().any(): df_notas.drop_duplicates(subset=['ID'], keep='first', inplace=True, ignore_index=True)

        if 'ID' in df_origem.columns and 'ID' in df_notas.columns:
            origem_ids = set(df_origem[df_origem['ID'] != '']['ID'])
            notas_ids = set(df_notas[df_notas['ID'] != '']['ID'])
            novas_tarefas_ids = origem_ids - notas_ids
            
            if novas_tarefas_ids:
                novas_tarefas_df = df_origem[df_origem['ID'].isin(novas_tarefas_ids)].copy()
                st.toast(f"Sincronizando {len(novas_tarefas_df)} nova(s) tarefa(s)...")
                for col in df_notas.columns:
                    if col not in novas_tarefas_df.columns: novas_tarefas_df[col] = ''
                novas_tarefas_df = novas_tarefas_df[df_notas.columns.tolist()]
                df_notas_atualizado = pd.concat([df_notas, novas_tarefas_df], ignore_index=True)
                if salvar_df_na_aba(PLANILHA_NOTAS_NOME, df_notas_atualizado, show_success=False):
                    st.cache_data.clear(); st.rerun()

        if 'Posição' in df_equipes.columns and 'Nome' in df_equipes.columns:
            true_leader_names = df_equipes[df_equipes['Posição'] == 'Lider']['Nome'].tolist()
        else: true_leader_names = []
        
        true_leader_names_lower = {str(name).lower() for name in true_leader_names}
        lideres = [col for col in df_notas.columns if str(col).lower() in true_leader_names_lower]
        
        encarregados = sorted(df_notas['Encarregado'].astype(str).unique())
        
        colunas_de_notas = ['Peso'] + lideres
        for col in colunas_de_notas:
            if col in df_notas.columns: df_notas[col] = pd.to_numeric(df_notas[col], errors='coerce')
        
        return df_notas, df_equipes, encarregados, lideres

    except Exception as e:
        st.error(f"Erro no processo de sincronização e carregamento: {e}")
        return None, None, None, None

def adicionar_coluna_e_promover_lider(aba_notas_nome, aba_equipes_nome, nome_lider):
    client = autorizar_cliente()
    url_da_planilha = st.secrets.get("SHEET_URL")
    if not url_da_planilha:
        st.error("URL da planilha não configurada nos segredos para promover líder.")
        return False
        
    spreadsheet = client.open_by_url(url_da_planilha)
    ws_notas = spreadsheet.worksheet(aba_notas_nome)
    df_notas = pd.DataFrame(ws_notas.get_all_records())

    if nome_lider not in df_notas.columns:
        try:
            ws_notas.add_cols(1)
            nova_coluna_index = len(df_notas.columns) + 1
            ws_notas.update_cell(1, nova_coluna_index, nome_lider)
        except Exception as e:
            st.error(f"Não foi possível adicionar a coluna: {e}"); return False
    
    try:
        ws_equipes = spreadsheet.worksheet(aba_equipes_nome)
        cell = ws_equipes.find(nome_lider, in_column=1)
        if cell:
            headers = ws_equipes.row_values(1)
            if "Posição" in headers:
                posicao_col_index = headers.index("Posição") + 1
                ws_equipes.update_cell(cell.row, posicao_col_index, "Lider")
                st.toast(f"👑 {nome_lider} promovido a Líder com sucesso!"); return True
        return False
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
    
    if df_equipes is not None and not df_equipes.empty:
        nomes_ativos = df_equipes[df_equipes['Status'] == 'Ativo']['Nome'].tolist()
        nomes_ativos_lower = {str(nome).lower() for nome in nomes_ativos}
    else:
        nomes_ativos = df_completo['Encarregado'].unique().tolist()
        nomes_ativos_lower = {str(nome).lower() for nome in nomes_ativos}

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_completo.to_excel(writer, sheet_name='Relatório Geral', index=False)

        df_calculo = df_completo.copy()
        colunas_de_notas = ['Peso'] + lideres
        for col in colunas_de_notas:
            if col in df_calculo.columns:
                df_calculo[col] = pd.to_numeric(df_calculo[col], errors='coerce')

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

        df_semanal = df_calculo.copy()
        df_semanal = df_semanal[pd.notna(df_semanal['Data Final'])]
        
        if not df_semanal.empty:
            start_date = pd.to_datetime("2025-07-04")
            df_semanal['Semana_dt'] = df_semanal['Data Final'].apply(lambda x: start_date + pd.to_timedelta(((x - start_date).days // 7) * 7, unit='d'))
            df_semanal['Semana'] = df_semanal['Semana_dt'].dt.strftime('%d/%m/%Y')
            sorted_weeks = sorted(df_semanal['Semana'].unique(), key=lambda d: pd.to_datetime(d, format='%d/%m/%Y'))

            pivot_peso = df_semanal.pivot_table(index='Encarregado', columns='Semana', values='Peso', aggfunc='sum').fillna(0)
            pivot_peso_filtrado = pivot_peso[pivot_peso.index.str.lower().isin(nomes_ativos_lower)]
            pivot_peso_filtrado = pivot_peso_filtrado.reindex(columns=sorted_weeks, fill_value=0)
            pivot_peso_filtrado.to_excel(writer, sheet_name='Soma Semanal', startrow=0)
            
            for lider_col in lideres:
                if lider_col in df_semanal.columns:
                    df_semanal[lider_col] = pd.to_numeric(df_semanal[lider_col], errors='coerce').fillna(0)
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
    col_titulo, col_botao = st.columns([3, 1])
    with col_titulo:
        st.title("📝 Gerenciador de Notas de Tarefas")
        st.write(f"Sincronize tarefas, edite notas, promova líderes e exporte relatórios. (Modo: **{st.session_state.user_role}**)")

    df_notas, df_equipes, encarregados, lideres = sincronizar_e_processar_dados()

    with st.sidebar:
        st.image("media portal logo.png", width=200)
        st.sidebar.button("Sair / Logout", on_click=lambda: st.session_state.update(authenticated=False, user_role=None))

        # --- SEÇÕES CONDICIONAIS PARA EDITORES ---
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
        
        # --- SEÇÕES CONDICIONAIS PARA EDITORES ---
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

    # O resto da lógica principal só é executado se os dados foram carregados
    if df_notas is not None:
        df_para_exibir = df_notas.copy()

        if "Todos" not in filtro_encarregado:
            df_para_exibir = df_para_exibir[df_para_exibir['Encarregado'].isin(filtro_encarregado)]
            
        if ordem_primaria and ordem_secundaria and ordem_primaria != ordem_secundaria:
            df_para_exibir = df_para_exibir.sort_values(by=[ordem_primaria, ordem_secundaria])

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
        
        st.markdown("---")
        st.header("Editor de Notas")

        # --- LÓGICA CONDICIONAL DE EDIÇÃO ---
        is_read_only = st.session_state.user_role == "Visualizador"
        
        if is_read_only:
            st.info("Você está no modo 'Visualizador'. A edição está desabilitada.")
            colunas_desabilitadas = df_para_exibir.columns.tolist() # Desabilita todas
        else:
            st.info("Você está no modo 'Editor'. Clique nas células para editar 'Peso' e as colunas de líderes.")
            colunas_editaveis = ['Peso'] + lideres
            colunas_desabilitadas = [col for col in df_para_exibir.columns if col not in colunas_editaveis]
        
        df_editado = st.data_editor(
            df_para_exibir, 
            disabled=colunas_desabilitadas,
            key="editor_notas",
            column_config={
                "Link": st.column_config.LinkColumn(
                    "Link da Tarefa",
                    display_text="Abrir ↗"
                )
            }
        )
        
        st.markdown("---")
        
        # --- BOTÃO DE SALVAR CONDICIONAL ---
        if not is_read_only:
            if st.button("Salvar Alterações na Planilha Google", type="primary"):
                with st.spinner("Salvando..."):
                    df_atualizado = df_notas.copy()
                    df_atualizado['ID'] = df_atualizado['ID'].astype(str)
                    df_editado['ID'] = df_editado['ID'].astype(str)
                    df_atualizado.set_index('ID', inplace=True)
                    df_editado.set_index('ID', inplace=True)
                    df_atualizado.update(df_editado)
                    df_atualizado.reset_index(inplace=True)
                    df_atualizado = df_atualizado[df_atualizado['ID'] != '']
                    if salvar_df_na_aba(PLANILHA_NOTAS_NOME, df_atualizado):
                        st.cache_data.clear()
                        st.success("Alterações salvas! A página será recarregada."); 
                        st.rerun()
                    else: 
                        st.error("Houve um erro ao salvar.")