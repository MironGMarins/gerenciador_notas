import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# --- CONFIGURAÇÃO ---
ABA_MES = "Julho 2025"
ABA_CONSOLIDADA = "Total BaseCamp para Notas"

# --- CONEXÃO ---
print("🕵️ Conectando ao Google Sheets...")
scopes = ["https://www.googleapis.com/auth/spreadsheets"]
creds = Credentials.from_service_account_file("google_credentials.json", scopes=scopes) # Ou use st.secrets se rodar no streamlit
client = gspread.authorize(creds)
spreadsheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1juyOfIh0ZqsfJjN0p3gD8pKaAIX0R6IAPG9vysl7yWI/edit?gid=1943416862#gid=1943416862") # <--- COLOQUE SUA URL AQUI

# --- CARREGAR DADOS ---
print(f"📂 Lendo aba original: '{ABA_MES}'...")
ws_mes = spreadsheet.worksheet(ABA_MES)
# Pega tudo, sem filtros
dados_mes = ws_mes.get_all_records()
df_mes = pd.DataFrame(dados_mes)

print(f"📂 Lendo aba consolidada: '{ABA_CONSOLIDADA}'...")
ws_cons = spreadsheet.worksheet(ABA_CONSOLIDADA)
dados_cons = ws_cons.get_all_records()
df_cons = pd.DataFrame(dados_cons)

# --- ANÁLISE ---
print("\n--- RESULTADOS DA CONTAGEM ---")
qtd_mes = len(df_mes)
print(f"Original ({ABA_MES}): {qtd_mes} linhas encontradas pelo Python.")

# Filtra na consolidada apenas o que veio de Julho
df_cons_julho = df_cons[df_cons['Fonte_Dados'].astype(str).str.contains(ABA_MES, na=False)]
qtd_cons = len(df_cons_julho)
print(f"Consolidada (Vieram de Julho): {qtd_cons} linhas.")

diferenca = qtd_cons - qtd_mes

if diferenca == 0:
    print("\n✅ Os números batem! O mistério pode ser apenas visualização.")
else:
    print(f"\n⚠️ Diferença de {diferenca} linhas!")

# --- QUEM SÃO AS LINHAS EXTRAS? ---
# Vamos verificar se há duplicatas DENTRO da aba original
if 'ID' in df_mes.columns:
    duplicatas_na_origem = df_mes[df_mes.duplicated(subset=['ID'], keep=False)]
    if not duplicatas_na_origem.empty:
        print(f"\n🚨 ACHEI! Existem {len(duplicatas_na_origem)} linhas com IDs duplicados DENTRO de '{ABA_MES}'.")
        print("Isso explica por que o consolidado tem mais linhas (ele manteve as duplicatas).")
        print("\nExemplos de IDs duplicados na aba original:")
        print(duplicatas_na_origem[['ID', 'Encarregado', 'Tarefa']].head(10).to_string())
    else:
        print("\nℹ️ Não há IDs duplicados na aba original.")
else:
    # Se não tiver ID, tenta pelo Link
    if 'Link' in df_mes.columns:
        duplicatas_na_origem = df_mes[df_mes.duplicated(subset=['Link'], keep=False)]
        if not duplicatas_na_origem.empty:
            print(f"\n🚨 ACHEI! Existem {len(duplicatas_na_origem)} Links duplicados DENTRO de '{ABA_MES}'.")
            print(duplicatas_na_origem[['Link', 'Encarregado']].head(10))