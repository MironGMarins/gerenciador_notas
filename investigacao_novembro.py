import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# --- CONFIGURAÇÃO ---
ABA_ALVO = "Novembro 2025"
MES_ESPERADO = 11
ANO_ESPERADO = 2025

# --- FUNÇÃO DE DATA CORRIGIDA ---
def converter_data_hibrida(series):
    # Prioridade ISO (YYYY-MM-DD) para não confundir mês/dia
    datas_prioridade = pd.to_datetime(series, format='%Y-%m-%d', errors='coerce')
    falhas = datas_prioridade.isna()
    if falhas.any():
        # Fallback para BR (DD/MM/YYYY)
        datas_br = pd.to_datetime(series[falhas], dayfirst=True, errors='coerce')
        datas_prioridade = datas_prioridade.fillna(datas_br)
    return datas_prioridade

# --- CONEXÃO ---
print("🕵️ Conectando...")
scopes = ["https://www.googleapis.com/auth/spreadsheets"]
creds = Credentials.from_service_account_file("google_credentials.json", scopes=scopes)
client = gspread.authorize(creds)

# !!! COLOQUE SEU LINK AQUI !!!
spreadsheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1juyOfIh0ZqsfJjN0p3gD8pKaAIX0R6IAPG9vysl7yWI/edit") 

# --- ANÁLISE ---
print(f"📂 Lendo aba: '{ABA_ALVO}'...")
try:
    ws = spreadsheet.worksheet(ABA_ALVO)
    df = pd.DataFrame(ws.get_all_records())
    df.columns = df.columns.str.strip() # Limpeza de espaços
except Exception as e:
    print(f"Erro ao ler aba: {e}")
    exit()

print(f"📊 Total de linhas na aba original: {len(df)}")

if 'Data Final' in df.columns:
    # Aplica a conversão
    df['Data_Obj'] = converter_data_hibrida(df['Data Final'])
    
    # Filtros
    filtro_mes_certo = (df['Data_Obj'].dt.month == MES_ESPERADO) & (df['Data_Obj'].dt.year == ANO_ESPERADO)
    
    aceitas = df[filtro_mes_certo]
    rejeitadas = df[~filtro_mes_certo]
    
    print(f"✅ Aceitas (São de Novembro): {len(aceitas)}")
    print(f"❌ Rejeitadas (Total): {len(rejeitadas)}")
    
    if not rejeitadas.empty:
        print("\n🔍 ANÁLISE DAS REJEITADAS:")
        
        # 1. Datas de Outros Meses
        outros_meses = rejeitadas[rejeitadas['Data_Obj'].notna()]
        if not outros_meses.empty:
            print(f"   -> {len(outros_meses)} linhas têm datas válidas, mas fora de Novembro:")
            print(outros_meses[['Data Final', 'Data_Obj']].head(10).to_string(index=False))
            
        # 2. Datas Inválidas/Vazias
        invalidas = rejeitadas[rejeitadas['Data_Obj'].isna()]
        if not invalidas.empty:
            print(f"\n   -> {len(invalidas)} linhas têm Data Final vazia ou ilegível:")
            # Tenta mostrar o Link ou ID para você achar
            col_ref = 'Link' if 'Link' in df.columns else 'ID'
            if col_ref in df.columns:
                print(invalidas[[col_ref, 'Data Final']].head(5).to_string(index=False))
            else:
                print(invalidas['Data Final'].head(5).to_string(index=False))
else:
    print("❌ Erro: Coluna 'Data Final' não encontrada.")