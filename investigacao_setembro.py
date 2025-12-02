import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# --- CONFIGURAÇÃO ---
ABA_ALVO = "Setembro 2025"
MES_ESPERADO = 9
ANO_ESPERADO = 2025

# --- CONEXÃO ---
print("🕵️ Conectando...")
scopes = ["https://www.googleapis.com/auth/spreadsheets"]
# Se estiver rodando local e tiver o arquivo json:
creds = Credentials.from_service_account_file("google_credentials.json", scopes=scopes)
# Se estiver no Streamlit Cloud, teria que usar st.secrets, mas assumo que está rodando local para teste.

client = gspread.authorize(creds)
# !!! SUBSTITUA PELA SUA URL !!!
spreadsheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1juyOfIh0ZqsfJjN0p3gD8pKaAIX0R6IAPG9vysl7yWI/edit?gid=1605683471#gid=1605683471") 

# --- ANÁLISE ---
print(f"📂 Lendo aba: '{ABA_ALVO}'...")
ws = spreadsheet.worksheet(ABA_ALVO)
df = pd.DataFrame(ws.get_all_records())

# Limpeza de colunas
df.columns = df.columns.str.strip()

print(f"📊 Total de linhas na aba: {len(df)}")

if 'Data Final' not in df.columns:
    print("❌ ERRO CRÍTICO: Coluna 'Data Final' não encontrada!")
    exit()

# Conversão Híbrida (A mesma do seu sistema)
print("⚙️ Analisando datas...")
datas_br = pd.to_datetime(df['Data Final'], dayfirst=True, errors='coerce')
falhas = datas_br.isna()
if falhas.any():
    datas_iso = pd.to_datetime(df.loc[falhas, 'Data Final'], format='mixed', errors='coerce')
    datas_br = datas_br.fillna(datas_iso)

df['Data_Obj'] = datas_br

# --- O PENTE FINO ---
# 1. Datas Inválidas (Vazias ou Lixo)
df_invalidas = df[df['Data_Obj'].isna()]

# 2. Datas de Outros Meses (Válidas, mas fora de Setembro)
df_outros_meses = df[
    (df['Data_Obj'].notna()) & 
    ((df['Data_Obj'].dt.month != MES_ESPERADO) | (df['Data_Obj'].dt.year != ANO_ESPERADO))
]

# 3. Datas Corretas
df_corretas = df[
    (df['Data_Obj'].dt.month == MES_ESPERADO) & 
    (df['Data_Obj'].dt.year == ANO_ESPERADO)
]

print("\n" + "="*40)
print("RESUMO DA INVESTIGAÇÃO")
print("="*40)
print(f"✅ Aceitas (Setembro): {len(df_corretas)}")
print(f"❌ Rejeitadas (Total):  {len(df) - len(df_corretas)}")
print("-" * 20)
print(f"   -> Datas Vazias ou Inválidas: {len(df_invalidas)}")
print(f"   -> Datas de OUTROS Meses:     {len(df_outros_meses)}")

if not df_outros_meses.empty:
    print("\n⚠️  EXEMPLOS DE OUTROS MESES ENCONTRADOS:")
    print(df_outros_meses[['Data Final', 'Data_Obj']].head(5).to_string(index=False))

if not df_invalidas.empty:
    print("\n⚠️  EXEMPLOS DE DATAS INVÁLIDAS/VAZIAS:")
    # Mostra Link ou ID para ajudar a achar
    col_ref = 'Link' if 'Link' in df.columns else 'ID'
    print(df_invalidas[[col_ref, 'Data Final']].head(5).to_string(index=False))