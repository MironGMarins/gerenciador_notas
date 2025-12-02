import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# --- CONFIGURAÇÃO ---
ABA_ORIGEM = "Novembro 2025"
ABA_DESTINO = "Total BaseCamp para Notas"

# --- CONEXÃO ---
print("🕵️ Conectando...")
scopes = ["https://www.googleapis.com/auth/spreadsheets"]
creds = Credentials.from_service_account_file("google_credentials.json", scopes=scopes)
client = gspread.authorize(creds)

# !!! SUBSTITUA PELO SEU LINK !!!
spreadsheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1juyOfIh0ZqsfJjN0p3gD8pKaAIX0R6IAPG9vysl7yWI/edit?gid=1052350069#gid=1052350069") 

# --- CARREGAR DADOS ---
print(f"1️⃣ Lendo aba de origem: '{ABA_ORIGEM}'...")
ws_origem = spreadsheet.worksheet(ABA_ORIGEM)
df_origem = pd.DataFrame(ws_origem.get_all_records())

# Garante que temos um ID para comparar (Gera pelo Link se necessário)
if 'Link' in df_origem.columns:
    df_origem['ID_Real'] = df_origem['Link'].astype(str).str.split('/').str[-1].str.strip()
    # Remove linhas vazias
    df_origem = df_origem[df_origem['ID_Real'] != '']
else:
    print("Erro: Aba origem sem coluna Link.")
    exit()

print(f"2️⃣ Lendo aba de destino: '{ABA_DESTINO}'...")
ws_destino = spreadsheet.worksheet(ABA_DESTINO)
df_destino = pd.DataFrame(ws_destino.get_all_records())

if 'Link' in df_destino.columns:
    df_destino['ID_Real'] = df_destino['Link'].astype(str).str.split('/').str[-1].str.strip()
else:
    print("Erro: Aba destino sem coluna Link.")
    exit()

# --- O CRUZAMENTO (QUEM FALTOU?) ---
ids_origem = set(df_origem['ID_Real'])
ids_destino = set(df_destino['ID_Real'])

# IDs que estão na Origem mas NÃO estão no Destino
ids_perdidos = ids_origem - ids_destino

print("\n" + "="*40)
print("RELATÓRIO DE PERDAS")
print("="*40)
print(f"Total na Origem (bruto): {len(df_origem)}")
print(f"Total no Destino (consolidado): {len(df_destino)}")
print(f"⚠️  Tarefas que existem em '{ABA_ORIGEM}' mas sumiram no Destino: {len(ids_perdidos)}")

if len(ids_perdidos) > 0:
    print("\n🔍 LISTA DAS TAREFAS PERDIDAS E SUAS DATAS:")
    print(f"{'ID':<15} | {'Data Final (Original)':<20} | {'Encarregado'}")
    print("-" * 60)
    
    # Pega os detalhes das perdidas
    df_perdidas = df_origem[df_origem['ID_Real'].isin(ids_perdidos)]
    
    for index, row in df_perdidas.iterrows():
        d_id = str(row.get('ID_Real', ''))
        d_data = str(row.get('Data Final', 'VAZIO'))
        d_enc = str(row.get('Encarregado', ''))
        
        # Filtra para não mostrar as 1000 antigas (Junho/Maio)
        # Mostra apenas se a data parecer ser de 2025 e próxima a Novembro
        print(f"{d_id:<15} | {d_data:<20} | {d_enc}")

    print("\n💡 DICA:")
    print("Se as datas acima forem de Outubro (10) ou Dezembro (12), elas foram barradas pelo filtro de mês.")
    print("Se forem datas antigas (Junho, Maio), o filtro funcionou corretamente.")
else:
    print("✅ Nenhuma perda inexplicável encontrada. Todos os IDs da origem estão no destino.")