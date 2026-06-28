"""
Lê a planilha BOLAO.xlsx (abas ORGANIZAÇÃO e MATA-MATA 32) e gera results.json
Execute localmente com:  python extract.py
O GitHub Actions roda isso automaticamente quando você faz upload do Excel.
"""
import json, sys
import pandas as pd

EXCEL = "BOLAO.xlsx"

try:
    excel = pd.ExcelFile(EXCEL)
except FileNotFoundError:
    print(f"❌ Arquivo {EXCEL} não encontrado.")
    sys.exit(1)

SHEETS = ["ORGANIZAÇÃO", "MATA-MATA 32"]
RESULTADO_COL = "RESULTADO"
MATCH_COL     = "CONFRONTOS"
DATE_COL      = "DATAS"
GOLS_A_COL    = "Gols A"
GOLS_B_COL    = "Gols B"

all_games = []
participants = []
results = {}

# Processar ambas as abas
for sheet_name in SHEETS:
    if sheet_name not in excel.sheet_names:
        print(f"⚠️ Aba '{sheet_name}' não encontrada.")
        continue
    
    df = pd.read_excel(EXCEL, sheet_name=sheet_name, header=0)
    cols = list(df.columns)
    
    # Encontrar participantes (entre "Unnamed: 5" e "Gols A")
    try:
        start = cols.index("Unnamed: 5") + 1
        end   = cols.index("Gols A")
        sheet_participants = cols[start:end]
        
        if not participants:
            participants = sheet_participants
    except ValueError:
        print(f"⚠️ Estrutura não reconhecida em '{sheet_name}'")
        continue

    game_id = len(all_games)  # ID global
    
    for i, row in df.iterrows():
        match = row.get(MATCH_COL, "")
        if pd.isna(match) or str(match).strip() in ["", "nan"]:
            continue
        if str(match).strip() == str(participants[0]):
            continue

        resultado = row.get(RESULTADO_COL, "")
        resultado = "" if pd.isna(resultado) or str(resultado) in ["#x#","nan"] else str(resultado).strip()

        date_val = row.get(DATE_COL, "")
        date_str = date_val.strftime("%d/%m") if pd.notna(date_val) and hasattr(date_val, "strftime") else ""

        ga = row.get(GOLS_A_COL, "")
        gb = row.get(GOLS_B_COL, "")
        gols_a = "" if pd.isna(ga) else str(int(ga))
        gols_b = "" if pd.isna(gb) else str(int(gb))

        guesses = {}
        for p in participants:
            val = row.get(p, "")
            guesses[p] = "" if pd.isna(val) or str(val) == "nan" else str(val).strip()

        all_games.append({
            "id":        game_id,
            "sheet":     sheet_name,
            "match":     str(match).strip(),
            "date":      date_str,
            "resultado": resultado,
            "gols_a":    gols_a,
            "gols_b":    gols_b,
            "guesses":   guesses,
        })

        if resultado:
            results[game_id] = resultado

        game_id += 1

output = {
    "participants": list(participants),
    "games":        all_games,
    "results":      results,
}

with open("results.json", "w", encoding="utf-8") as f:
    json.dump(output, f, ensure_ascii=False, indent=2)

org_games = sum(1 for g in all_games if g["sheet"] == "ORGANIZAÇÃO")
mm_games = sum(1 for g in all_games if g["sheet"] == "MATA-MATA 32")
played = sum(1 for g in all_games if g["resultado"])

print(f"✅ results.json gerado: {org_games} jogos (Organização) + {mm_games} jogos (Mata-Mata) = {len(all_games)} total")
print(f"   {played} com resultado | {len(participants)} participantes")
