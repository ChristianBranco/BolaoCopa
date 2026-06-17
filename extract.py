"""
Lê a planilha BOLAO.xlsx e gera results.json para o site.
Execute localmente com:  python extract.py
O GitHub Actions roda isso automaticamente quando você faz upload do Excel.
"""
import json, sys
import pandas as pd

EXCEL = "BOLAO.xlsx"

try:
    df = pd.read_excel(EXCEL, sheet_name="ORGANIZAÇÃO", header=0)
except FileNotFoundError:
    print(f"❌ Arquivo {EXCEL} não encontrado.")
    sys.exit(1)

RESULTADO_COL = "RESULTADO"
MATCH_COL     = "CONFRONTOS"
DATE_COL      = "DATAS"
GOLS_A_COL    = "Gols A"
GOLS_B_COL    = "Gols B"

cols = list(df.columns)
start = cols.index("Unnamed: 5") + 1
end   = cols.index("Gols A")
participants = cols[start:end]

results = {}
games   = []

for i, row in df.iterrows():
    match = row.get(MATCH_COL, "")
    if pd.isna(match) or str(match).strip() == "" or str(match) == "nan":
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

    games.append({
        "id":        i,
        "match":     str(match).strip(),
        "date":      date_str,
        "resultado": resultado,
        "gols_a":    gols_a,
        "gols_b":    gols_b,
        "guesses":   guesses,
    })

    if resultado:
        results[i] = resultado

output = {
    "participants": participants,
    "games":        games,
    "results":      results,
}

with open("results.json", "w", encoding="utf-8") as f:
    json.dump(output, f, ensure_ascii=False, indent=2)

played = sum(1 for g in games if g["resultado"])
print(f"✅ results.json gerado: {len(games)} jogos, {played} com resultado, {len(participants)} participantes.")
