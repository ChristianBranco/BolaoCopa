import pandas as pd
import json

excel = pd.ExcelFile('BOLAO.xlsx')
org = pd.read_excel('BOLAO.xlsx', sheet_name='ORGANIZAÇÃO')
mm = pd.read_excel('BOLAO.xlsx', sheet_name='MATA-MATA 32')

participants = sorted([c for c in org.columns[6:-2] if c and 'Unnamed' not in str(c)])

games = []

for sheet_name, df in [('ORGANIZAÇÃO', org), ('MATA-MATA 32', mm)]:
  for idx, row in df.iterrows():
    match = row.get('CONFRONTOS','')
    if not match or pd.isna(match):
      continue

    game_id = len(games)
    date = row.get('DATAS')
    date_str = ''
    if pd.notna(date):
      if isinstance(date, str):
        date_str = date.split()[0]
      else:
        try:
          date_str = date.strftime('%d/%m')
        except:
          date_str = str(date)
    
    guesses = {}
    for p in participants:
      guess = row.get(p, '')
      guesses[p] = str(guess).upper() if pd.notna(guess) and guess != '' else ''

    resultado = row.get('RESULTADO','')
    resultado = str(resultado).upper() if pd.notna(resultado) and resultado != '' else ''

    gols_a = row.get('Gols A','')
    gols_a = str(int(gols_a)) if pd.notna(gols_a) and gols_a != '' else ''

    gols_b = row.get('Gols B','')
    gols_b = str(int(gols_b)) if pd.notna(gols_b) and gols_b != '' else ''

    # Ler coluna "Passou" (classificado)
    passou = row.get('Passou', '')
    passou = str(passou).strip() if pd.notna(passou) and passou != '' else ''

    # Detectar se foi decidido nos pênaltis
    penalties = False
    if resultado == 'E' and passou:  # Empate E alguém se classificou = pênaltis
      penalties = True

    games.append({
      'id': game_id,
      'match': match,
      'date': date_str,
      'resultado': resultado,
      'gols_a': gols_a,
      'gols_b': gols_b,
      'guesses': guesses,
      'sheet': sheet_name,
      'passou': passou,
      'penalties': penalties
    })

results = {'participants': participants, 'games': games}
with open('results.json', 'w', encoding='utf-8') as f:
  json.dump(results, f, ensure_ascii=False, indent=2)

print(f"✅ results.json gerado: {sum(1 for g in games if g['sheet']=='ORGANIZAÇÃO')} jogos (Organização) + {sum(1 for g in games if g['sheet']=='MATA-MATA 32')} jogos (Mata-Mata) = {len(games)} total")
print(f"   {sum(1 for g in games if g.get('resultado'))} com resultado | {len(participants)} participantes")
print(f"   {sum(1 for g in games if g.get('penalties'))} jogos decididos nos pênaltis")
