import pandas as pd
import json

excel = pd.ExcelFile('BOLAO.xlsx')
org = pd.read_excel('BOLAO.xlsx', sheet_name='ORGANIZAÇÃO')
mm = pd.read_excel('BOLAO.xlsx', sheet_name='MATA-MATA 32')

# Excluir colunas que não são participantes
exclude_cols = ['DATAS', 'CONFRONTOS', 'TIME A', 'TIME B', 'PLACAR', 'Penalti', 'Prorrogação', 
                'Unnamed: 7', 'Gols A', 'Gols B', 'RESULTADO', 'Passou']

participants = sorted([c for c in org.columns 
                       if c and 'Unnamed' not in str(c) and c not in exclude_cols])

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

    # Ler colunas de Pênalti e Prorrogação
    penalti = row.get('Penalti', '')
    penalti = str(penalti).strip().upper() if pd.notna(penalti) and penalti != '' else ''

    prorrogacao = row.get('Prorrogação', '')
    prorrogacao = str(prorrogacao).strip().upper() if pd.notna(prorrogacao) and prorrogacao != '' else ''

    # Determinar quem passou e como
    passou = ""
    tipo_decisao = ""
    
    times = match.split(" x ")
    timeA = times[0].strip() if len(times) > 0 else ""
    timeB = times[1].strip() if len(times) > 1 else ""

    # Se há pênalti
    if penalti:
      tipo_decisao = "Pênaltis"
      passou = timeA if penalti == "A" else timeB if penalti == "B" else ""
    
    # Se há prorrogação
    elif prorrogacao:
      tipo_decisao = "Prorrogação"
      passou = timeA if prorrogacao == "A" else timeB if prorrogacao == "B" else ""
    
    # Se resultado normal (90 min)
    elif resultado:
      if resultado == "A":
        passou = timeA
      elif resultado == "B":
        passou = timeB
      elif resultado == "E":
        passou = ""

    penalties = (penalti != "")
    overtime = (prorrogacao != "")

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
      'penalties': penalties,
      'overtime': overtime,
      'tipo_decisao': tipo_decisao
    })

results = {'participants': participants, 'games': games}
with open('results.json', 'w', encoding='utf-8') as f:
  json.dump(results, f, ensure_ascii=False, indent=2)

print(f"✅ results.json gerado")
print(f"   {sum(1 for g in games if g['sheet']=='ORGANIZAÇÃO')} jogos (Organização) + {sum(1 for g in games if g['sheet']=='MATA-MATA 32')} (Mata-Mata)")
print(f"   {len(participants)} participantes")
print(f"   {sum(1 for g in games if g.get('penalties'))} jogos nos pênaltis")
print(f"   {sum(1 for g in games if g.get('overtime'))} jogos na prorrogação")
