#!/usr/bin/env python3
"""
MODELO 3 — IMPACTO DE UM AUMENTO DE 10% NA POPULAÇÃO IDOSA (65+)
════════════════════════════════════════════════════════════════════════════
OBJETIVO
Simular o que acontece à procura e cobertura hospitalar do país (camas
necessárias, défice de camas por ULS) se a população idosa (faixa
[65-120[) crescer 10%, mantendo a oferta de camas constante. Isto é
relevante porque os idosos geram muito mais procura hospitalar por
habitante do que as restantes faixas etárias.

MÉTODO
1) A faixa etária [65-120[ é escalada em +10% em dois locais:
     a) freguesias_limpo.xlsx -> coluna de população idosa por freguesia
     b) Excel_Procura_Portugal.xlsx -> linhas da faixa [65-120[
        (usadas para recalcular a procura de camas, mantendo a mesma
        "Demora Média Ponderada" -> Internamentos e Camas Necessárias
        escalam linearmente com a população da faixa, tal como no
        modelo base: Internamentos = Σ(Prob × Pop_faixa))
2) Recalcula-se a Cobertura Camas_Necessárias vs Camas (Oferta) por ULS,
   igual ao KPI5 do modelo base, para ANTES e DEPOIS do choque demográfico.

OUTPUT: Modelo3_Impacto_Pop_Idosa.xlsx com:
  - Resumo executivo (antes vs depois)
  - Cobertura de camas por ULS (antes vs depois)
  - Por freguesia: população e camas necessárias por faixa etária (antes/depois)
"""
import pandas as pd
import numpy as np
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ── PARÂMETROS DO MODELO ────────────────────────────────────────────────────
AUMENTO_IDOSOS   = 0.10          # +10% população idosa
FAIXA_IDOSA_FREG = '[65,120['    # nome da coluna em freguesias_limpo.xlsx
FAIXA_IDOSA_PROC = '[65-120['    # nome da faixa em Excel_Procura_Portugal.xlsx

print(f"📁 Pasta: {SCRIPT_DIR}\n")

# ═══════════════════════════════════════════════════════════════════════════
# 1. CARREGAR DADOS
# ═══════════════════════════════════════════════════════════════════════════
print("📂 Carregando Excel...")
df_freg = pd.read_excel(os.path.join(SCRIPT_DIR, 'freguesias_limpo.xlsx'), sheet_name=0)
df_freg.columns = df_freg.columns.str.strip()

df_oferta = pd.read_excel(os.path.join(SCRIPT_DIR, 'Excel_Oferta_Portugal.xlsx'),
                           sheet_name='1. Oferta por ULS', header=2)
df_oferta.columns = df_oferta.columns.str.strip()

df_procura = pd.read_excel(os.path.join(SCRIPT_DIR, 'Excel_Procura_Portugal.xlsx'),
                            sheet_name='2. Agregado por Faixa', header=4)
df_procura.columns = df_procura.columns.str.strip()
print("✅ Ficheiros carregados\n")


def normalize_text(s):
    import re, unicodedata
    s = str(s).strip().lower()
    s = unicodedata.normalize('NFKD', s).encode('ascii', 'ignore').decode()
    return re.sub(r'\s+', ' ', s)


def normalize_uls(name):
    import re, unicodedata
    s = str(name).strip()
    s = re.sub(r',?\s*E\.?P\.?E\.?\.?$', '', s, flags=re.IGNORECASE)
    s = re.sub(r',?\s*PPP\.?$', '', s, flags=re.IGNORECASE)
    s = s.lower()
    s = unicodedata.normalize('NFKD', s).encode('ascii', 'ignore').decode()
    s = re.sub(r'[\/\-]', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()


ULS_ALIASES = {
    'unidade local de saude castelo branco': 'unidade local de saude de castelo branco',
    'unidade local de saude de vila nova de gaia espinho': 'unidade local de saude de gaia espinho',
    'unidade local de saude de dao lafoes': 'unidade local de saude de viseu dao lafoes',
}


def normalize_uls_final(name):
    return ULS_ALIASES.get(normalize_uls(name), normalize_uls(name))


# ═══════════════════════════════════════════════════════════════════════════
# 2. CAPACIDADE DE CAMAS POR ULS
# ═══════════════════════════════════════════════════════════════════════════
df_freg['ULS_norm'] = df_freg['ULS'].apply(normalize_uls_final)
df_oferta['ULS_norm'] = df_oferta['ULS'].apply(normalize_uls_final)
camas_col = [c for c in df_oferta.columns if 'cama' in c.lower()][0]

# Nome de exibição da ULS (um por ULS_norm) + capacidade agregada de camas
capacidade_uls = df_freg[['ULS_norm', 'ULS']].drop_duplicates('ULS_norm').reset_index(drop=True)
oferta_uls = df_oferta.groupby('ULS_norm')[camas_col].sum().rename('Capacidade').reset_index()
capacidade_uls = capacidade_uls.merge(oferta_uls, on='ULS_norm', how='left')
capacidade_uls['Capacidade'] = capacidade_uls['Capacidade'].fillna(capacidade_uls['Capacidade'].median())

# ═══════════════════════════════════════════════════════════════════════════
# 3. PROCURA ANTES E DEPOIS DO AUMENTO DEMOGRÁFICO (por freguesia)
# ═══════════════════════════════════════════════════════════════════════════
print(f"👴 A simular +{int(AUMENTO_IDOSOS*100)}% de população idosa...")

demanda_col = [c for c in df_procura.columns if 'camas necess' in c.lower()][0]
distrito_p = [c for c in df_procura.columns if 'distrito' in c.lower()][0]
concelho_p = [c for c in df_procura.columns if 'concelho' in c.lower()][0]
freguesia_p = [c for c in df_procura.columns if c.lower() == 'freguesia'][0]
faixa_p = [c for c in df_procura.columns if 'faixa' in c.lower()][0]

df_procura['_key'] = (df_procura[distrito_p].apply(normalize_text) + '|' +
                       df_procura[concelho_p].apply(normalize_text) + '|' +
                       df_procura[freguesia_p].apply(normalize_text))

# ANTES: procura de camas necessárias, agregada por freguesia (todas as faixas)
demanda_antes = df_procura.groupby('_key')[demanda_col].sum().rename('Demanda_Antes')

# DEPOIS: escala só a linha da faixa idosa em +10% (Internamentos, Dias e
# Camas Necessárias são todos diretamente proporcionais à população da
# faixa, mantendo a mesma probabilidade/demora média -> escalar a coluna
# final "Camas Necessárias" da faixa idosa é equivalente a escalar a
# população de origem)
df_procura_depois = df_procura.copy()
mask_idoso = df_procura_depois[faixa_p] == FAIXA_IDOSA_PROC
if mask_idoso.sum() == 0:
    raise ValueError(f"Não encontrei a faixa '{FAIXA_IDOSA_PROC}' na coluna '{faixa_p}'. "
                      f"Valores existentes: {df_procura_depois[faixa_p].unique()}")
df_procura_depois.loc[mask_idoso, demanda_col] *= (1 + AUMENTO_IDOSOS)
demanda_depois = df_procura_depois.groupby('_key')[demanda_col].sum().rename('Demanda_Depois')

df_freg['_key'] = (df_freg['Distrito'].apply(normalize_text) + '|' +
                    df_freg['Concelho'].apply(normalize_text) + '|' +
                    df_freg['Freguesia'].apply(normalize_text))
df_freg = df_freg.merge(demanda_antes, on='_key', how='left').merge(demanda_depois, on='_key', how='left')
df_freg['Demanda_Antes'] = df_freg['Demanda_Antes'].fillna(df_freg['Demanda_Antes'].median())
df_freg['Demanda_Depois'] = df_freg['Demanda_Depois'].fillna(df_freg['Demanda_Depois'].median())

acrescimo_nacional = df_freg['Demanda_Depois'].sum() - df_freg['Demanda_Antes'].sum()
print(f"   Camas necessárias nacionais: {df_freg['Demanda_Antes'].sum():.1f} -> "
      f"{df_freg['Demanda_Depois'].sum():.1f}  (+{acrescimo_nacional:.1f})\n")

# Também escalamos a coluna de população idosa da freguesia (para métricas de
# tempo de acesso ponderado por idade)
df_freg['pop_idoso_antes'] = df_freg[FAIXA_IDOSA_FREG]
df_freg['pop_idoso_depois'] = df_freg[FAIXA_IDOSA_FREG] * (1 + AUMENTO_IDOSOS)
df_freg['Populacao_Depois'] = df_freg['Populacão total'] + (df_freg['pop_idoso_depois'] - df_freg['pop_idoso_antes'])


# ═══════════════════════════════════════════════════════════════════════════
# 4. COBERTURA CAMAS NECESSÁRIAS x OFERTA POR ULS — ANTES vs DEPOIS
# ═══════════════════════════════════════════════════════════════════════════
print("📊 A recalcular cobertura de camas por ULS...")
cobertura_antes = df_freg.groupby('ULS_norm')['Demanda_Antes'].sum().rename('Camas_Necessarias_Antes')
cobertura_depois = df_freg.groupby('ULS_norm')['Demanda_Depois'].sum().rename('Camas_Necessarias_Depois')
cobertura = pd.concat([cobertura_antes, cobertura_depois], axis=1).reset_index()
cobertura = cobertura.merge(capacidade_uls[['ULS_norm', 'ULS', 'Capacidade']],
                             on='ULS_norm', how='left')
cobertura['Cobertura_%_Antes'] = np.minimum(100, cobertura['Capacidade'] / cobertura['Camas_Necessarias_Antes'] * 100).round(1)
cobertura['Cobertura_%_Depois'] = np.minimum(100, cobertura['Capacidade'] / cobertura['Camas_Necessarias_Depois'] * 100).round(1)
cobertura['Camas_Em_Falta_Antes'] = np.maximum(0, cobertura['Camas_Necessarias_Antes'] - cobertura['Capacidade']).round(1)
cobertura['Camas_Em_Falta_Depois'] = np.maximum(0, cobertura['Camas_Necessarias_Depois'] - cobertura['Capacidade']).round(1)
cobertura['Delta_Camas_Em_Falta'] = (cobertura['Camas_Em_Falta_Depois'] - cobertura['Camas_Em_Falta_Antes']).round(1)


print("\n" + "=" * 70)
print("RESUMO — IMPACTO DE +10% DE POPULAÇÃO IDOSA")
print("=" * 70)
print(f"Camas necessárias (nacional)   : {df_freg['Demanda_Antes'].sum():.0f} -> {df_freg['Demanda_Depois'].sum():.0f}  "
      f"(+{acrescimo_nacional:.0f})")
print(f"Cobertura média nacional (%)   : {cobertura['Cobertura_%_Antes'].mean():.1f} -> {cobertura['Cobertura_%_Depois'].mean():.1f}")
print("=" * 70)


# ═══════════════════════════════════════════════════════════════════════════
# 7. OUTPUT EXCEL
# ═══════════════════════════════════════════════════════════════════════════
print("\n💾 A gerar ficheiro Excel...")
wb = Workbook()
HEADER_FILL = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF")


def style_header(ws, ncols, row=1):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal='center')


def autofit(ws, ncols, max_width=45):
    for c in range(1, ncols + 1):
        letter = get_column_letter(c)
        maxlen = max((len(str(ws.cell(row=r, column=c).value or '')) for r in range(1, ws.max_row + 1)), default=10)
        ws.column_dimensions[letter].width = min(max_width, max(10, maxlen + 2))


# --- Folha 1: Resumo ----------------------------------------------------------
ws = wb.active
ws.title = "1. Resumo"
ws['B2'] = f"MODELO 3 — IMPACTO DE +{int(AUMENTO_IDOSOS*100)}% NA POPULAÇÃO IDOSA (65+)"
ws['B2'].font = Font(bold=True, size=14, color="C00000")
linhas = [
    ("", ""),
    ("Camas necessárias nacionais (ANTES)", round(df_freg['Demanda_Antes'].sum(), 1)),
    ("Camas necessárias nacionais (DEPOIS)", round(df_freg['Demanda_Depois'].sum(), 1)),
    ("Acréscimo de camas necessárias", round(acrescimo_nacional, 1)),
    ("", ""),
    ("Cobertura de camas média nacional (ANTES) %", round(cobertura['Cobertura_%_Antes'].mean(), 1)),
    ("Cobertura de camas média nacional (DEPOIS) %", round(cobertura['Cobertura_%_Depois'].mean(), 1)),
    ("", ""),
    ("Nota", "A população idosa (65+) foi escalada em +10% e a procura de camas "
             "recalculada proporcionalmente; a oferta de camas por ULS mantém-se "
             "constante, pelo que a cobertura oferta/necessidade piora."),
]
for i, (k, v) in enumerate(linhas, 4):
    ws.cell(row=i, column=2, value=k).font = Font(bold=True)
    ws.cell(row=i, column=3, value=v)
ws.column_dimensions['B'].width = 48
ws.column_dimensions['C'].width = 30

# --- Folha 2: Cobertura por ULS -----------------------------------------------
ws2 = wb.create_sheet("2. Cobertura_ULS")
cols2 = ['ULS', 'Camas_Necessarias_Antes', 'Camas_Necessarias_Depois', 'Capacidade',
         'Cobertura_%_Antes', 'Cobertura_%_Depois', 'Camas_Em_Falta_Antes',
         'Camas_Em_Falta_Depois', 'Delta_Camas_Em_Falta']
out2 = cobertura[cols2].sort_values('Delta_Camas_Em_Falta', ascending=False)
for j, h in enumerate(out2.columns, 1):
    ws2.cell(row=1, column=j, value=h)
for i, (_, row) in enumerate(out2.iterrows(), 2):
    for j, h in enumerate(out2.columns, 1):
        ws2.cell(row=i, column=j, value=row[h])
style_header(ws2, len(out2.columns))
ws2.freeze_panes = 'A2'
autofit(ws2, len(out2.columns))

# --- Folha 3: Por Freguesia — População e camas necessárias por faixa etária -
# Por freguesia: população em cada faixa etária e camas necessárias, antes
# e depois do choque demográfico (+10% na faixa idosa [65,120[).
ws4 = wb.create_sheet("3. Por_Freguesia")
ws4['A1'] = ("O que muda ANTES vs DEPOIS é a Pop_[65,120[ (+10%) e a procura de "
             "camas que essa faixa gera; as restantes faixas etárias mantêm-se.")
ws4['A1'].font = Font(italic=True, size=9, color="666666")

faixas_fixas = ['[0-1[', '[1,5[', '[5,15[', '[15,25[', '[25,45[', '[45,65[']

out4 = df_freg[['Distrito', 'Concelho', 'Freguesia', 'ULS']].copy()
for faixa in faixas_fixas:
    out4[f'Pop_{faixa}'] = df_freg[faixa]
# Faixa idosa: ANTES e DEPOIS lado a lado (a única que muda no cenário)
out4['Pop_[65,120[_Antes'] = df_freg['pop_idoso_antes'].round(0)
out4['Pop_[65,120[_Depois'] = df_freg['pop_idoso_depois'].round(0)
out4['Camas_Necessarias_Antes'] = df_freg['Demanda_Antes'].round(2)
out4['Camas_Necessarias_Depois'] = df_freg['Demanda_Depois'].round(2)
out4 = out4.sort_values(['Distrito', 'Concelho', 'Freguesia'])

for j, h in enumerate(out4.columns, 1):
    ws4.cell(row=2, column=j, value=h)
for i, (_, row) in enumerate(out4.iterrows(), 3):
    for j, h in enumerate(out4.columns, 1):
        ws4.cell(row=i, column=j, value=row[h])
style_header(ws4, len(out4.columns), row=2)
ws4.freeze_panes = 'A3'
autofit(ws4, len(out4.columns))

output_path = os.path.join(SCRIPT_DIR, 'Modelo3_Impacto_Pop_Idosa.xlsx')
wb.save(output_path)
print(f"✅ Ficheiro guardado em: {output_path}")