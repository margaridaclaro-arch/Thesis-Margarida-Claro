#!/usr/bin/env python3
"""
MODELO DE IMPACTO COVID NOS INTERNAMENTOS E CAMAS
═══════════════════════════════════════════════════════════════════════════
Simula o que acontece à procura de camas hospitalares se os internamentos
por "Códigos para fins especiais" (categoria COVID) voltarem ao nível do
pico de 2021, por faixa etária, mantendo a oferta de camas e a geografia
(tempo de acesso) inalteradas.

FICHEIROS DE ENTRADA NECESSÁRIOS (na mesma pasta deste script):
  - freguesias_limpo.xlsx                     (população por freguesia/faixa, ULS, coords)
  - Excel_Oferta_Portugal.xlsx                 (camas por ULS)
  - Excel_Procura_Portugal.xlsx                (probabilidades e internamentos esperados
                                                 por freguesia/faixa/doença - baseline)
  - variação_de_internamentos.xlsx             (internamentos 2021 vs 2025, por faixa etária,
                                                 categoria "Códigos para fins especiais")

  O tempo de acesso (tempo_min) de cada freguesia ao hospital do seu ULS é calculado
  por este próprio script, pela rede viária real (OSRM), a partir das coordenadas em
  freguesias_limpo.xlsx — exatamente como no modelo.py original. Precisa de ligação
  à Internet e demora alguns minutos da primeira vez; fica em cache (_osrm_cache.json)
  para reexecuções seguintes serem instantâneas nesse passo.

FICHEIROS DE SAÍDA:
  - Dashboard_Impacto_COVID_Internamentos.xlsx  (dashboard interativo, estilo Modelo3)
  - Excel_Procura_Portugal_Cenario_COVID.xlsx   (procura antes/depois, por freguesia/faixa/ULS)
  - Excel_Oferta_Portugal_Referencia.xlsx       (cópia formatada da oferta de camas por ULS — constante no
                                                 cenário; tem nome diferente do ficheiro de entrada de propósito,
                                                 para nunca o sobrescrever)

LÓGICA DO CENÁRIO:
  1. Na tabela de probabilidades nacionais (Excel_Procura_Portugal.xlsx), a doença de
     código 22 ("Códigos para fins especiais") é a categoria COVID - confirma-se porque os
     seus internamentos nacionais por faixa batem com os valores de 2025 no ficheiro de
     variação.
  2. Fator_Covid[faixa] = 1 / (1 - Variação[faixa]), onde Variação = (Internam_2021 -
     Internam_2025) / Internam_2021, medido no ficheiro de variação.
  3. Para cada freguesia × faixa etária: os internamentos do código 22 são multiplicados
     por Fator_Covid[faixa]; as restantes doenças mantêm-se. Dias e Camas Necessárias
     são recalculados a partir daí (Camas = Dias / 365).
  4. Agrega-se por freguesia e por ULS, cruza-se com a oferta (camas) para obter
     Cobertura_% e Camas_Em_Falta, ANTES e DEPOIS do choque.
"""
import os
import json
import time
from math import radians, cos, sin, asin, sqrt
import numpy as np
import pandas as pd
import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.chart import BarChart, Reference

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
print(f"📁 Pasta: {SCRIPT_DIR}\n")

# ═══════════════════════════════════════════════════════════════════════════
# CONFIGURAÇÃO DO MOTOR DE ROTAS (OpenStreetMap via OSRM) — igual ao modelo.py
# ═══════════════════════════════════════════════════════════════════════════
# Servidor público de demonstração do projeto OSRM. Gratuito, sem chave, mas
# com política de "fair use" (pode ficar lento com muitos pedidos seguidos).
# Para uso pesado/estável, ver instruções de servidor OSRM próprio no modelo.py.
OSRM_BASE_URL = "https://router.project-osrm.org"
OSRM_BATCH_SIZE = 90
OSRM_SLEEP_SECONDS = 1.0
OSRM_CACHE_FILE = os.path.join(SCRIPT_DIR, '_osrm_cache.json')

# ═══════════════════════════════════════════════════════════════════════════
# CONFIGURAÇÃO DE FICHEIROS
# ═══════════════════════════════════════════════════════════════════════════
F_FREGUESIAS = os.path.join(SCRIPT_DIR, 'freguesias_limpo.xlsx')
F_OFERTA = os.path.join(SCRIPT_DIR, 'Excel_Oferta_Portugal.xlsx')
F_PROCURA = os.path.join(SCRIPT_DIR, 'Excel_Procura_Portugal.xlsx')
F_VARIACAO = os.path.join(SCRIPT_DIR, 'variação_de_internamentos.xlsx')

OUT_DASHBOARD = os.path.join(SCRIPT_DIR, 'Dashboard_Impacto_COVID_Internamentos.xlsx')
OUT_PROCURA = os.path.join(SCRIPT_DIR, 'Excel_Procura_Portugal_Cenario_COVID.xlsx')
OUT_OFERTA = os.path.join(SCRIPT_DIR, 'Excel_Oferta_Portugal_Referencia.xlsx')

# ── Verificação prévia: os 4 ficheiros de entrada têm de estar na mesma
#    pasta deste script. Se faltar algum, avisa já aqui em vez de rebentar
#    a meio da execução. ────────────────────────────────────────────────
_ficheiros_necessarios = {
    'freguesias_limpo.xlsx': F_FREGUESIAS,
    'Excel_Oferta_Portugal.xlsx': F_OFERTA,
    'Excel_Procura_Portugal.xlsx': F_PROCURA,
    'variação_de_internamentos.xlsx': F_VARIACAO,
}
_em_falta = [nome for nome, caminho in _ficheiros_necessarios.items() if not os.path.isfile(caminho)]
if _em_falta:
    print("❌ Faltam ficheiros nesta pasta:")
    for nome in _em_falta:
        print(f"   - {nome}")
    print(f"\nColoca-os em: {SCRIPT_DIR}\ne volta a correr o script.")
    raise SystemExit(1)

COD_DOENCA_COVID = 22  # "Códigos para fins especiais"

KEYS = ['Distrito', 'Concelho', 'Freguesia', 'Faixa Etária']

header_font = Font(bold=True, color="FFFFFF")
header_fill_blue = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
header_fill_red = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")


def style_header(ws, ncols, row=1, fill=header_fill_blue):
    for col in range(1, ncols + 1):
        c = ws.cell(row=row, column=col)
        c.font = header_font
        c.fill = fill
        c.alignment = Alignment(horizontal="center")


def autofit(ws, ncols, min_width=10, max_width=48):
    for col in range(1, ncols + 1):
        letter = get_column_letter(col)
        max_len = min_width
        for cell in ws[letter]:
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[letter].width = min(max_len + 2, max_width)


def normalize_uls(name):
    import re
    import unicodedata
    s = str(name).strip()
    s = re.sub(r',?\s*E\.?P\.?E\.?\.?$', '', s, flags=re.IGNORECASE)
    s = re.sub(r',?\s*PPP\.?$', '', s, flags=re.IGNORECASE)
    s = s.lower()
    s = unicodedata.normalize('NFKD', s).encode('ascii', 'ignore').decode()
    s = re.sub(r'[\/\-]', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s


ULS_ALIASES = {
    'unidade local de saude castelo branco': 'unidade local de saude de castelo branco',
    'unidade local de saude de vila nova de gaia espinho': 'unidade local de saude de gaia espinho',
    'unidade local de saude de dao lafoes': 'unidade local de saude de viseu dao lafoes',
}


def normalize_uls_final(name):
    key = normalize_uls(name)
    return ULS_ALIASES.get(key, key)


def haversine(lat1, lon1, lat2, lon2):
    lat1, lon1, lat2, lon2 = map(radians, [lat1, lon1, lat2, lon2])
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    a = sin(dlat / 2) ** 2 + cos(lat1) * cos(lat2) * sin(dlon / 2) ** 2
    c = 2 * asin(sqrt(a))
    return 6371 * c


# Cache em disco: evita recalcular tudo de novo em reexecuções do script,
# já que as distâncias reais por estrada não mudam de um dia para o outro.
if os.path.exists(OSRM_CACHE_FILE):
    try:
        with open(OSRM_CACHE_FILE, 'r', encoding='utf-8') as f:
            _osrm_cache = json.load(f)
    except Exception:
        _osrm_cache = {}
else:
    _osrm_cache = {}


def _osrm_cache_key(lat_f, lon_f, lat_h, lon_h):
    return f"{round(lat_f, 5)},{round(lon_f, 5)}|{round(lat_h, 5)},{round(lon_h, 5)}"


def _osrm_save_cache():
    try:
        with open(OSRM_CACHE_FILE, 'w', encoding='utf-8') as f:
            json.dump(_osrm_cache, f)
    except Exception as e:
        print(f"   ⚠️  Não consegui gravar cache OSRM: {e}")


def osrm_table_batch(freg_rows, lat_h, lon_h):
    """Calcula distância (km) e tempo (min) REAIS POR ESTRADA de várias
    freguesias até UM hospital, numa única chamada ao serviço 'table' do OSRM.
    freg_rows: lista de tuplos (índice_no_df, lat, lon) das freguesias.
    Devolve: dict {índice_no_df: (distancia_km, tempo_min)}.
    Em caso de falha, ou de par sem rota, cai em haversine como reserva.
    """
    resultados = {}
    por_pedir = []
    for idx, lat_f, lon_f in freg_rows:
        key = _osrm_cache_key(lat_f, lon_f, lat_h, lon_h)
        if key in _osrm_cache:
            resultados[idx] = tuple(_osrm_cache[key])
        else:
            por_pedir.append((idx, lat_f, lon_f, key))

    if not por_pedir:
        return resultados

    coords = [f"{lon_f},{lat_f}" for _, lat_f, lon_f, _ in por_pedir]
    coords.append(f"{lon_h},{lat_h}")  # o hospital é sempre o destino (último ponto)
    n = len(por_pedir)
    sources = ";".join(str(i) for i in range(n))
    destination = str(n)

    url = (
        f"{OSRM_BASE_URL}/table/v1/driving/{';'.join(coords)}"
        f"?sources={sources}&destinations={destination}&annotations=distance,duration"
    )

    try:
        resp = requests.get(url, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        if data.get('code') != 'Ok':
            raise ValueError(f"OSRM devolveu code={data.get('code')}")
        distancias = data['distances']
        duracoes = data['durations']
        for i, (idx, lat_f, lon_f, key) in enumerate(por_pedir):
            dist_m = distancias[i][0]
            dur_s = duracoes[i][0]
            if dist_m is None or dur_s is None:
                dist_km = haversine(lat_f, lon_f, lat_h, lon_h)
                tempo_min = dist_km / 1.0
            else:
                dist_km = dist_m / 1000.0
                tempo_min = dur_s / 60.0
            resultados[idx] = (dist_km, tempo_min)
            _osrm_cache[key] = (dist_km, tempo_min)
    except Exception as e:
        print(f"   ⚠️  Erro OSRM neste lote (a usar haversine como reserva): {e}")
        for idx, lat_f, lon_f, key in por_pedir:
            dist_km = haversine(lat_f, lon_f, lat_h, lon_h)
            tempo_min = dist_km / 1.0
            resultados[idx] = (dist_km, tempo_min)

    time.sleep(OSRM_SLEEP_SECONDS)
    return resultados


# ═══════════════════════════════════════════════════════════════════════════
# 1. CARREGAR FICHEIROS
# ═══════════════════════════════════════════════════════════════════════════
print("📂 Carregando ficheiros...")

df_freg = pd.read_excel(F_FREGUESIAS, sheet_name=0)
df_freg.columns = df_freg.columns.str.strip()
if 'ULS' not in df_freg.columns:
    print(f"❌ '{os.path.basename(F_FREGUESIAS)}' não tem uma coluna 'ULS'. "
          f"Colunas encontradas: {list(df_freg.columns)}\n"
          f"   O ficheiro pode estar corrompido ou não ser o original — confirma-o e tenta de novo.")
    raise SystemExit(1)
print(f"   ✅ Freguesias: {len(df_freg)} linhas")

df_oferta = pd.read_excel(F_OFERTA, sheet_name='1. Oferta por ULS', header=2)
df_oferta.columns = df_oferta.columns.str.strip()
if 'ULS' not in df_oferta.columns:
    print(f"❌ '{os.path.basename(F_OFERTA)}' não tem uma coluna 'ULS' logo na 3ª linha do ficheiro "
          f"(como seria esperado). Colunas encontradas: {list(df_oferta.columns)}\n"
          f"   Isto costuma acontecer se o ficheiro foi substituído por um output deste script por engano, "
          f"ou aberto/gravado noutro programa que alterou as linhas. Usa o ficheiro ORIGINAL de "
          f"Excel_Oferta_Portugal.xlsx (3 linhas de cabeçalho + 1 linha por ULS) e tenta de novo.")
    raise SystemExit(1)
print(f"   ✅ Oferta: {len(df_oferta)} ULS")

df_agg = pd.read_excel(F_PROCURA, sheet_name='2. Agregado por Faixa', header=4)
df_agg.columns = df_agg.columns.str.strip()
df_agg = df_agg.rename(columns={
    'Pop. da Faixa': 'Pop_Faixa',
    'Internamentos Esperados  =  Σ(Prob × Pop. Faixa)': 'Internamentos_Antes',
    'Demora Média Pond. (dias)': 'Demora_Media_Pond',
    'Dias Necessários  =  Internamentos × Demora Média': 'Dias_Antes',
    'Camas Necessárias  =  Dias / 365': 'Camas_Antes',
})
print(f"   ✅ Procura agregada (baseline): {len(df_agg)} linhas (freguesia x faixa)")

df_det = pd.read_excel(F_PROCURA, sheet_name='1. Detalhe por Doenca', header=4)
df_det.columns = df_det.columns.str.strip()
df_covid = df_det[df_det['Cód.'] == COD_DOENCA_COVID].copy()
df_covid = df_covid.rename(columns={
    'Pop. da Faixa': 'Pop_Faixa',
    'Probabilidade': 'Prob_Covid',
    'Internamentos Esperados  =  Prob × Pop. Faixa': 'Internamentos_Covid_Antes',
    'Demora Média Nacional (dias)': 'Demora_Media_Covid',
    'Dias Necessários  =  Internamentos × Demora Média': 'Dias_Covid_Antes',
    'Camas Necessárias  =  Dias / 365': 'Camas_Covid_Antes',
})
print(f"   ✅ Detalhe filtrado a Cód={COD_DOENCA_COVID} (Códigos para fins especiais / COVID): {len(df_covid)} linhas")

# Tempo de acesso (tempo_min) calculado pela rede viária real (OSRM), a
# partir das coordenadas de freguesias_limpo.xlsx — agrupado por hospital
# (mesmo par lat/lon) para pedir, numa única chamada, a distância de várias
# freguesias até esse hospital.
print("   Calculando tempos de acesso reais por estrada (OpenStreetMap/OSRM)...")

lat_freq, lon_freq = 'Lat_Freguesia', 'Long_Freguesia'
lat_hosp, lon_hosp = 'Lat do Hospital', 'Long do Hospital'

df_freg['distancia_km'] = np.nan
df_freg['tempo_min'] = np.nan

_validos = df_freg.dropna(subset=[lat_freq, lon_freq, lat_hosp, lon_hosp])
_hospitais_unicos = _validos[[lat_hosp, lon_hosp]].drop_duplicates()
_n_hosp = len(_hospitais_unicos)

for _h_i, (_, _hrow) in enumerate(_hospitais_unicos.iterrows(), 1):
    _lat_h, _lon_h = _hrow[lat_hosp], _hrow[lon_hosp]
    _subset = _validos[(_validos[lat_hosp] == _lat_h) & (_validos[lon_hosp] == _lon_h)]
    print(f"   Hospital {_h_i}/{_n_hosp}: {len(_subset)} freguesia(s)...")
    _freg_rows = list(zip(_subset.index, _subset[lat_freq], _subset[lon_freq]))
    for _b in range(0, len(_freg_rows), OSRM_BATCH_SIZE):
        _lote = _freg_rows[_b:_b + OSRM_BATCH_SIZE]
        _res = osrm_table_batch(_lote, _lat_h, _lon_h)
        for _idx, (_dist_km, _tempo_min) in _res.items():
            df_freg.loc[_idx, 'distancia_km'] = _dist_km
            df_freg.loc[_idx, 'tempo_min'] = _tempo_min

_osrm_save_cache()

tempo = df_freg[['Distrito', 'Concelho', 'Freguesia', 'ULS', 'tempo_min']].copy()
print(f"   ✅ Tempo de acesso calculado: {len(tempo)} freguesias")

# Fatores COVID por faixa etária (2021 vs 2025)
var21 = pd.read_excel(F_VARIACAO, sheet_name='2021', header=None)
variacao = var21.iloc[18:25, [0, 1]].copy()
variacao.columns = ['Faixa', 'Variacao_2021_vs_2025']
variacao['Faixa'] = variacao['Faixa'].str.strip()
variacao['Multiplicador_Covid'] = 1 / (1 - variacao['Variacao_2021_vs_2025'])
variacao = variacao.set_index('Faixa')
print("\n🦠 Fatores de agravamento COVID por faixa etária:")
print(variacao.round(3))

# ═══════════════════════════════════════════════════════════════════════════
# 2. CENÁRIO: aplicar o choque COVID aos internamentos "código 22"
# ═══════════════════════════════════════════════════════════════════════════
print("\n🔄 Calculando cenário COVID (Antes/Depois)...")

df_covid = df_covid.merge(variacao[['Multiplicador_Covid']], left_on='Faixa Etária',
                           right_index=True, how='left')
if df_covid['Multiplicador_Covid'].isna().any():
    faltam = df_covid[df_covid['Multiplicador_Covid'].isna()]['Faixa Etária'].unique()
    print(f"   ⚠️  Faixas sem fator COVID (ficam sem alteração): {faltam}")
    df_covid['Multiplicador_Covid'] = df_covid['Multiplicador_Covid'].fillna(1.0)

df_covid['Internamentos_Covid_Depois'] = df_covid['Internamentos_Covid_Antes'] * df_covid['Multiplicador_Covid']
df_covid['Dias_Covid_Depois'] = df_covid['Internamentos_Covid_Depois'] * df_covid['Demora_Media_Covid']
df_covid['Camas_Covid_Depois'] = df_covid['Dias_Covid_Depois'] / 365
df_covid['Delta_Camas_Covid'] = df_covid['Camas_Covid_Depois'] - df_covid['Camas_Covid_Antes']
df_covid['Delta_Internamentos_Covid'] = df_covid['Internamentos_Covid_Depois'] - df_covid['Internamentos_Covid_Antes']

covid_delta = df_covid[KEYS + [
    'Internamentos_Covid_Antes', 'Internamentos_Covid_Depois', 'Delta_Internamentos_Covid',
    'Camas_Covid_Antes', 'Camas_Covid_Depois', 'Delta_Camas_Covid'
]]

# "Depois" = "Antes" + delta do código 22 (as restantes doenças não mudam)
proc = df_agg.merge(covid_delta, on=KEYS, how='left')
proc['Internamentos_Depois'] = proc['Internamentos_Antes'] + proc['Delta_Internamentos_Covid']
proc['Camas_Depois'] = proc['Camas_Antes'] + proc['Delta_Camas_Covid']
proc['Dias_Depois'] = proc['Camas_Depois'] * 365

print(f"   Internamentos nacionais ANTES  (todas as doenças): {round(proc['Internamentos_Antes'].sum()):,}")
print(f"   Internamentos nacionais DEPOIS (cenário COVID):    {round(proc['Internamentos_Depois'].sum()):,}")
print(f"   Camas necessárias nacionais ANTES:  {round(proc['Camas_Antes'].sum(), 1)}")
print(f"   Camas necessárias nacionais DEPOIS: {round(proc['Camas_Depois'].sum(), 1)}")

# ═══════════════════════════════════════════════════════════════════════════
# 3. AGREGAÇÃO POR FREGUESIA E POR ULS
# ═══════════════════════════════════════════════════════════════════════════
print("\n🔄 Agregando por freguesia e por ULS...")

freg_tot = proc.groupby(['Distrito', 'Concelho', 'Freguesia'], as_index=False).agg(
    Pop_Total=('Pop_Faixa', 'sum'),
    Internamentos_Antes=('Internamentos_Antes', 'sum'),
    Internamentos_Depois=('Internamentos_Depois', 'sum'),
    Internamentos_Covid_Antes=('Internamentos_Covid_Antes', 'sum'),
    Internamentos_Covid_Depois=('Internamentos_Covid_Depois', 'sum'),
    Camas_Necessarias_Antes=('Camas_Antes', 'sum'),
    Camas_Necessarias_Depois=('Camas_Depois', 'sum'),
)
freg_tot = freg_tot.merge(tempo, on=['Distrito', 'Concelho', 'Freguesia'], how='left')
freg_tot['Cobertura_%'] = np.where(freg_tot['tempo_min'] <= 60, 100.0, 0.0)

freg_tot['ULS_norm'] = freg_tot['ULS'].apply(normalize_uls_final)
df_oferta['ULS_norm'] = df_oferta['ULS'].apply(normalize_uls_final)
nome_uls_bonito = freg_tot.dropna(subset=['ULS']).drop_duplicates('ULS_norm').set_index('ULS_norm')['ULS']

uls_tot = freg_tot.groupby('ULS_norm', as_index=False).agg(
    Pop_ULS=('Pop_Total', 'sum'),
    Camas_Necessarias_Antes=('Camas_Necessarias_Antes', 'sum'),
    Camas_Necessarias_Depois=('Camas_Necessarias_Depois', 'sum'),
)
camas_col = 'Camas Totais (média 2025)'
oferta_uls = df_oferta.groupby('ULS_norm', as_index=False)[camas_col].sum().rename(columns={camas_col: 'Capacidade'})
uls_tot = uls_tot.merge(oferta_uls, on='ULS_norm', how='left')
uls_tot['ULS'] = uls_tot['ULS_norm'].map(nome_uls_bonito)
uls_tot['ULS'] = uls_tot['ULS'].fillna(uls_tot['ULS_norm'].str.title())

uls_tot['Cobertura_%_Antes'] = pd.Series(np.where(
    uls_tot['Camas_Necessarias_Antes'] > 0,
    (uls_tot['Capacidade'] / uls_tot['Camas_Necessarias_Antes'] * 100).clip(upper=100), np.nan
), index=uls_tot.index).round(1)
uls_tot['Cobertura_%_Depois'] = pd.Series(np.where(
    uls_tot['Camas_Necessarias_Depois'] > 0,
    (uls_tot['Capacidade'] / uls_tot['Camas_Necessarias_Depois'] * 100).clip(upper=100), np.nan
), index=uls_tot.index).round(1)
uls_tot['Camas_Em_Falta_Antes'] = (uls_tot['Camas_Necessarias_Antes'] - uls_tot['Capacidade']).clip(lower=0).round(1)
uls_tot['Camas_Em_Falta_Depois'] = (uls_tot['Camas_Necessarias_Depois'] - uls_tot['Capacidade']).clip(lower=0).round(1)
uls_tot['Delta_Camas_Em_Falta'] = (uls_tot['Camas_Em_Falta_Depois'] - uls_tot['Camas_Em_Falta_Antes']).round(1)
uls_tot['Camas_Necessarias_Antes'] = uls_tot['Camas_Necessarias_Antes'].round(2)
uls_tot['Camas_Necessarias_Depois'] = uls_tot['Camas_Necessarias_Depois'].round(2)
uls_tot['Pop_ULS'] = uls_tot['Pop_ULS'].round(0)
uls_tot = uls_tot.sort_values('Delta_Camas_Em_Falta', ascending=False).reset_index(drop=True)

n_sem_oferta = int(uls_tot['Capacidade'].isna().sum())
if n_sem_oferta:
    print(f"   ⚠️  {n_sem_oferta} ULS sem correspondência na Oferta -> Cobertura ficará em branco")

# Resumo nacional
cam_antes = uls_tot['Camas_Necessarias_Antes'].sum()
cam_depois = uls_tot['Camas_Necessarias_Depois'].sum()
cap_total = uls_tot['Capacidade'].sum()

resumo = {
    'Camas necessárias nacionais (2025)': round(cam_antes, 1),
    'Camas necessárias nacionais (2021)': round(cam_depois, 1),
    'Acréscimo de camas necessárias': round(cam_depois - cam_antes, 1),
    'Cobertura de camas média nacional (2025) %': round(uls_tot['Cobertura_%_Antes'].mean(), 1),
    'Cobertura de camas média nacional (2021) %': round(uls_tot['Cobertura_%_Depois'].mean(), 1),
}
print("\n📊 RESUMO NACIONAL:")
for k, v in resumo.items():
    print(f"   {k}: {v}")

freg_tot = freg_tot.sort_values(['Distrito', 'Concelho', 'Freguesia']).reset_index(drop=True)

# ═══════════════════════════════════════════════════════════════════════════
# 4. EXCEL — DASHBOARD (estilo Modelo3)
# ═══════════════════════════════════════════════════════════════════════════
print("\n📊 Criando Dashboard_Impacto_COVID_Internamentos.xlsx...")

wb = Workbook()
wb.remove(wb.active)

# --- 1. Resumo -------------------------------------------------------------
ws1 = wb.create_sheet('1. Resumo')
ws1['B2'] = 'MODELO — IMPACTO DE UM CHOQUE COVID (NÍVEL 2021) NOS INTERNAMENTOS E CAMAS'
ws1['B2'].font = Font(bold=True, size=13, color="C00000")
labels_order = [
    'Camas necessárias nacionais (2025)', 'Camas necessárias nacionais (2021)',
    'Acréscimo de camas necessárias', None,
    'Cobertura de camas média nacional (2025) %', 'Cobertura de camas média nacional (2021) %',
]
row = 5
for lab in labels_order:
    if lab is None:
        row += 1
        continue
    ws1.cell(row=row, column=2, value=lab)
    ws1.cell(row=row, column=3, value=resumo[lab])
    row += 1
row += 1
ws1.cell(row=row, column=2, value='Nota')
ws1.cell(row=row, column=3, value=(
    "O tempo de trajeto não se altera (é geografia pura); a oferta de camas (capacidade) mantém-se constante. "
    "O que piora é a PROCURA de camas: aplica-se, a cada freguesia e faixa etária, o fator de agravamento observado "
    "em 2021 (pico COVID) vs 2025 nos internamentos por 'Códigos para fins especiais', mantendo inalterados os "
    "internamentos das restantes doenças."
))
ws1.cell(row=row, column=3).alignment = Alignment(wrap_text=True, vertical='top')
ws1.column_dimensions['B'].width = 46
ws1.column_dimensions['C'].width = 60

# --- 2. Cobertura_ULS --------------------------------------------------------
ws2 = wb.create_sheet('2. Cobertura_ULS')
def rotulo(h):
    """Traduz nomes internos de colunas (_Antes/_Depois) para rótulos amigáveis (2025/2021)."""
    return h.replace('_Antes', '_2025').replace('_Depois', '_2021').replace('Antes', '2025').replace('Depois', '2021')


cols2 = ['ULS', 'Camas_Necessarias_Antes', 'Camas_Necessarias_Depois', 'Capacidade',
         'Cobertura_%_Antes', 'Cobertura_%_Depois', 'Camas_Em_Falta_Antes', 'Camas_Em_Falta_Depois',
         'Delta_Camas_Em_Falta']
for j, h in enumerate(cols2, 1):
    ws2.cell(row=1, column=j, value=rotulo(h))
for i, (_, r) in enumerate(uls_tot.iterrows(), 2):
    for j, h in enumerate(cols2, 1):
        v = r[h]
        ws2.cell(row=i, column=j, value=(None if pd.isna(v) else v))
    ws2.cell(row=i, column=11, value=f"=F{i}-ROW()/100000000")  # K: chave única p/ SMALL
ws2.cell(row=1, column=11, value='Rank_Cobertura_2021')
style_header(ws2, len(cols2))
ws2.cell(row=1, column=11).font = header_font
ws2.cell(row=1, column=11).fill = header_fill_blue
ws2.freeze_panes = 'A2'
autofit(ws2, 11)
n_uls = len(uls_tot)

# --- 3. Por_Freguesia --------------------------------------------------------
ws3 = wb.create_sheet('3. Por_Freguesia')
ws3['A1'] = ("Cobertura_% = 100 se tempo_min<=60, senão 0 (geografia pura — não muda com o choque COVID). "
             "Internamentos_Covid_[2025/2021] = só a categoria 'Códigos para fins especiais': 2025 é o nível atual, "
             "2021 é a simulação ao nível do pico da pandemia, escalada pelo fator daquela faixa etária. "
             "Camas_Necessarias_[2025/2021] = total nacional na freguesia (todas as doenças), incluindo o efeito "
             "do choque COVID em 2021.")
ws3['A1'].font = Font(italic=True, size=9, color="666666")
cols3 = ['Distrito', 'Concelho', 'Freguesia', 'ULS', 'tempo_min', 'Pop_Total',
         'Internamentos_Covid_Antes', 'Internamentos_Covid_Depois', 'Cobertura_%',
         'Camas_Necessarias_Antes', 'Camas_Necessarias_Depois']
for j, h in enumerate(cols3, 1):
    ws3.cell(row=2, column=j, value=rotulo(h))
for i, (_, r) in enumerate(freg_tot.iterrows(), 3):
    for j, h in enumerate(cols3, 1):
        v = r[h]
        if h == 'tempo_min':
            v = round(v, 1)
        elif h in ('Internamentos_Covid_Antes', 'Internamentos_Covid_Depois'):
            v = round(v, 2)
        elif h in ('Camas_Necessarias_Antes', 'Camas_Necessarias_Depois'):
            v = round(v, 4)
        ws3.cell(row=i, column=j, value=v)
style_header(ws3, len(cols3), row=2)
ws3.freeze_panes = 'A3'
autofit(ws3, len(cols3))
n_freg = len(freg_tot)
last_row = 2 + n_freg

# --- 0. Dashboard (interativo, criado por último porque referencia as outras) ---
ws0 = wb.create_sheet('0. Dashboard', 0)
ws0.sheet_view.showGridLines = False
ws0.column_dimensions['A'].width = 2
for col in 'BCDEFGHIJ':
    ws0.column_dimensions[col].width = 15

distritos = sorted(freg_tot['Distrito'].unique().tolist())
uls_lista = uls_tot.sort_values('ULS')['ULS'].tolist()

ws0['N1'] = 'lista_distritos'
for i, d in enumerate(distritos, 2):
    ws0.cell(row=i, column=14, value=d)
ws0['O1'] = 'lista_uls'
for i, u in enumerate(uls_lista, 2):
    ws0.cell(row=i, column=15, value=u)

titulo_font = Font(bold=True, size=16, color="C00000")
sub_font = Font(size=10, color="666666")
sec_font = Font(bold=True, size=12, color="C00000")
kpi_lab_font = Font(bold=True, size=9, color="666666")
kpi_val_font = Font(bold=True, size=15, color="C00000")
box_fill = PatternFill(start_color="F7F7F7", end_color="F7F7F7", fill_type="solid")
thin_bottom = Border(bottom=Side(style='thin', color="C00000"))

# ── Título ───────────────────────────────────────────────────────────────
ws0['B2'] = 'DASHBOARD INTERATIVO — IMPACTO DE UM CHOQUE COVID (NÍVEL 2021) NOS INTERNAMENTOS'
ws0['B2'].font = titulo_font
ws0.merge_cells('B2:J2')
ws0['B3'] = (f"O que acontece à procura e à cobertura hospitalar do país se os internamentos por "
             f"'Códigos para fins especiais' voltarem ao nível do pico de 2021, por faixa etária? — "
             f"{n_freg} freguesias, {n_uls} ULS, mesma rede geográfica e mesma oferta de camas")
ws0['B3'].font = sub_font
ws0.merge_cells('B3:J3')

# ── KPIs nacionais ───────────────────────────────────────────────────────
kpi_headers = ['CAMAS NECESSÁRIAS (2025)', 'CAMAS NECESSÁRIAS (2021)', 'ACRÉSCIMO DE CAMAS',
               'COBERTURA MÉDIA (2025)', 'COBERTURA MÉDIA (2021)', 'QUEDA DE COBERTURA']
for j, h in enumerate(kpi_headers):
    cell = ws0.cell(row=5, column=2 + j, value=h)
    cell.font = kpi_lab_font
    cell.alignment = Alignment(horizontal='center', wrap_text=True)
    cell.fill = box_fill

kpi_formulas = [
    "='1. Resumo'!C5", "='1. Resumo'!C6", "='1. Resumo'!C7",
    "='1. Resumo'!C9/100", "='1. Resumo'!C10/100", "=('1. Resumo'!C10-'1. Resumo'!C9)/100",
]
for j, f in enumerate(kpi_formulas):
    cell = ws0.cell(row=6, column=2 + j, value=f)
    cell.font = kpi_val_font
    cell.alignment = Alignment(horizontal='center')
    cell.number_format = '0.0%' if ('COBERTURA' in kpi_headers[j] or 'QUEDA' in kpi_headers[j]) else '#,##0.0'
    cell.fill = box_fill

ws0['B7'] = ('="Nota: o tempo de trajeto NÃO muda (é geografia pura) e a oferta de camas mantém-se — "'
             '&"o que piora é a procura de camas por internamentos ao nível COVID de 2021, por faixa etária."')
ws0['B7'].font = Font(italic=True, size=9, color="666666")
ws0.merge_cells('B7:J7')

# ── Filtro interativo por DISTRITO ──────────────────────────────────────
row = 9
ws0.cell(row=row, column=2, value='🔎  FILTRO INTERATIVO — DISTRITO').font = sec_font
ws0.cell(row=row, column=2).border = thin_bottom
ws0.merge_cells(start_row=row, start_column=2, end_row=row, end_column=9)
for c in range(2, 10):
    ws0.cell(row=row, column=c).border = thin_bottom
row += 1
row_dist_header = row
filt_headers_dist = ['Nº Freguesias', 'Internamentos (2025)', 'Internamentos (2021)',
                      'Camas Necess. (2025)', 'Camas Necess. (2021)', 'Acréscimo de Camas']
for j, h in enumerate(filt_headers_dist):
    cell = ws0.cell(row=row, column=4 + j, value=h)
    cell.font = kpi_lab_font
    cell.alignment = Alignment(horizontal='center', wrap_text=True)
row += 1
row_dist_value = row
ws0.cell(row=row, column=2, value='Escolher Distrito:').font = Font(bold=True)
c_dist = ws0.cell(row=row, column=3, value=distritos[0] if distritos else '')
c_dist.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
c_dist.font = Font(bold=True, color="C00000")
for c in range(4, 10):
    ws0.cell(row=row, column=c).fill = box_fill

dv_dist = DataValidation(type="list", formula1="=lista_distritos", allow_blank=False)
ws0.add_data_validation(dv_dist)
dv_dist.add(f'C{row_dist_value}')
wb.defined_names['lista_distritos'] = DefinedName(
    'lista_distritos', attr_text=f"'0. Dashboard'!$N$2:$N${1 + len(distritos)}")

ws0.cell(row=row_dist_value, column=4, value=f"=COUNTIF('3. Por_Freguesia'!A3:A{last_row},$C${row_dist_value})")
ws0.cell(row=row_dist_value, column=5, value=f"=ROUND(SUMIF('3. Por_Freguesia'!A3:A{last_row},$C${row_dist_value},'3. Por_Freguesia'!G3:G{last_row}),0)")
ws0.cell(row=row_dist_value, column=6, value=f"=ROUND(SUMIF('3. Por_Freguesia'!A3:A{last_row},$C${row_dist_value},'3. Por_Freguesia'!H3:H{last_row}),0)")
ws0.cell(row=row_dist_value, column=7, value=f"=ROUND(SUMIF('3. Por_Freguesia'!A3:A{last_row},$C${row_dist_value},'3. Por_Freguesia'!J3:J{last_row}),1)")
ws0.cell(row=row_dist_value, column=8, value=f"=ROUND(SUMIF('3. Por_Freguesia'!A3:A{last_row},$C${row_dist_value},'3. Por_Freguesia'!K3:K{last_row}),1)")
ws0.cell(row=row_dist_value, column=9, value=(
    f"=ROUND(SUMIF('3. Por_Freguesia'!A3:A{last_row},$C${row_dist_value},'3. Por_Freguesia'!K3:K{last_row})"
    f"-SUMIF('3. Por_Freguesia'!A3:A{last_row},$C${row_dist_value},'3. Por_Freguesia'!J3:J{last_row}),1)"))

# ── Filtro interativo por ULS ────────────────────────────────────────────
row += 2
ws0.cell(row=row, column=2, value='🏥  FILTRO INTERATIVO — ULS').font = sec_font
ws0.cell(row=row, column=2).border = thin_bottom
ws0.merge_cells(start_row=row, start_column=2, end_row=row, end_column=9)
for c in range(2, 10):
    ws0.cell(row=row, column=c).border = thin_bottom
row += 1
row_uls_header = row
filt_headers_uls = ['Camas Necess. (2025)', 'Camas Necess. (2021)', 'Capacidade (camas)',
                     'Cobertura % (2025)', 'Cobertura % (2021)', 'Camas em Falta (2021)']
for j, h in enumerate(filt_headers_uls):
    cell = ws0.cell(row=row, column=4 + j, value=h)
    cell.font = kpi_lab_font
    cell.alignment = Alignment(horizontal='center', wrap_text=True)
row += 1
row_uls_value = row
ws0.cell(row=row, column=2, value='Escolher ULS:').font = Font(bold=True)
c_uls = ws0.cell(row=row, column=3, value=uls_lista[0] if uls_lista else '')
c_uls.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
c_uls.font = Font(bold=True, color="C00000")
for c in range(4, 10):
    ws0.cell(row=row, column=c).fill = box_fill

dv_uls = DataValidation(type="list", formula1="=lista_uls", allow_blank=False)
ws0.add_data_validation(dv_uls)
dv_uls.add(f'C{row_uls_value}')
wb.defined_names['lista_uls'] = DefinedName(
    'lista_uls', attr_text=f"'0. Dashboard'!$O$2:$O${1 + len(uls_lista)}")

# '2. Cobertura_ULS' colunas: A=ULS, B=Camas_2025, C=Camas_2021, D=Capacidade,
# E=Cobertura%_2025, F=Cobertura%_2021, G=Falta_2025, H=Falta_2021
for col_i, src_col in enumerate(['B', 'C', 'D', 'E', 'F', 'H']):
    ws0.cell(row=row_uls_value, column=4 + col_i, value=(
        f"=IFERROR(INDEX('2. Cobertura_ULS'!${src_col}$2:${src_col}${1 + n_uls},"
        f"MATCH($C${row_uls_value},'2. Cobertura_ULS'!$A$2:$A${1 + n_uls},0)),\"—\")"))
# Cobertura % já vem em escala 0-100 na folha de origem, por isso só se acrescenta o símbolo "%"
ws0.cell(row=row_uls_value, column=7).number_format = '0.0"%"'
ws0.cell(row=row_uls_value, column=8).number_format = '0.0"%"'

# ── Tabela por distrito + Top 15 ULS pior agravamento ───────────────────
row += 3
row_section2 = row
ws0.cell(row=row, column=2, value='📊  ANÁLISE POR DISTRITO — CAMAS NECESSÁRIAS E INTERNAMENTOS COVID').font = sec_font
ws0.cell(row=row, column=2).border = thin_bottom
ws0.merge_cells(start_row=row, start_column=2, end_row=row, end_column=6)
for c in range(2, 7):
    ws0.cell(row=row, column=c).border = thin_bottom
ws0.cell(row=row, column=8, value='🏥  TOP 15 ULS — MAIOR AGRAVAMENTO DO DÉFICE').font = sec_font
ws0.cell(row=row, column=8).border = thin_bottom
ws0.merge_cells(start_row=row, start_column=8, end_row=row, end_column=9)
for c in (8, 9):
    ws0.cell(row=row, column=c).border = thin_bottom

row_tbl_header = row + 1
dist_headers = ['Distrito', 'Camas Necess. 2025', 'Camas Necess. 2021', 'Internamentos 2025', 'Internamentos 2021']
for j, h in enumerate(dist_headers):
    ws0.cell(row=row_tbl_header, column=2 + j, value=h).font = kpi_lab_font
ws0.cell(row=row_tbl_header, column=8, value='ULS').font = kpi_lab_font
ws0.cell(row=row_tbl_header, column=9, value='Agravamento (camas)').font = kpi_lab_font

row_tbl_start = row_tbl_header + 1
for i, d in enumerate(distritos, row_tbl_start):
    ws0.cell(row=i, column=2, value=d)
    ws0.cell(row=i, column=3, value=f"=ROUND(SUMIF('3. Por_Freguesia'!$A$3:$A${last_row},$B{i},'3. Por_Freguesia'!$J$3:$J${last_row}),1)")
    ws0.cell(row=i, column=4, value=f"=ROUND(SUMIF('3. Por_Freguesia'!$A$3:$A${last_row},$B{i},'3. Por_Freguesia'!$K$3:$K${last_row}),1)")
    ws0.cell(row=i, column=5, value=f"=ROUND(SUMIF('3. Por_Freguesia'!$A$3:$A${last_row},$B{i},'3. Por_Freguesia'!$G$3:$G${last_row}),0)")
    ws0.cell(row=i, column=6, value=f"=ROUND(SUMIF('3. Por_Freguesia'!$A$3:$A${last_row},$B{i},'3. Por_Freguesia'!$H$3:$H${last_row}),0)")

n_top = min(15, n_uls)
for k in range(n_top):
    r = row_tbl_start + k
    ws0.cell(row=r, column=8, value=f"='2. Cobertura_ULS'!A{2 + k}")
    ws0.cell(row=r, column=9, value=f"='2. Cobertura_ULS'!I{2 + k}")

row_tbl_end = row_tbl_start + len(distritos) - 1

# ── Top 15 piores coberturas em 2021 ─────────────────────────────────────
row_top = row_tbl_end + 3
ws0.cell(row=row_top, column=2, value='📉  TOP 15 ULS COM PIOR COBERTURA EM 2021 (CENÁRIO COVID)').font = sec_font
ws0.cell(row=row_top, column=2).border = thin_bottom
ws0.merge_cells(start_row=row_top, start_column=2, end_row=row_top, end_column=6)
for c in range(2, 7):
    ws0.cell(row=row_top, column=c).border = thin_bottom
h_row = row_top + 1
worst_headers = ['ULS', 'Cobertura % 2025', 'Cobertura % 2021', 'Camas em Falta 2025', 'Camas em Falta 2021']
for j, h in enumerate(worst_headers):
    ws0.cell(row=h_row, column=2 + j, value=h).font = kpi_lab_font

n_worst = min(15, n_uls)
for k in range(n_worst):
    r = h_row + 1 + k
    ws0.cell(row=r, column=12, value=f"=SMALL('2. Cobertura_ULS'!$K$2:$K${1 + n_uls},{k + 1})")
    ws0.cell(row=r, column=2, value=f"=INDEX('2. Cobertura_ULS'!$A$2:$A${1 + n_uls},MATCH($L{r},'2. Cobertura_ULS'!$K$2:$K${1 + n_uls},0))")
    ws0.cell(row=r, column=3, value=f"=INDEX('2. Cobertura_ULS'!$E$2:$E${1 + n_uls},MATCH($L{r},'2. Cobertura_ULS'!$K$2:$K${1 + n_uls},0))")
    ws0.cell(row=r, column=4, value=f"=INDEX('2. Cobertura_ULS'!$F$2:$F${1 + n_uls},MATCH($L{r},'2. Cobertura_ULS'!$K$2:$K${1 + n_uls},0))")
    ws0.cell(row=r, column=5, value=f"=INDEX('2. Cobertura_ULS'!$G$2:$G${1 + n_uls},MATCH($L{r},'2. Cobertura_ULS'!$K$2:$K${1 + n_uls},0))")
    ws0.cell(row=r, column=6, value=f"=INDEX('2. Cobertura_ULS'!$H$2:$H${1 + n_uls},MATCH($L{r},'2. Cobertura_ULS'!$K$2:$K${1 + n_uls},0))")

# ── Gráficos ──────────────────────────────────────────────────────────────
chart_row = h_row + n_worst + 3
chart1 = BarChart()
chart1.type = "col"
chart1.title = "Camas Necessárias por Distrito — 2025 vs 2021"
chart1.y_axis.title = "Camas"
chart1.style = 10
data1 = Reference(ws0, min_col=3, max_col=4, min_row=row_tbl_header, max_row=row_tbl_end)
cats1 = Reference(ws0, min_col=2, min_row=row_tbl_start, max_row=row_tbl_end)
chart1.add_data(data1, titles_from_data=True)
chart1.set_categories(cats1)
chart1.height, chart1.width = 9, 18
chart1.legend.position = 'b'
ws0.add_chart(chart1, f'B{chart_row}')

chart2 = BarChart()
chart2.type = "bar"
chart2.title = f"Top {n_top} ULS — Agravamento do Défice de Camas"
chart2.x_axis.title = "Camas em falta a mais"
chart2.style = 12
data2 = Reference(ws0, min_col=9, min_row=row_tbl_header, max_row=row_tbl_start + n_top - 1)
cats2 = Reference(ws0, min_col=8, min_row=row_tbl_start, max_row=row_tbl_start + n_top - 1)
chart2.add_data(data2, titles_from_data=True)
chart2.set_categories(cats2)
chart2.height, chart2.width = 9, 18
chart2.legend.position = 'b'
ws0.add_chart(chart2, f'H{chart_row}')

wb.save(OUT_DASHBOARD)
print(f"   ✅ Guardado: {OUT_DASHBOARD}")

# ═══════════════════════════════════════════════════════════════════════════
# 5. EXCEL — PROCURA (cenário COVID)
# ═══════════════════════════════════════════════════════════════════════════
print("\n📊 Criando Excel_Procura_Portugal_Cenario_COVID.xlsx...")

wbp = Workbook()
wbp.remove(wbp.active)

wsp0 = wbp.create_sheet('0. Fatores COVID por Faixa')
wsp0['A1'] = ("Fator = multiplicador aplicado aos internamentos da categoria 'Códigos para fins especiais' (COVID) "
              "de cada faixa etária, para simular um choque ao nível observado em 2021. Fonte: "
              "'variação_de_internamentos.xlsx' (2021 vs 2025). Multiplicador = 1 / (1 - Variação); "
              "Variação = (Internamentos_2021 - Internamentos_2025) / Internamentos_2021.")
wsp0['A1'].font = Font(italic=True, size=9, color="666666")
wsp0.merge_cells('A1:C1')
wsp0.row_dimensions[1].height = 45
wsp0['A1'].alignment = Alignment(wrap_text=True, vertical='top')
cols_p0 = ['Faixa', 'Variacao_2021_vs_2025', 'Multiplicador_Covid']
for j, h in enumerate(cols_p0, 1):
    wsp0.cell(row=2, column=j, value=h)
for i, (faixa, r) in enumerate(variacao.iterrows(), 3):
    wsp0.cell(row=i, column=1, value=faixa)
    wsp0.cell(row=i, column=2, value=round(r['Variacao_2021_vs_2025'], 4))
    wsp0.cell(row=i, column=3, value=round(r['Multiplicador_Covid'], 3))
style_header(wsp0, 3, row=2)
autofit(wsp0, 3)

wsp1 = wbp.create_sheet('1. Procura por Freg x Faixa')
wsp1['A1'] = ("Internamentos_2025/2021 = todas as doenças (só 'Códigos para fins especiais' é escalada pelo "
              "fator COVID da faixa; as restantes doenças mantêm-se). Camas_Necessarias = Dias Necessários / 365.")
wsp1['A1'].font = Font(italic=True, size=9, color="666666")
cols_p1 = ['Distrito', 'Concelho', 'Freguesia', 'Faixa Etária', 'Pop_Faixa',
           'Internamentos_Antes', 'Internamentos_Depois', 'Camas_Antes', 'Camas_Depois']
for j, h in enumerate(cols_p1, 1):
    wsp1.cell(row=2, column=j, value=rotulo(h))
proc_out = proc[cols_p1].copy()
for i, (_, r) in enumerate(proc_out.iterrows(), 3):
    for j, h in enumerate(cols_p1, 1):
        v = r[h]
        if h in ('Internamentos_Antes', 'Internamentos_Depois'):
            v = round(v, 2)
        elif h in ('Camas_Antes', 'Camas_Depois'):
            v = round(v, 4)
        wsp1.cell(row=i, column=j, value=v)
style_header(wsp1, len(cols_p1), row=2)
wsp1.freeze_panes = 'A3'
autofit(wsp1, len(cols_p1))

wsp2 = wbp.create_sheet('2. Procura Agregada Freguesia')
cols_p2 = ['Distrito', 'Concelho', 'Freguesia', 'ULS', 'Pop_Total',
           'Internamentos_Antes', 'Internamentos_Depois', 'Camas_Necessarias_Antes', 'Camas_Necessarias_Depois']
for j, h in enumerate(cols_p2, 1):
    wsp2.cell(row=1, column=j, value=rotulo(h))
for i, (_, r) in enumerate(freg_tot.iterrows(), 2):
    wsp2.cell(row=i, column=1, value=r['Distrito'])
    wsp2.cell(row=i, column=2, value=r['Concelho'])
    wsp2.cell(row=i, column=3, value=r['Freguesia'])
    wsp2.cell(row=i, column=4, value=r['ULS'])
    wsp2.cell(row=i, column=5, value=r['Pop_Total'])
    wsp2.cell(row=i, column=6, value=round(r['Internamentos_Antes'], 1))
    wsp2.cell(row=i, column=7, value=round(r['Internamentos_Depois'], 1))
    wsp2.cell(row=i, column=8, value=round(r['Camas_Necessarias_Antes'], 3))
    wsp2.cell(row=i, column=9, value=round(r['Camas_Necessarias_Depois'], 3))
style_header(wsp2, len(cols_p2))
wsp2.freeze_panes = 'A2'
autofit(wsp2, len(cols_p2))

wsp3 = wbp.create_sheet('3. Procura Agregada por ULS')
cols_p3 = ['ULS', 'Pop_ULS', 'Camas_Necessarias_Antes', 'Camas_Necessarias_Depois', 'Acrescimo_Camas', 'Acrescimo_%']
for j, h in enumerate(cols_p3, 1):
    wsp3.cell(row=1, column=j, value=rotulo(h))
uls_sorted = uls_tot.sort_values('Camas_Necessarias_Depois', ascending=False)
for i, (_, r) in enumerate(uls_sorted.iterrows(), 2):
    acresc = r['Camas_Necessarias_Depois'] - r['Camas_Necessarias_Antes']
    pct = (acresc / r['Camas_Necessarias_Antes'] * 100) if r['Camas_Necessarias_Antes'] else np.nan
    wsp3.cell(row=i, column=1, value=r['ULS'])
    wsp3.cell(row=i, column=2, value=r['Pop_ULS'])
    wsp3.cell(row=i, column=3, value=round(r['Camas_Necessarias_Antes'], 1))
    wsp3.cell(row=i, column=4, value=round(r['Camas_Necessarias_Depois'], 1))
    wsp3.cell(row=i, column=5, value=round(acresc, 1))
    wsp3.cell(row=i, column=6, value=(None if pd.isna(pct) else round(pct, 1)))
style_header(wsp3, len(cols_p3))
wsp3.freeze_panes = 'A2'
autofit(wsp3, len(cols_p3))

wbp.save(OUT_PROCURA)
print(f"   ✅ Guardado: {OUT_PROCURA}")

# ═══════════════════════════════════════════════════════════════════════════
# 6. EXCEL — OFERTA (constante no cenário; entregue como referência)
# ═══════════════════════════════════════════════════════════════════════════
print("\n📊 Criando Excel_Oferta_Portugal_Referencia.xlsx...")

wbo = Workbook()
wbo.remove(wbo.active)
wso = wbo.create_sheet('1. Oferta por ULS')
wso['A1'] = ("Capacidade Real = Camas Totais × Taxa de Ocupação (nº médio de camas efetivamente ocupadas por dia, "
             "em 2025). A oferta NÃO se altera no cenário COVID — este ficheiro é a mesma oferta-base usada como "
             "referência de 'Capacidade' no modelo de cobertura (coluna 'Camas Totais').")
wso['A1'].font = Font(italic=True, size=9, color="666666")
wso.merge_cells('A1:D1')
wso.row_dimensions[1].height = 30
wso['A1'].alignment = Alignment(wrap_text=True, vertical='top')
cols_o = ['ULS', 'Camas Totais (média 2025)', 'Taxa Ocupação (%)', 'Expectativa de capacidade utilizada']
for j, h in enumerate(cols_o, 1):
    wso.cell(row=2, column=j, value=h)
for i, (_, r) in enumerate(df_oferta.sort_values('ULS').iterrows(), 3):
    for j, h in enumerate(cols_o, 1):
        wso.cell(row=i, column=j, value=r[h])
style_header(wso, len(cols_o), row=2)
wso.freeze_panes = 'A3'
autofit(wso, len(cols_o))
wbo.save(OUT_OFERTA)
print(f"   ✅ Guardado: {OUT_OFERTA}")

# ═══════════════════════════════════════════════════════════════════════════
print("\n" + "=" * 70)
print("✅ CONCLUÍDO!")
print("=" * 70)