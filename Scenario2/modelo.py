#!/usr/bin/env python3
"""
MODELO 2 — LOCALIZAÇÃO ÓTIMA DE UM NOVO HOSPITAL NO ALGARVE (COM CAPACIDADE)
════════════════════════════════════════════════════════════════════════════
OBJETIVO
Hoje, todas as freguesias do distrito de Faro (região do Algarve) têm como
hospital de referência o Hospital de Faro. Este modelo pergunta: "se
construíssemos UM novo hospital com CAPACIDADE_NOVO_HOSPITAL camas algures
no Algarve, em que freguesia deveria ficar para minimizar o tempo de
trajeto ponderado pela população, tendo em conta que nem toda a procura
cabe no novo hospital?"

MÉTODO — Para cada uma das 67 freguesias candidatas, resolve-se um pequeno
problema de transportes capacitado (2 "hospitais": o Hospital de Faro
existente e o novo hospital candidato), igual em espírito ao Modelo 1:

    minimizar   Σ_i Σ_j  pop_i * tempo_ij * x_ij  +  penalização * overflow_j
    sujeito a   x_i,Faro + x_i,Novo = 1                 para cada freguesia i
                Σ_i demanda_i * x_i,Faro  ≤ capacidade_Faro  + overflow_Faro
                Σ_i demanda_i * x_i,Novo  ≤ CAPACIDADE_NOVO_HOSPITAL + overflow_Novo

demanda_i = camas necessárias estimadas da freguesia (Excel_Procura).
capacidade_Faro = camas atuais da ULS do Algarve (Excel_Oferta).
x_i,Faro / x_i,Novo = fração da freguesia i encaminhada para cada hospital
(o solver reparte a procura entre os dois quando a capacidade aperta).

Escolhe-se o candidato c* que minimiza o custo total (tempo + penalização
por défice de capacidade). Reporta-se o ranking completo das 67 hipóteses.

OUTPUT: Modelo2_Localizacao_Algarve.xlsx com:
  - Resumo executivo (melhor localização, capacidades, défices)
  - Ranking de todas as freguesias candidatas
  - Impacto por freguesia servida (tempo antes/depois, hospital atribuído)
"""
import pandas as pd
import numpy as np
import os
import json
import time
import hashlib
import pulp
import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ── PARÂMETROS DO MODELO ────────────────────────────────────────────────────
VELOCIDADE_KMH = 60   # só usado no FALLBACK, se o OSRM estiver indisponível
TEMPO_FIXO_MIN = 15
DISTRITO_ALVO  = 'Faro'   # distrito de Faro = região estatística do Algarve
CAPACIDADE_NOVO_HOSPITAL = 742     # camas do novo hospital (assumido pelo utilizador)
PENALIZACAO_OVERFLOW = 1_000_000   # penalização por cada unidade de procura não coberta

# ── PARÂMETROS DO OSRM (distâncias/tempos reais de estrada, não euclidianos) ─
OSRM_BASE_URL = "http://router.project-osrm.org"   # servidor público de demonstração
OSRM_CHUNK_SIZE = 90       # nº de pontos por bloco em cada pedido /table (limite do servidor)
OSRM_TIMEOUT = 30
OSRM_MAX_RETRIES = 3
OSRM_CACHE_PATH = os.path.join(SCRIPT_DIR, '.osrm_cache_algarve.json')

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


def haversine_matrix(lat1, lon1, lat2, lon2):
    lat1r, lon1r = np.radians(lat1)[:, None], np.radians(lon1)[:, None]
    lat2r, lon2r = np.radians(lat2)[None, :], np.radians(lon2)[None, :]
    dlat = lat2r - lat1r
    dlon = lon2r - lon1r
    a = np.sin(dlat / 2) ** 2 + np.cos(lat1r) * np.cos(lat2r) * np.sin(dlon / 2) ** 2
    c = 2 * np.arcsin(np.sqrt(a))
    return 6371 * c


# ── CACHE EM DISCO PARA O OSRM (evita repetir pedidos ao servidor) ──────────
if os.path.exists(OSRM_CACHE_PATH):
    try:
        with open(OSRM_CACHE_PATH) as f:
            _osrm_cache = json.load(f)
    except Exception:
        _osrm_cache = {}
else:
    _osrm_cache = {}


def _osrm_save_cache():
    try:
        with open(OSRM_CACHE_PATH, 'w') as f:
            json.dump(_osrm_cache, f)
    except Exception:
        pass


def _osrm_cache_key(lats, lons):
    raw = "|".join(f"{a:.5f},{b:.5f}" for a, b in zip(lats, lons))
    return hashlib.md5(raw.encode()).hexdigest()


def osrm_table_km_min(lats, lons):

    n = len(lats)
    key = _osrm_cache_key(lats, lons)
    if key in _osrm_cache:
        print("   (OSRM: distâncias lidas da cache local)")
        cached = _osrm_cache[key]
        return np.array(cached['tempo_min']), np.array(cached['dist_km'])

    tempo_min = np.full((n, n), np.nan)
    dist_km = np.full((n, n), np.nan)
    blocos = [list(range(i, min(i + OSRM_CHUNK_SIZE, n))) for i in range(0, n, OSRM_CHUNK_SIZE)]

    try:
        for src_bloco in blocos:
            for dst_bloco in blocos:
                idx_usados = sorted(set(src_bloco) | set(dst_bloco))
                pos = {orig: k for k, orig in enumerate(idx_usados)}
                coord_str = ";".join(f"{lons[i]:.6f},{lats[i]:.6f}" for i in idx_usados)
                src_str = ";".join(str(pos[i]) for i in src_bloco)
                dst_str = ";".join(str(pos[i]) for i in dst_bloco)
                url = (f"{OSRM_BASE_URL}/table/v1/driving/{coord_str}"
                       f"?sources={src_str}&destinations={dst_str}&annotations=duration,distance")

                resp = None
                for tentativa in range(OSRM_MAX_RETRIES):
                    try:
                        resp = requests.get(url, timeout=OSRM_TIMEOUT)
                        if resp.status_code == 200:
                            break
                    except requests.exceptions.RequestException:
                        resp = None
                    time.sleep(1 + tentativa)
                if resp is None or resp.status_code != 200:
                    raise RuntimeError(f"status {resp.status_code if resp is not None else 'sem resposta'}")

                data = resp.json()
                if data.get('code') != 'Ok':
                    raise RuntimeError(f"OSRM devolveu code={data.get('code')}")
                durations = data['durations']
                distances = data['distances']
                for a, si in enumerate(src_bloco):
                    for b, di in enumerate(dst_bloco):
                        d = durations[a][b]
                        m = distances[a][b]
                        tempo_min[si, di] = (d / 60) if d is not None else np.nan
                        dist_km[si, di] = (m / 1000) if m is not None else np.nan
        print(f"   ✅ OSRM: matriz {n}x{n} de distâncias/tempos reais de estrada obtida com sucesso.")
    except Exception as e:
        print(f"   ⚠️  OSRM indisponível ({e}) — a usar fallback (haversine + {VELOCIDADE_KMH} km/h).")

    # preenche o que faltar (falha total ou parcial) com o fallback
    if np.isnan(tempo_min).any():
        dist_fb = haversine_matrix(np.array(lats), np.array(lons), np.array(lats), np.array(lons))
        tempo_fb = dist_fb / (VELOCIDADE_KMH / 60)
        mask = np.isnan(tempo_min)
        n_fallback = int(mask.sum())
        if n_fallback and n_fallback < mask.size:
            print(f"   ⚠️  {n_fallback}/{mask.size} pares sem resposta do OSRM — preenchidos por fallback.")
        tempo_min[mask] = tempo_fb[mask]
        dist_km[mask] = dist_fb[mask]

    _osrm_cache[key] = {'tempo_min': tempo_min.tolist(), 'dist_km': dist_km.tolist()}
    _osrm_save_cache()
    return tempo_min, dist_km


# ═══════════════════════════════════════════════════════════════════════════
# 2. ISOLAR AS FREGUESIAS DO ALGARVE + DEMANDA (camas necessárias) + CAPACIDADE ATUAL
# ═══════════════════════════════════════════════════════════════════════════
df_freg['ULS_norm'] = df_freg['ULS'].apply(normalize_uls_final)
df_oferta['ULS_norm'] = df_oferta['ULS'].apply(normalize_uls_final)
camas_col = [c for c in df_oferta.columns if 'cama' in c.lower()][0]

capacidade_faro = df_oferta.loc[df_oferta['ULS_norm'].str.contains('algarve'), camas_col].sum()
print(f"🏥 Capacidade atual do Hospital de Faro (ULS do Algarve): {capacidade_faro:.1f} camas")
print(f"🏥 Capacidade assumida do novo hospital: {CAPACIDADE_NOVO_HOSPITAL} camas\n")

demanda_col = [c for c in df_procura.columns if 'camas necess' in c.lower()][0]
distrito_p = [c for c in df_procura.columns if 'distrito' in c.lower()][0]
concelho_p = [c for c in df_procura.columns if 'concelho' in c.lower()][0]
freguesia_p = [c for c in df_procura.columns if c.lower() == 'freguesia'][0]
df_procura['_key'] = (df_procura[distrito_p].apply(normalize_text) + '|' +
                       df_procura[concelho_p].apply(normalize_text) + '|' +
                       df_procura[freguesia_p].apply(normalize_text))
demanda_freg = df_procura.groupby('_key')[demanda_col].sum().rename('Demanda')

df_freg['_key'] = (df_freg['Distrito'].apply(normalize_text) + '|' +
                    df_freg['Concelho'].apply(normalize_text) + '|' +
                    df_freg['Freguesia'].apply(normalize_text))
df_freg = df_freg.merge(demanda_freg, on='_key', how='left')
df_freg['Demanda'] = df_freg['Demanda'].fillna(df_freg['Demanda'].median())

algarve = df_freg[df_freg['Distrito'] == DISTRITO_ALVO].reset_index(drop=True)
n = len(algarve)
demanda_total_algarve = algarve['Demanda'].sum()
print(f"🌴 Região do Algarve: {n} freguesias, {int(algarve['Populacão total'].sum())} habitantes, "
      f"{demanda_total_algarve:.1f} camas necessárias\n")

# Tempo atual até ao Hospital de Faro + tempo entre todas as freguesias
# candidatas — calculados de uma só vez com o OSRM (distâncias/tempos reais
# de estrada, seguindo a rede viária do OpenStreetMap, não linha reta).
print("🗺️  A calcular distâncias/tempos reais de estrada via OSRM...")
lat_todos = np.concatenate([algarve['Lat_Freguesia'].values, algarve['Lat do Hospital'].values[:1]])
lon_todos = np.concatenate([algarve['Long_Freguesia'].values, algarve['Long do Hospital'].values[:1]])
tempo_osrm_min, dist_osrm_km = osrm_table_km_min(lat_todos, lon_todos)

dist_atual_km = dist_osrm_km[:n, n]                 # freguesia -> Hospital de Faro
tempo_atual = tempo_osrm_min[:n, n] + TEMPO_FIXO_MIN
algarve['tempo_atual_min'] = tempo_atual

tempo_medio_atual = np.average(tempo_atual, weights=algarve['Populacão total'])
pct_criticas_atual = round((tempo_atual > 60).mean() * 100, 1)
defice_atual = max(0, demanda_total_algarve - capacidade_faro)
print(f"   Tempo médio ponderado ATUAL (só Hospital de Faro): {tempo_medio_atual:.1f} min")
print(f"   Freguesias críticas (>60min) ATUAL: {pct_criticas_atual}%")
print(f"   Défice de camas ATUAL (região só com H. de Faro): {defice_atual:.1f} camas\n")


# ═══════════════════════════════════════════════════════════════════════════
# 3. AVALIAR TODAS AS FREGUESIAS COMO CANDIDATAS (LP capacitado de 2 hospitais)
# ═══════════════════════════════════════════════════════════════════════════
print("⚙️  A avaliar as 67 freguesias do Algarve como candidatas (com capacidade)...")

lat = algarve['Lat_Freguesia'].values
lon = algarve['Long_Freguesia'].values
pop = algarve['Populacão total'].values
demanda = algarve['Demanda'].values

# Reaproveita a matriz OSRM já calculada acima (freguesia <-> freguesia é o
# submatriz [0:n, 0:n] da matriz combinada com o Hospital de Faro).
dist_candidatos_km = dist_osrm_km[:n, :n]
tempo_candidatos = tempo_osrm_min[:n, :n] + TEMPO_FIXO_MIN


def resolver_2_hospitais(tempo_novo_col):
    """Resolve o problema de transportes capacitado com 2 hospitais
    (Faro existente vs candidato) e devolve: custo total ponderado,
    défice residual, e vetor de atribuição ótima (fração p/ novo hospital)."""
    prob = pulp.LpProblem("Loc_Algarve", pulp.LpMinimize)
    x_novo = [pulp.LpVariable(f"x_novo_{i}", lowBound=0, upBound=1) for i in range(n)]
    overflow_faro = pulp.LpVariable("overflow_faro", lowBound=0)
    overflow_novo = pulp.LpVariable("overflow_novo", lowBound=0)

    custo = pulp.lpSum(
        pop[i] * (x_novo[i] * tempo_novo_col[i] + (1 - x_novo[i]) * tempo_atual[i])
        for i in range(n)
    ) + PENALIZACAO_OVERFLOW * (overflow_faro + overflow_novo)
    prob += custo

    prob += pulp.lpSum(demanda[i] * x_novo[i] for i in range(n)) - overflow_novo <= CAPACIDADE_NOVO_HOSPITAL
    prob += pulp.lpSum(demanda[i] * (1 - x_novo[i]) for i in range(n)) - overflow_faro <= capacidade_faro

    prob.solve(pulp.PULP_CBC_CMD(msg=0))
    frac_novo = np.array([v.value() if v.value() is not None else 0 for v in x_novo])
    custo_val = pulp.value(prob.objective)
    defice_val = (overflow_faro.value() or 0) + (overflow_novo.value() or 0)
    return custo_val, defice_val, frac_novo


custos, defices = [], []
for c in range(n):
    custo_c, defice_c, _ = resolver_2_hospitais(tempo_candidatos[:, c])
    custos.append(custo_c)
    defices.append(defice_c)

algarve['custo_ponderado'] = custos
algarve['defice_camas'] = defices
algarve['tempo_medio_se_aqui'] = np.array(custos) / pop.sum()   # aproximação (inclui penalização se houver défice)

melhor_idx = int(np.argmin(custos))
melhor_local = algarve.iloc[melhor_idx]
print(f"✅ Melhor localização: {melhor_local['Freguesia']} ({melhor_local['Concelho']})\n")


# ═══════════════════════════════════════════════════════════════════════════
# 4. IMPACTO DETALHADO NA MELHOR LOCALIZAÇÃO
# ═══════════════════════════════════════════════════════════════════════════
_, defice_melhor, frac_novo_melhor = resolver_2_hospitais(tempo_candidatos[:, melhor_idx])
tempo_final_melhor = frac_novo_melhor * tempo_candidatos[:, melhor_idx] + (1 - frac_novo_melhor) * tempo_atual
algarve['tempo_com_novo_hospital_min'] = tempo_final_melhor
algarve['fracao_novo_hospital'] = frac_novo_melhor
algarve['usa_novo_hospital'] = frac_novo_melhor > 0.5
algarve['poupanca_min'] = tempo_atual - tempo_final_melhor

tempo_medio_otimo = np.average(tempo_final_melhor, weights=pop)
pct_criticas_otimo = round((tempo_final_melhor > 60).mean() * 100, 1)
n_muda = int(algarve['usa_novo_hospital'].sum())
pop_muda = int(algarve.loc[algarve['usa_novo_hospital'], 'Populacão total'].sum())
demanda_novo = float((algarve['Demanda'] * frac_novo_melhor).sum())
demanda_faro_resid = float((algarve['Demanda'] * (1 - frac_novo_melhor)).sum())

print("=" * 70)
print("RESUMO")
print("=" * 70)
print(f"Melhor localização         : {melhor_local['Freguesia']} ({melhor_local['Concelho']})")
print(f"Tempo médio ATUAL           : {tempo_medio_atual:.1f} min")
print(f"Tempo médio COM novo hosp.  : {tempo_medio_otimo:.1f} min  "
      f"({tempo_medio_atual - tempo_medio_otimo:+.1f} min)")
print(f"Freguesias críticas (>60min) ATUAL -> ÓTIMO: {pct_criticas_atual}% -> {pct_criticas_otimo}%")
print(f"Freguesias que passariam a usar o novo hospital: {n_muda}/{n} "
      f"({pop_muda} habitantes, {round(pop_muda/pop.sum()*100,1)}% da região)")
print(f"Procura atribuída: Novo hospital = {demanda_novo:.1f}/{CAPACIDADE_NOVO_HOSPITAL} camas | "
      f"H. de Faro (residual) = {demanda_faro_resid:.1f}/{capacidade_faro:.1f} camas")
print(f"Défice de camas residual (região): {defice_melhor:.1f} (ATUAL era {defice_atual:.1f})")
print("=" * 70)


# ═══════════════════════════════════════════════════════════════════════════
# 5. OUTPUT EXCEL
# ═══════════════════════════════════════════════════════════════════════════
print("\n💾 A gerar ficheiro Excel...")
wb = Workbook()
HEADER_FILL = PatternFill(start_color="2E9E6B", end_color="2E9E6B", fill_type="solid")
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
ws['B2'] = "MODELO 2 — LOCALIZAÇÃO ÓTIMA DE NOVO HOSPITAL NO ALGARVE (COM CAPACIDADE)"
ws['B2'].font = Font(bold=True, size=14, color="2E9E6B")
ws['B3'] = ("Ponderação por população: cada freguesia pesa Peso_i = População_i / População_total_Algarve. "
            "O modelo minimiza Σ Peso_i × Tempo_i (não a distância mínima 'no abstrato') — o critério já é "
            "'onde estão as pessoas', não só a melhor localização geométrica. Ver detalhe por freguesia na folha "
            "'3. Impacto_Melhor_Local' (colunas Peso_Populacional_% e Tempo_Ponderado).")
ws['B3'].font = Font(italic=True, size=9, color="666666")
ws.merge_cells('B3:C3')
linhas = [
    ("", ""),
    ("Nº de freguesias no Algarve", n),
    ("População total do Algarve", int(pop.sum())),
    ("Camas necessárias totais na região", round(demanda_total_algarve, 1)),
    ("", ""),
    ("★ Melhor localização (Freguesia)", melhor_local['Freguesia']),
    ("★ Concelho", melhor_local['Concelho']),
    ("★ Coordenadas (lat, lon)", f"{melhor_local['Lat_Freguesia']:.4f}, {melhor_local['Long_Freguesia']:.4f}"),
    ("", ""),
    ("Tempo médio ponderado ATUAL (só H. de Faro)", round(tempo_medio_atual, 1)),
    ("Tempo médio ponderado COM novo hospital", round(tempo_medio_otimo, 1)),
    ("Redução de tempo médio (min)", round(tempo_medio_atual - tempo_medio_otimo, 1)),
    ("", ""),
    ("% freguesias críticas (>60min) ATUAL", pct_criticas_atual),
    ("% freguesias críticas (>60min) COM novo hospital", pct_criticas_otimo),
    ("", ""),
    ("Freguesias que passam a usar o novo hospital", n_muda),
    ("População que passa a usar o novo hospital", pop_muda),
    ("% da população regional que muda", round(pop_muda / pop.sum() * 100, 1)),
    ("", ""),
    ("Capacidade do Hospital de Faro (existente)", round(capacidade_faro, 1)),
    ("Capacidade do novo hospital (assumida)", CAPACIDADE_NOVO_HOSPITAL),
    ("Procura atribuída ao novo hospital", round(demanda_novo, 1)),
    ("Procura residual no Hospital de Faro", round(demanda_faro_resid, 1)),
    ("Défice de camas ATUAL (região, só H. Faro)", round(defice_atual, 1)),
    ("Défice de camas COM novo hospital", round(defice_melhor, 1)),
    ("", ""),
    ("Parâmetros do modelo", ""),
    ("Velocidade assumida (km/h)", VELOCIDADE_KMH),
    ("Tempo fixo de acesso (min)", TEMPO_FIXO_MIN),
    ("Nº de candidatos avaliados", n),
]
for i, (k, v) in enumerate(linhas, 4):
    ws.cell(row=i, column=2, value=k).font = Font(bold=True)
    ws.cell(row=i, column=3, value=v)
ws.column_dimensions['B'].width = 48
ws.column_dimensions['C'].width = 24

# --- Folha 2: Ranking de candidatos -------------------------------------------
ws2 = wb.create_sheet("2. Ranking_Candidatos")
ws2['A1'] = ("Ranking das 67 freguesias do Algarve como candidatas a novo hospital "
             f"({CAPACIDADE_NOVO_HOSPITAL} camas) — menor custo = melhor")
ws2['A1'].font = Font(italic=True, size=9, color="666666")
cols2 = ['Concelho', 'Freguesia', 'Populacão total', 'Lat_Freguesia', 'Long_Freguesia',
         'tempo_medio_se_aqui', 'defice_camas', 'custo_ponderado']
out2 = algarve[cols2].rename(columns={'Populacão total': 'Populacao'}).copy()
out2['tempo_medio_se_aqui'] = out2['tempo_medio_se_aqui'].round(2)
out2['defice_camas'] = out2['defice_camas'].round(1)
out2['custo_ponderado'] = out2['custo_ponderado'].round(0)
out2 = out2.sort_values('custo_ponderado').reset_index(drop=True)
out2.insert(0, 'Ranking', out2.index + 1)
for j, h in enumerate(out2.columns, 1):
    ws2.cell(row=2, column=j, value=h)
for i, (_, row) in enumerate(out2.iterrows(), 3):
    for j, h in enumerate(out2.columns, 1):
        ws2.cell(row=i, column=j, value=row[h])
style_header(ws2, len(out2.columns), row=2)
ws2.freeze_panes = 'A3'
autofit(ws2, len(out2.columns))

# --- Folha 3: Impacto por freguesia (com a melhor localização) ---------------
# Peso populacional explícito: cada freguesia pesa proporcionalmente à sua
# população no total do Algarve (Σ peso_i = 1). O "tempo ponderado" de cada
# freguesia é peso_i × tempo_i — a SOMA de todos os tempos ponderados é
# exatamente a média ponderada por população reportada no Resumo (não é
# apenas "a melhor localização no abstrato", é literalmente "onde estão as
# pessoas" a pesar mais no resultado). Isto já era o critério usado pelo
# solver (minimizar Σ pop_i×tempo_i) — aqui só se torna visível/auditável.
algarve['peso_populacional'] = algarve['Populacão total'] / algarve['Populacão total'].sum()
algarve['tempo_ponderado_atual'] = algarve['peso_populacional'] * algarve['tempo_atual_min']
algarve['tempo_ponderado_novo'] = algarve['peso_populacional'] * algarve['tempo_com_novo_hospital_min']

ws3 = wb.create_sheet("3. Impacto_Melhor_Local")
ws3['A1'] = (f"Impacto se o novo hospital ({CAPACIDADE_NOVO_HOSPITAL} camas) ficar em: "
             f"{melhor_local['Freguesia']} ({melhor_local['Concelho']})  |  "
             "Peso_Populacional = População_da_freguesia / População_total_Algarve  |  "
             "Tempo_Ponderado = Peso_Populacional × Tempo  |  "
             "Σ Tempo_Ponderado (todas as freguesias) = tempo médio ponderado por população (ver Resumo e bloco de verificação à direita).")
ws3['A1'].font = Font(italic=True, size=9, color="666666")
cols3 = ['Concelho', 'Freguesia', 'Populacão total', 'Demanda', 'tempo_atual_min',
         'tempo_com_novo_hospital_min', 'poupanca_min', 'fracao_novo_hospital', 'usa_novo_hospital',
         'peso_populacional', 'tempo_ponderado_atual', 'tempo_ponderado_novo']
out3 = algarve[cols3].rename(columns={'Populacão total': 'Populacao', 'Demanda': 'Camas_Necessarias',
                                       'peso_populacional': 'Peso_Populacional_%',
                                       'tempo_ponderado_atual': 'Tempo_Ponderado_Atual',
                                       'tempo_ponderado_novo': 'Tempo_Ponderado_Novo'}).copy()
out3['tempo_atual_min'] = out3['tempo_atual_min'].round(1)
out3['tempo_com_novo_hospital_min'] = out3['tempo_com_novo_hospital_min'].round(1)
out3['poupanca_min'] = out3['poupanca_min'].round(1)
out3['fracao_novo_hospital'] = out3['fracao_novo_hospital'].round(2)
out3['usa_novo_hospital'] = out3['usa_novo_hospital'].map({True: 'Sim', False: 'Não'})
out3['Peso_Populacional_%'] = out3['Peso_Populacional_%'].round(4)
out3['Tempo_Ponderado_Atual'] = out3['Tempo_Ponderado_Atual'].round(3)
out3['Tempo_Ponderado_Novo'] = out3['Tempo_Ponderado_Novo'].round(3)
out3 = out3.sort_values('poupanca_min', ascending=False)
for j, h in enumerate(out3.columns, 1):
    ws3.cell(row=2, column=j, value=h)
for i, (_, row) in enumerate(out3.iterrows(), 3):
    for j, h in enumerate(out3.columns, 1):
        ws3.cell(row=i, column=j, value=row[h])
last_row3 = 2 + len(out3)
for col_letter in ('J',):
    for r in range(3, last_row3 + 1):
        ws3[f'{col_letter}{r}'].number_format = '0.00%'
style_header(ws3, len(out3.columns), row=2)
ws3.freeze_panes = 'A3'
autofit(ws3, len(out3.columns))

# ── Bloco de verificação (fora do intervalo A:I usado pelo dashboard) ───────
# Confirma que a soma dos tempos ponderados por freguesia bate certo com as
# médias ponderadas por população já reportadas na folha "1. Resumo".
ws3['N2'] = 'Verificação — soma dos tempos ponderados (deve bater com a folha 1. Resumo)'
ws3['N2'].font = Font(italic=True, size=9, color="666666")
ws3['N3'] = 'Σ Peso_Populacional (deve ser 100%)'
ws3['O3'] = f'=SUM(J3:J{last_row3})'
ws3['O3'].number_format = '0.00%'
ws3['N4'] = 'Σ Tempo_Ponderado_Atual = Tempo médio ATUAL'
ws3['O4'] = f'=ROUND(SUM(K3:K{last_row3}),1)'
ws3['N5'] = 'Σ Tempo_Ponderado_Novo = Tempo médio COM novo hospital'
ws3['O5'] = f'=ROUND(SUM(L3:L{last_row3}),1)'
for r in (3, 4, 5):
    ws3[f'N{r}'].font = Font(size=10, color="666666")
    ws3[f'O{r}'].font = Font(bold=True, size=10, color="2E9E6B")
ws3.column_dimensions['N'].width = 45
ws3.column_dimensions['O'].width = 14

output_path = os.path.join(SCRIPT_DIR, 'Modelo2_Localizacao_Algarve.xlsx')
wb.save(output_path)
print(f"✅ Ficheiro guardado em: {output_path}")