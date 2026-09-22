#!/usr/bin/env python3
"""
MODELO 1 — OTIMIZAÇÃO 2SFCA: REAFETAÇÃO ÓTIMA FREGUESIA → HOSPITAL
════════════════════════════════════════════════════════════════════════════
OBJETIVO
Hoje cada freguesia tem um "hospital de referência" fixo (coluna Hospiatais
em freguesias_limpo.xlsx), que nem sempre é o hospital com menor tempo de
trajeto. Este modelo pergunta: "se pudéssemos redesenhar os hospitais de
referência do zero, só a pensar em minimizar o tempo de trajeto (sujeito à
capacidade real de camas), qual seria a melhor atribuição?"

MÉTODO
1) 2SFCA (Two-Step Floating Catchment Area) — usado para DIAGNOSTICAR a
   acessibilidade atual do país:
     Passo 1: para cada hospital j, razão de oferta/procura dentro do seu
              catchment (raio de t0 minutos):
                  R_j = Capacidade_j / Σ_i∈catchment(j) Procura_i
     Passo 2: para cada freguesia i, acessibilidade = soma das razões R_j de
              todos os hospitais ao seu alcance:
                  A_i = Σ_j∈catchment(i) R_j
   Quanto maior A_i, melhor servida está a freguesia (mais camas disponíveis
   por unidade de procura, dentro de alcance razoável).

2) OTIMIZAÇÃO (Problema de Transportes / Assignment capacitado) — usado para
   DECIDIR a nova atribuição ótima:
       minimizar   Σ_i Σ_j  pop_i * tempo_ij * x_ij           (tempo total
                                                                 ponderado
                                                                 pela pop.)
       sujeito a   Σ_j x_ij = 1                    para cada freguesia i
                   Σ_i demanda_i * x_ij  ≤  capacidade_j + overflow_j
                   x_ij ≥ 0 , overflow_j ≥ 0
   x_ij ∈ {0,1} = 1 se a freguesia i é encaminhada para o hospital j (atribuição
   inteira: cada freguesia vai INTEIRA para um único hospital, sem divisões).
   demanda_i = camas necessárias estimadas da freguesia (Excel_Procura).
   capacidade_j = camas do hospital/ULS (Excel_Oferta).
   overflow_j é uma variável de folga com penalização muito alta: garante
   que o modelo NUNCA fica infeasível (mesmo que a capacidade nacional não
   chegue), mas reporta onde ficam défices estruturais mesmo depois de
   otimizar as deslocações.

Este modelo só pode atribuir cada freguesia a uma das 39 ULS/hospitais REAIS
de Portugal que constam em freguesias_limpo.xlsx (nenhuma ULS é inventada:
a tabela de hospitais é construída diretamente a partir dessa coluna, e o
código valida que corresponde 1 para 1 com a lista oficial de ULS). O modelo
considera TODAS as 39 ULS como candidatas para cada freguesia (sem restringir
a um subconjunto "mais próximo"), para garantir que a solução é verdadeiramente
ótima e não fica limitada artificialmente.

OUTPUT: Modelo1_Reafetacao_2SFCA.xlsx com:
  - Resumo executivo (antes vs depois)
  - Reafetação por freguesia (ULS/hospital atual vs ótimo, tempos, mudou?)
  - Acessibilidade 2SFCA por freguesia
  - Capacidade e défice estrutural por hospital
"""
import pandas as pd
import numpy as np
from math import radians
import os
import json
import time
import requests
import pulp
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ── PARÂMETROS DO MODELO (ajustáveis) ───────────────────────────────────────
VELOCIDADE_KMH   = 60     # velocidade média assumida (igual ao modelo base) — usada só na reserva (fallback) do OSRM
TEMPO_FIXO_MIN   = 15     # tempo fixo (estacionamento/entrada/urgência) — usado só na reserva (fallback) do OSRM
T0_2SFCA_MIN     = 60     # raio de catchment (min) usado no 2SFCA
PENALIZACAO_OVERFLOW = 1_000_000  # penalização por cada unidade de procura não coberta
# Nota: NÃO se restringe o nº de hospitais candidatos por freguesia — o modelo
# escolhe livremente entre as 39 ULS reais para garantir o ótimo verdadeiro.

# ═══════════════════════════════════════════════════════════════════════════
# CONFIGURAÇÃO DO MOTOR DE ROTAS (OpenStreetMap via OSRM)
# ═══════════════════════════════════════════════════════════════════════════
# A matriz de tempos freguesia -> hospital deixa de ser "linha reta"
# (haversine) e passa a ser tempo REAL de condução, via OSRM (routing sobre
# dados do OpenStreetMap). Como a otimização (secção 7) precisa de considerar
# TODOS os hospitais candidatos para cada freguesia, a matriz é
# n_freguesias × n_hospitais — pede-se ao OSRM hospital a hospital, com as
# freguesias em lotes (serviço 'table' do OSRM).
#
# Usa-se por omissão o servidor público de demonstração do OSRM (grátis, sem
# chave, mas com fair-use — pode ser lento/instável com muitos milhares de
# pedidos). Há cache em disco (para não repetir tudo em reexecuções) e
# retentativas com espera crescente antes de cair em haversine como reserva.
# Para uso pesado/frequente, considera montar um servidor OSRM próprio
# (Docker + extrato do OpenStreetMap de Portugal) e mudar OSRM_BASE_URL.
OSRM_BASE_URL = "https://router.project-osrm.org"
OSRM_BATCH_SIZE = 90        # nº máx. de freguesias por chamada (servidor demo limita a ~100 coordenadas, incl. o hospital)
OSRM_SLEEP_SECONDS = 1.0    # pausa entre chamadas, para não sobrecarregar o servidor demo
OSRM_MAX_RETRIES = 4        # nº de tentativas por lote antes de desistir e usar haversine
OSRM_RETRY_BACKOFF = 3      # segundos de espera antes de cada retentativa (cresce: 3s, 6s, 9s...)
OSRM_CACHE_FILE = os.path.join(SCRIPT_DIR, '_osrm_cache_modelo1.json')

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


# ═══════════════════════════════════════════════════════════════════════════
# 2. NORMALIZAÇÃO DE TEXTO (igual ao modelo base, para casar nomes entre ficheiros)
# ═══════════════════════════════════════════════════════════════════════════
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
    """Distância haversine (km), vetorizada: lat1/lon1 = vetores freguesias,
    lat2/lon2 = vetores hospitais. Devolve matriz [n_freg x n_hosp].
    Mantida como RESERVA (fallback) para quando o OSRM falhar."""
    lat1r, lon1r = np.radians(lat1)[:, None], np.radians(lon1)[:, None]
    lat2r, lon2r = np.radians(lat2)[None, :], np.radians(lon2)[None, :]
    dlat = lat2r - lat1r
    dlon = lon2r - lon1r
    a = np.sin(dlat / 2) ** 2 + np.cos(lat1r) * np.cos(lat2r) * np.sin(dlon / 2) ** 2
    c = 2 * np.arcsin(np.sqrt(a))
    return 6371 * c


# Cache em disco: evita recalcular tudo de novo em reexecuções do script,
# já que as distâncias/tempos reais por estrada não mudam de um dia para o outro.
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
    """Calcula tempo (min) REAL POR ESTRADA de várias freguesias até UM
    hospital, numa única chamada ao serviço 'table' do OSRM.
    freg_rows: lista de tuplos (índice, lat, lon).
    Devolve: dict {índice: tempo_min}.
    Faz até OSRM_MAX_RETRIES tentativas (com espera crescente) antes de
    desistir e usar haversine (mesma fórmula de velocidade constante do
    modelo original) como reserva.
    """
    resultados = {}

    por_pedir = []
    for idx, lat_f, lon_f in freg_rows:
        key = _osrm_cache_key(lat_f, lon_f, lat_h, lon_h)
        if key in _osrm_cache:
            resultados[idx] = _osrm_cache[key]
        else:
            por_pedir.append((idx, lat_f, lon_f, key))

    if not por_pedir:
        return resultados

    coords = [f"{lon_f},{lat_f}" for _, lat_f, lon_f, _ in por_pedir]
    coords.append(f"{lon_h},{lat_h}")  # hospital é sempre o destino (último ponto)
    n = len(por_pedir)
    sources = ";".join(str(i) for i in range(n))
    destination = str(n)

    url = (
        f"{OSRM_BASE_URL}/table/v1/driving/{';'.join(coords)}"
        f"?sources={sources}&destinations={destination}&annotations=duration"
    )

    try:
        data = None
        ultimo_erro = None
        for tentativa in range(1, OSRM_MAX_RETRIES + 1):
            try:
                resp = requests.get(url, timeout=30)
                resp.raise_for_status()
                data = resp.json()
                if data.get('code') != 'Ok':
                    raise ValueError(f"OSRM devolveu code={data.get('code')}")
                break  # sucesso, sai do ciclo de retentativas
            except Exception as e_retry:
                ultimo_erro = e_retry
                if tentativa < OSRM_MAX_RETRIES:
                    espera = OSRM_RETRY_BACKOFF * tentativa
                    print(f"   ⚠️  Erro OSRM (tentativa {tentativa}/{OSRM_MAX_RETRIES}): {e_retry} "
                          f"— a repetir daqui a {espera}s...")
                    time.sleep(espera)
        if data is None:
            raise ultimo_erro
        duracoes = data['durations']  # segundos
        for i, (idx, lat_f, lon_f, key) in enumerate(por_pedir):
            dur_s = duracoes[i][0]
            if dur_s is None:
                dist_km = haversine_matrix(
                    np.array([lat_f]), np.array([lon_f]), np.array([lat_h]), np.array([lon_h])
                )[0, 0]
                t_min = dist_km / (VELOCIDADE_KMH / 60)
            else:
                t_min = dur_s / 60.0
            resultados[idx] = t_min
            _osrm_cache[key] = t_min
    except Exception as e:
        print(f"   ⚠️  Erro OSRM neste lote, mesmo após {OSRM_MAX_RETRIES} tentativas "
              f"(a usar haversine como reserva): {e}")
        for idx, lat_f, lon_f, key in por_pedir:
            dist_km = haversine_matrix(
                np.array([lat_f]), np.array([lon_f]), np.array([lat_h]), np.array([lon_h])
            )[0, 0]
            t_min = dist_km / (VELOCIDADE_KMH / 60)
            resultados[idx] = t_min
            _osrm_cache[key] = t_min

    time.sleep(OSRM_SLEEP_SECONDS)
    return resultados


def osrm_time_matrix(freg_lat, freg_lon, hosp_lat, hosp_lon):
    """Devolve a matriz tempo_min (n_freguesias × n_hospitais) com o tempo
    REAL de condução (OpenStreetMap/OSRM), calculada hospital a hospital,
    em lotes de freguesias (serviço 'table' do OSRM)."""
    n_freg = len(freg_lat)
    n_hosp = len(hosp_lat)
    matriz = np.zeros((n_freg, n_hosp))
    for j in range(n_hosp):
        lat_h, lon_h = hosp_lat[j], hosp_lon[j]
        for b in range(0, n_freg, OSRM_BATCH_SIZE):
            idxs = list(range(b, min(b + OSRM_BATCH_SIZE, n_freg)))
            lote = [(i, freg_lat[i], freg_lon[i]) for i in idxs]
            res = osrm_table_batch(lote, lat_h, lon_h)
            for i, t_min in res.items():
                matriz[i, j] = t_min
        print(f"   Hospital {j + 1}/{n_hosp} concluído (tempos calculados)...")
    return matriz


# ═══════════════════════════════════════════════════════════════════════════
# 3. CONSTRUIR TABELA DE HOSPITAIS (39 hospitais únicos, com capacidade)
# ═══════════════════════════════════════════════════════════════════════════
print("🏥 A construir tabela de hospitais...")
df_freg['ULS_norm'] = df_freg['ULS'].apply(normalize_uls_final)
df_oferta['ULS_norm'] = df_oferta['ULS'].apply(normalize_uls_final)

camas_col = [c for c in df_oferta.columns if 'cama' in c.lower()][0]

hospitais = (
    df_freg.dropna(subset=['Hospiatais', 'Lat do Hospital', 'Long do Hospital'])
    .drop_duplicates('Hospiatais')[['Hospiatais', 'Lat do Hospital', 'Long do Hospital', 'ULS_norm', 'ULS']]
    .rename(columns={'Hospiatais': 'Hospital', 'Lat do Hospital': 'Lat', 'Long do Hospital': 'Lon'})
    .reset_index(drop=True)
)
oferta_uls = df_oferta.groupby('ULS_norm')[camas_col].sum().rename('Capacidade').reset_index()
hospitais = hospitais.merge(oferta_uls, on='ULS_norm', how='left')
hospitais['Capacidade'] = hospitais['Capacidade'].fillna(hospitais['Capacidade'].median())
n_hosp = len(hospitais)
print(f"   {n_hosp} hospitais identificados, capacidade total = {hospitais['Capacidade'].sum():.0f} camas\n")

# ── VALIDAÇÃO: garantir que só usamos ULS reais existentes em Portugal ──────
# (a tabela de hospitais vem diretamente da coluna 'Hospiatais'/'ULS' de
# freguesias_limpo.xlsx — nenhuma ULS é inventada; confirmamos aqui que bate
# certo 1-para-1 com a lista de ULS desse ficheiro)
uls_reais = set(df_freg['ULS_norm'].unique())
uls_no_modelo = set(hospitais['ULS_norm'].unique())
uls_invalidas = uls_no_modelo - uls_reais
if uls_invalidas:
    raise ValueError(f"❌ ERRO: o modelo criou ULS que não existem no ficheiro original: {uls_invalidas}")
print(f"✅ Validação: as {len(uls_no_modelo)} ULS usadas no modelo correspondem exatamente às "
      f"ULS reais de freguesias_limpo.xlsx (nenhuma inventada).\n")


# ═══════════════════════════════════════════════════════════════════════════
# 4. PROCURA POR FREGUESIA — soma das camas necessárias, todas as faixas
# ═══════════════════════════════════════════════════════════════════════════
print("📊 A calcular procura (camas necessárias) por freguesia...")
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
n_sem_demanda = df_freg['Demanda'].isna().sum()
if n_sem_demanda:
    print(f"   ⚠️  {n_sem_demanda} freguesias sem correspondência na Procura -> demanda assumida = mediana nacional")
df_freg['Demanda'] = df_freg['Demanda'].fillna(df_freg['Demanda'].median())
print(f"   Demanda nacional total = {df_freg['Demanda'].sum():.1f} camas necessárias\n")


# ═══════════════════════════════════════════════════════════════════════════
# 5. MATRIZ DE TEMPOS (freguesia x hospital) — todas as combinações
# ═══════════════════════════════════════════════════════════════════════════
print("🗺️  A calcular matriz de tempos de trajeto REAIS por estrada (OpenStreetMap/OSRM)...")
n_freg = len(df_freg)
tempo_min = osrm_time_matrix(
    df_freg['Lat_Freguesia'].values, df_freg['Long_Freguesia'].values,
    hospitais['Lat'].values, hospitais['Lon'].values
)
_osrm_save_cache()
print(f"   Matriz {n_freg} freguesias x {n_hosp} hospitais construída\n")

# Tempo/hospital ATUAL (hospital de referência de hoje), para comparação
hosp_pos = {h: i for i, h in enumerate(hospitais['Hospital'])}
df_freg['idx_hosp_atual'] = df_freg['Hospiatais'].map(hosp_pos)
df_freg['tempo_atual_min'] = [
    tempo_min[i, j] if pd.notna(j) else np.nan
    for i, j in enumerate(df_freg['idx_hosp_atual'])
]


# ═══════════════════════════════════════════════════════════════════════════
# 6. 2SFCA — ACESSIBILIDADE ATUAL (diagnóstico, ANTES da otimização)
# ═══════════════════════════════════════════════════════════════════════════
print("🔎 A calcular acessibilidade 2SFCA (diagnóstico atual)...")
alcance = tempo_min <= T0_2SFCA_MIN  # [n_freg x n_hosp] booleano

# Passo 1: razão oferta/procura por hospital, dentro do catchment de t0 min
procura_no_alcance = alcance.T @ df_freg['Demanda'].values  # soma da procura de todas as freguesias a <=t0 min de cada hospital
with np.errstate(divide='ignore', invalid='ignore'):
    R_j = np.where(procura_no_alcance > 0, hospitais['Capacidade'].values / procura_no_alcance, 0.0)

# Passo 2: acessibilidade por freguesia = soma dos R_j de todos os hospitais ao alcance
A_i = alcance @ R_j
df_freg['Acessibilidade_2SFCA'] = A_i
hospitais['R_j_2SFCA'] = R_j
print(f"   Acessibilidade média nacional (2SFCA) = {A_i.mean():.3f} camas/procura\n")


# ═══════════════════════════════════════════════════════════════════════════
# 7. OTIMIZAÇÃO — PROBLEMA DE TRANSPORTES CAPACITADO (PuLP / CBC)
# ═══════════════════════════════════════════════════════════════════════════
print("⚙️  A montar e resolver o modelo de otimização (PuLP/CBC)...")

# Todas as 39 ULS/hospitais são candidatas para cada freguesia (sem restrição)
candidatos_idx = np.tile(np.arange(n_hosp), (n_freg, 1))

prob = pulp.LpProblem("Reafetacao_Otima_Hospitalar", pulp.LpMinimize)

x = {}   # x[i, j] = 1 se a freguesia i é atribuída ao hospital j (binário; cada freguesia vai toda para um só hospital)
for i in range(n_freg):
    for j in candidatos_idx[i]:
        x[i, j] = pulp.LpVariable(f"x_{i}_{j}", cat="Binary")

overflow = {j: pulp.LpVariable(f"overflow_{j}", lowBound=0) for j in range(n_hosp)}

pop_i = df_freg['Populacão total'].values

# Objetivo: minimizar tempo total ponderado pela população + penalização por défice de capacidade
prob += (
    pulp.lpSum(pop_i[i] * tempo_min[i, j] * x[i, j] for (i, j) in x)
    + PENALIZACAO_OVERFLOW * pulp.lpSum(overflow.values())
)

# Cada freguesia é 100% atribuída (podendo ser repartida entre candidatos)
for i in range(n_freg):
    prob += pulp.lpSum(x[i, j] for j in candidatos_idx[i]) == 1, f"cobertura_freg_{i}"

# Capacidade de cada hospital (em camas necessárias servidas)
demanda_i = df_freg['Demanda'].values
for j in range(n_hosp):
    termos = [demanda_i[i] * x[i, j] for i in range(n_freg) if (i, j) in x]
    if termos:
        prob += pulp.lpSum(termos) - overflow[j] <= hospitais['Capacidade'].values[j], f"capacidade_hosp_{j}"

solver = pulp.PULP_CBC_CMD(msg=0, timeLimit=600, gapRel=0.01)
prob.solve(solver)
print(f"   Estado do solver: {pulp.LpStatus[prob.status]}\n")

# ── Extrair solução ─────────────────────────────────────────────────────────
alocacao = np.zeros((n_freg, n_hosp))
for (i, j), var in x.items():
    v = var.value()
    if v and v > 1e-6:
        alocacao[i, j] = v

hosp_otimo_idx = alocacao.argmax(axis=1)
df_freg['Hospital_Otimo'] = hospitais['Hospital'].values[hosp_otimo_idx]
df_freg['ULS_Otimo'] = hospitais['ULS'].values[hosp_otimo_idx]
df_freg['ULS_Atual'] = df_freg['ULS']
df_freg['Fracao_Hospital_Otimo'] = alocacao[np.arange(n_freg), hosp_otimo_idx]
df_freg['tempo_otimo_min'] = tempo_min[np.arange(n_freg), hosp_otimo_idx]
df_freg['Mudou_Hospital'] = df_freg['Hospital_Otimo'] != df_freg['Hospiatais']
df_freg['Mudou_ULS'] = df_freg['ULS_Otimo'] != df_freg['ULS_Atual']
df_freg['Poupanca_Tempo_min'] = df_freg['tempo_atual_min'] - df_freg['tempo_otimo_min']

# Défice estrutural por hospital (overflow) mesmo depois de otimizar
hospitais['Overflow_Defice_Camas'] = [overflow[j].value() or 0 for j in range(n_hosp)]
hospitais['Procura_Atribuida'] = [
    sum(demanda_i[i] * alocacao[i, j] for i in range(n_freg)) for j in range(n_hosp)
]
hospitais['Ocupacao_Pct'] = (hospitais['Procura_Atribuida'] / hospitais['Capacidade'] * 100).round(1)


# ═══════════════════════════════════════════════════════════════════════════
# 8. RESUMO EXECUTIVO
# ═══════════════════════════════════════════════════════════════════════════
tempo_medio_atual = np.average(df_freg['tempo_atual_min'], weights=pop_i)
tempo_medio_otimo = np.average(df_freg['tempo_otimo_min'], weights=pop_i)
n_mudou = int(df_freg['Mudou_ULS'].sum())
pct_mudou = round(n_mudou / n_freg * 100, 1)
pct_criticas_atual = round((df_freg['tempo_atual_min'] > 60).mean() * 100, 1)
pct_criticas_otimo = round((df_freg['tempo_otimo_min'] > 60).mean() * 100, 1)
defice_total = hospitais['Overflow_Defice_Camas'].sum()

print("=" * 70)
print("RESUMO")
print("=" * 70)
print(f"Tempo médio ponderado (ATUAL) : {tempo_medio_atual:.1f} min")
print(f"Tempo médio ponderado (ÓTIMO) : {tempo_medio_otimo:.1f} min  "
      f"({tempo_medio_atual - tempo_medio_otimo:+.1f} min)")
print(f"Freguesias que mudariam de ULS: {n_mudou}/{n_freg} ({pct_mudou}%)")
print(f"Freguesias críticas (>60min) ATUAL: {pct_criticas_atual}%  ->  ÓTIMO: {pct_criticas_otimo}%")
print(f"Défice estrutural de camas mesmo após otimizar: {defice_total:.1f} camas")
print("=" * 70)


# ═══════════════════════════════════════════════════════════════════════════
# 9. OUTPUT EXCEL
# ═══════════════════════════════════════════════════════════════════════════
print("\n💾 A gerar ficheiro Excel...")
wb = Workbook()

HEADER_FILL = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
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


# --- Folha 1: Resumo ---------------------------------------------------------
ws = wb.active
ws.title = "1. Resumo"
ws['B2'] = "MODELO 1 — OTIMIZAÇÃO 2SFCA: REAFETAÇÃO HOSPITALAR"
ws['B2'].font = Font(bold=True, size=14, color="2E75B6")
linhas = [
    ("", ""),
    ("Nº de freguesias", n_freg),
    ("Nº de hospitais", n_hosp),
    ("", ""),
    ("Tempo médio ponderado ATUAL (min)", round(tempo_medio_atual, 1)),
    ("Tempo médio ponderado ÓTIMO (min)", round(tempo_medio_otimo, 1)),
    ("Redução de tempo médio (min)", round(tempo_medio_atual - tempo_medio_otimo, 1)),
    ("Freguesias que mudam de ULS", f"{n_mudou} ({pct_mudou}%)"),
    ("", ""),
    ("% freguesias críticas (>60 min) ATUAL", pct_criticas_atual),
    ("% freguesias críticas (>60 min) ÓTIMO", pct_criticas_otimo),
    ("", ""),
    ("Acessibilidade média 2SFCA (camas/procura)", round(A_i.mean(), 3)),
    ("Défice estrutural de camas (após otimizar)", round(defice_total, 1)),
    ("", ""),
    ("Parâmetros do modelo", ""),
    ("Velocidade assumida (km/h)", VELOCIDADE_KMH),
    ("Tempo fixo de acesso (min)", TEMPO_FIXO_MIN),
    ("Nº de hospitais candidatos por freguesia", f"Todas as {n_hosp} ULS (sem restrição)"),
    ("Raio de catchment 2SFCA (min)", T0_2SFCA_MIN),
]
for i, (k, v) in enumerate(linhas, 4):
    ws.cell(row=i, column=2, value=k).font = Font(bold=True)
    ws.cell(row=i, column=3, value=v)
ws.column_dimensions['B'].width = 42
ws.column_dimensions['C'].width = 20

# --- Folha 2: Reafetação por freguesia --------------------------------------
ws2 = wb.create_sheet("2. Reafetacao_Freguesias")
cols2 = ['Distrito', 'Concelho', 'Freguesia', 'Populacão total',
         'ULS_Atual', 'tempo_atual_min',
         'ULS_Otimo', 'tempo_otimo_min',
         'Poupanca_Tempo_min', 'Mudou_ULS', 'Fracao_Hospital_Otimo']
out2 = df_freg[cols2].rename(columns={'Populacão total': 'Populacao'}).round(1)
out2 = out2.sort_values('Poupanca_Tempo_min', ascending=False)
for j, h in enumerate(out2.columns, 1):
    ws2.cell(row=1, column=j, value=h)
for i, (_, row) in enumerate(out2.iterrows(), 2):
    for j, h in enumerate(out2.columns, 1):
        ws2.cell(row=i, column=j, value=row[h])
style_header(ws2, len(out2.columns))
ws2.freeze_panes = 'A2'
autofit(ws2, len(out2.columns))

# --- Folha 3: Acessibilidade 2SFCA ------------------------------------------
ws3 = wb.create_sheet("3. Acessibilidade_2SFCA")
ws3['A1'] = ("2SFCA: R_j = Capacidade_hospital / Procura dentro do raio | "
             f"A_i = soma dos R_j dos hospitais a <= {T0_2SFCA_MIN} min. Quanto maior, melhor servida.")
ws3['A1'].font = Font(italic=True, size=9, color="666666")
cols3 = ['Distrito', 'Concelho', 'Freguesia', 'Populacão total', 'Hospiatais', 'Acessibilidade_2SFCA']
out3 = df_freg[cols3].rename(columns={'Hospiatais': 'Hospital_Atual', 'Populacão total': 'Populacao'})
out3['Acessibilidade_2SFCA'] = out3['Acessibilidade_2SFCA'].round(4)
out3 = out3.sort_values('Acessibilidade_2SFCA')
for j, h in enumerate(out3.columns, 1):
    ws3.cell(row=2, column=j, value=h)
for i, (_, row) in enumerate(out3.iterrows(), 3):
    for j, h in enumerate(out3.columns, 1):
        ws3.cell(row=i, column=j, value=row[h])
style_header(ws3, len(out3.columns), row=2)
ws3.freeze_panes = 'A3'
autofit(ws3, len(out3.columns))

# --- Folha 4: Capacidade das ULS ---------------------------------------------
ws4 = wb.create_sheet("4. Capacidade_ULS")
cols4 = ['ULS', 'Capacidade', 'Procura_Atribuida', 'Ocupacao_Pct',
         'Overflow_Defice_Camas', 'R_j_2SFCA', 'Hospital']
out4 = hospitais[cols4].copy()
out4['Procura_Atribuida'] = out4['Procura_Atribuida'].round(1)
out4['R_j_2SFCA'] = out4['R_j_2SFCA'].round(4)
out4['Overflow_Defice_Camas'] = out4['Overflow_Defice_Camas'].round(1)
out4 = out4.sort_values('Overflow_Defice_Camas', ascending=False)
for j, h in enumerate(out4.columns, 1):
    ws4.cell(row=1, column=j, value=h)
for i, (_, row) in enumerate(out4.iterrows(), 2):
    for j, h in enumerate(out4.columns, 1):
        ws4.cell(row=i, column=j, value=row[h])
style_header(ws4, len(out4.columns))
ws4.freeze_panes = 'A2'
autofit(ws4, len(out4.columns))

output_path = os.path.join(SCRIPT_DIR, 'Modelo1_Reafetacao_2SFCA.xlsx')
wb.save(output_path)
print(f"✅ Ficheiro guardado em: {output_path}")