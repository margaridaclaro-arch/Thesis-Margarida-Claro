#!/usr/bin/env python3
"""
MODELO DE COBERTURA HOSPITALAR
"""
import pandas as pd
import numpy as np
from math import radians, cos, sin, asin, sqrt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
import os
import json
import time
import requests

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
print(f"📁 Pasta: {SCRIPT_DIR}\n")

# ═══════════════════════════════════════════════════════════════════════════
# CONFIGURAÇÃO DO MOTOR DE ROTAS (OpenStreetMap via OSRM)
# ═══════════════════════════════════════════════════════════════════════════
# As distâncias/tempos deixam de ser em "linha reta" (haversine) e passam a
# ser calculados pela rede viária real, usando o OSRM (Open Source Routing
# Machine), que corre sobre dados do OpenStreetMap.
#
# Por omissão usa-se o servidor público de demonstração do projeto OSRM.
# É gratuito e não precisa de chave, mas tem uma política de "fair use" (sem
# SLA, pode ficar lento/instável com muitos milhares de pedidos seguidos).
#
# Para uso mais pesado/estável (todo o país, execuções frequentes), o
# recomendado é montar um servidor OSRM próprio com um extrato do
# OpenStreetMap de Portugal, por exemplo via Docker:
#
#   wget https://download.geofabrik.de/europe/portugal-latest.osm.pbf
#   docker run -t -v "${PWD}:/data" ghcr.io/project-osrm/osrm-backend osrm-extract -p /opt/car.lua /data/portugal-latest.osm.pbf
#   docker run -t -v "${PWD}:/data" ghcr.io/project-osrm/osrm-backend osrm-partition /data/portugal-latest.osrm
#   docker run -t -v "${PWD}:/data" ghcr.io/project-osrm/osrm-backend osrm-customize /data/portugal-latest.osrm
#   docker run -t -i -p 5000:5000 -v "${PWD}:/data" ghcr.io/project-osrm/osrm-backend osrm-routed --algorithm mld /data/portugal-latest.osrm
#
# E depois basta mudar a linha abaixo para:
#   OSRM_BASE_URL = "http://localhost:5000"
OSRM_BASE_URL = "https://router.project-osrm.org"
OSRM_BATCH_SIZE = 90        # nº máx. de freguesias por chamada (servidor demo limita a ~100 coordenadas, incl. o hospital)
OSRM_SLEEP_SECONDS = 1.0    # pausa entre chamadas, para não sobrecarregar o servidor demo
OSRM_CACHE_FILE = os.path.join(SCRIPT_DIR, '_osrm_cache.json')

# ═══════════════════════════════════════════════════════════════════════════
# CARREGAR FICHEIROS
# ═══════════════════════════════════════════════════════════════════════════
print("📂 Carregando Excel...")

try:
    df_oferta = pd.read_excel(
        os.path.join(SCRIPT_DIR, 'Excel_Oferta_Portugal.xlsx'),
        sheet_name='1. Oferta por ULS', header=2
    )
    print("✅ Oferta")
except Exception as e:
    print(f"❌ Oferta: {e}")
    exit(1)

try:
    # Excel_Procura_Portugal.xlsx: já vem com Internamentos Esperados e Dias
    # Necessários por Distrito/Concelho/Freguesia/Faixa Etária (folha
    # '2. Agregado por Faixa'), calculados a partir de probabilidades
    # nacionais x população da freguesia. É esta a fonte da Procura na folha
    # '4. Cobertura Procura x Oferta' (substitui a antiga estimativa via
    # Taxa de Internamento x População).
    df_procura = pd.read_excel(
        os.path.join(SCRIPT_DIR, 'Excel_Procura_Portugal.xlsx'),
        sheet_name='2. Agregado por Faixa', header=4
    )
    df_procura.columns = df_procura.columns.str.strip()
    print("✅ Procura (Internamentos Esperados por freguesia)")
except Exception as e:
    print(f"❌ Procura: {e}")
    exit(1)

try:
    df_freguesias = pd.read_excel(
        os.path.join(SCRIPT_DIR, 'freguesias_limpo.xlsx'),
        sheet_name=0
    )
    print("✅ Freguesias")
except Exception as e:
    print(f"❌ Freguesias: {e}")
    exit(1)

# ═══════════════════════════════════════════════════════════════════════════
# DETECTAR NOMES DE COLUNAS AUTOMATICAMENTE
# ═══════════════════════════════════════════════════════════════════════════
print("\n🔍 Detectando nomes de colunas...")

# Limpar espaços em nomes de colunas
df_freguesias.columns = df_freguesias.columns.str.strip()
df_oferta.columns = df_oferta.columns.str.strip()

# Mostrar nomes de colunas
print(f"   Colunas em freguesias: {list(df_freguesias.columns)}\n")


def find_column(df, keywords):
    """Procura coluna que contenha alguma keyword"""
    cols_lower = {col.lower(): col for col in df.columns}
    for keyword in keywords:
        for col_lower, col_orig in cols_lower.items():
            if keyword.lower() in col_lower:
                return col_orig
    return None


def find_column_all(df, must_contain_all):
    """Procura coluna cujo nome contenha TODAS as keywords dadas (evita
    confundir 'Lat_Freguesia' com 'Lat do Hospital', por exemplo)."""
    for col in df.columns:
        cl = col.lower()
        if all(k.lower() in cl for k in must_contain_all):
            return col
    return None


# Encontrar colunas de coordenadas
# NOTA: usa-se find_column_all (exige as duas palavras) para não confundir
# as coordenadas da freguesia com as do hospital, já que ambas contêm "lat"/"long".
lat_freq = find_column_all(df_freguesias, ['lat', 'freguesia']) or find_column(df_freguesias, ['lat', 'latitude'])
lon_freq = find_column_all(df_freguesias, ['long', 'freguesia']) or find_column(df_freguesias, ['long', 'longitude'])
lat_hosp = find_column_all(df_freguesias, ['lat', 'hospital'])
lon_hosp = find_column_all(df_freguesias, ['long', 'hospital'])
pop_col = find_column(df_freguesias, ['população', 'populacao', 'pop'])

# Coluna de capacidade (camas) na folha de Oferta - deteção automática, tal
# como as restantes colunas, porque o nome exato varia consoante a fonte.
camas_col = find_column(df_oferta, ['camas', 'cama', 'leito', 'leitos', 'lotação', 'lotacao', 'capacidade'])
uls_oferta_col = find_column(df_oferta, ['uls', 'instituição', 'instituicao']) or 'ULS'

# Colunas de identificação da freguesia (usadas no output detalhado)
distrito_col = find_column(df_freguesias, ['distrito'])
concelho_col = find_column(df_freguesias, ['concelho'])
freguesia_col = find_column(df_freguesias, ['freguesia e concelho', 'freguesia'])
hospital_col = find_column(df_freguesias, ['hosp'])


freguesia_pura_col = None
for _c in df_freguesias.columns:
    _cl = _c.lower()
    if 'freguesia' in _cl and 'concelho' not in _cl:
        freguesia_pura_col = _c
        break
if freguesia_pura_col is None:
    freguesia_pura_col = freguesia_col  # fallback (pode não casar bem)

# Colunas do Excel da Procura (folha '2. Agregado por Faixa')
distrito_procura_col = find_column(df_procura, ['distrito'])
concelho_procura_col = find_column(df_procura, ['concelho'])
freguesia_procura_col = find_column(df_procura, ['freguesia'])
internamentos_esp_col = find_column(df_procura, ['internamentos esperados', 'internamentos'])
dias_necessarios_col = find_column(df_procura, ['dias necess'])
camas_necessarias_col = find_column(df_procura, ['camas necess'])

if not all([distrito_procura_col, concelho_procura_col, freguesia_procura_col,
            internamentos_esp_col, dias_necessarios_col, camas_necessarias_col]):
    print("❌ Não consegui encontrar todas as colunas necessárias em Excel_Procura_Portugal.xlsx!")
    print(f"   Colunas disponíveis: {list(df_procura.columns)}")
    exit(1)

print(f"   Lat Freguesia: {lat_freq}")
print(f"   Lon Freguesia: {lon_freq}")
print(f"   Lat Hospital: {lat_hosp}")
print(f"   Lon Hospital: {lon_hosp}")
print(f"   População: {pop_col}")
print(f"   Distrito: {distrito_col} | Concelho: {concelho_col} | Freguesia: {freguesia_col} | Hospital: {hospital_col}")
print(f"   Camas (Oferta): {camas_col}\n")

# Validar
if not all([lat_freq, lon_freq, lat_hosp, lon_hosp, pop_col]):
    print("❌ Não consegui encontrar todas as colunas necessárias!")
    print(f"   Colunas disponíveis: {list(df_freguesias.columns)}")
    exit(1)


# ═══════════════════════════════════════════════════════════════════════════
# FUNÇÕES
# ═══════════════════════════════════════════════════════════════════════════
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
    freguesias até UM hospital, numa única chamada ao serviço 'table' do OSRM
    (muito mais eficiente do que uma chamada 'route' por freguesia).

    freg_rows: lista de tuplos (índice_no_df, lat, lon) das freguesias.
    Devolve: dict {índice_no_df: (distancia_km, tempo_min)}.
    Em caso de falha da chamada, ou de par sem rota (ex.: ilha sem ligação
    rodoviária ao hospital), cai em haversine como reserva.
    """
    resultados = {}

    # 1) Ver o que já está em cache
    por_pedir = []
    for idx, lat_f, lon_f in freg_rows:
        key = _osrm_cache_key(lat_f, lon_f, lat_h, lon_h)
        if key in _osrm_cache:
            resultados[idx] = tuple(_osrm_cache[key])
        else:
            por_pedir.append((idx, lat_f, lon_f, key))

    if not por_pedir:
        return resultados

    # 2) OSRM usa coordenadas no formato "longitude,latitude"
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
        distancias = data['distances']  # metros -> matriz [n origens][1 destino]
        duracoes = data['durations']    # segundos
        for i, (idx, lat_f, lon_f, key) in enumerate(por_pedir):
            dist_m = distancias[i][0]
            dur_s = duracoes[i][0]
            if dist_m is None or dur_s is None:
                # Sem rota rodoviária encontrada -> reserva: haversine
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


def normalize_uls(name):
    import re
    import unicodedata
    s = str(name).strip()
    s = re.sub(r',?\s*E\.?P\.?E\.?\.?$', '', s, flags=re.IGNORECASE)
    s = re.sub(r',?\s*PPP\.?$', '', s, flags=re.IGNORECASE)
    s = s.lower()
    s = unicodedata.normalize('NFKD', s).encode('ascii', 'ignore').decode()
    s = re.sub(r'[\/\-]', ' ', s)   # "Gaia/Espinho", "Almada-Seixal" -> "gaia espinho"
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def normalize_text(s):
    import re
    import unicodedata
    s = str(s).strip().lower()
    s = unicodedata.normalize('NFKD', s).encode('ascii', 'ignore').decode()
    s = re.sub(r'\s+', ' ', s)
    return s


# Alguns ULS têm nomes genuinamente diferentes entre ficheiros (não é só
# acentuação/maiúsculas) — ex.: "ULS Castelo Branco" vs "ULS de Castelo
# Branco", ou nome abreviado vs completo. Mapeia-se aqui manualmente para o
# mesmo ULS_norm depois da normalização automática.
ULS_ALIASES = {
    'unidade local de saude castelo branco': 'unidade local de saude de castelo branco',
    'unidade local de saude de vila nova de gaia espinho': 'unidade local de saude de gaia espinho',
    'unidade local de saude de dao lafoes': 'unidade local de saude de viseu dao lafoes',
}


def normalize_uls_final(name):
    key = normalize_uls(name)
    return ULS_ALIASES.get(key, key)


# ═══════════════════════════════════════════════════════════════════════════
# PROCESSAR
# ═══════════════════════════════════════════════════════════════════════════
print("🔄 Processando...")

# Normalizar ULS
df_freguesias['ULS_norm'] = df_freguesias['ULS'].apply(normalize_uls_final)
df_oferta['ULS_norm'] = df_oferta['ULS'].apply(normalize_uls_final)

# Mapear cada linha do Excel da Procura (por Distrito/Concelho/Freguesia) para
# o ULS respetivo, através da mesma chave (Distrito/Concelho/Freguesia) na
# folha de freguesias.
df_procura['_freg_key'] = (
    df_procura[distrito_procura_col].apply(normalize_text) + '|' +
    df_procura[concelho_procura_col].apply(normalize_text) + '|' +
    df_procura[freguesia_procura_col].apply(normalize_text)
)
df_freguesias['_freg_key'] = (
    df_freguesias[distrito_col].apply(normalize_text) + '|' +
    df_freguesias[concelho_col].apply(normalize_text) + '|' +
    df_freguesias[freguesia_pura_col].apply(normalize_text)
)
_freg_to_uls = df_freguesias.drop_duplicates('_freg_key').set_index('_freg_key')['ULS_norm']
df_procura['ULS_norm'] = df_procura['_freg_key'].map(_freg_to_uls)

_n_procura_total = len(df_procura)
_n_procura_sem_uls = int(df_procura['ULS_norm'].isna().sum())
if _n_procura_sem_uls:
    print(f"   ⚠️  {_n_procura_sem_uls}/{_n_procura_total} linhas de Excel_Procura_Portugal.xlsx "
          f"não casaram com nenhuma freguesia (Distrito/Concelho/Freguesia) -> excluídas da Procura por ULS.")

# Calcular distâncias REAIS POR ESTRADA (OpenStreetMap via OSRM), à FREGUESIA.
# Agrupa-se por hospital (mesmo par lat/lon) para poder pedir, numa única
# chamada, a distância de várias freguesias até esse hospital (serviço
# 'table' do OSRM), em vez de uma chamada por freguesia.
print("   Calculando distâncias reais por estrada (OpenStreetMap/OSRM)...")

df_freguesias['distancia_km'] = np.nan
df_freguesias['tempo_min'] = np.nan

_validos = df_freguesias.dropna(subset=[lat_freq, lon_freq, lat_hosp, lon_hosp])
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
            df_freguesias.loc[_idx, 'distancia_km'] = _dist_km
            df_freguesias.loc[_idx, 'tempo_min'] = _tempo_min

_osrm_save_cache()

df_freguesias['critico'] = df_freguesias['tempo_min'] > 60

# ═══════════════════════════════════════════════════════════════════════════
# KPIs
# ═══════════════════════════════════════════════════════════════════════════
print("   Calculando KPIs...")

# KPI 1 - Resumo por ULS (mantido para visão geral)
kpi1 = df_freguesias.groupby('ULS_norm').agg({
    'tempo_min': 'mean',
    'distancia_km': 'mean',
    pop_col: 'sum'
}).reset_index()
kpi1.columns = ['ULS', 'tempo_min', 'distancia_km', 'pop']
kpi1['tempo_min'] = kpi1['tempo_min'].round(1)
kpi1['distancia_km'] = kpi1['distancia_km'].round(1)

pop_crit = df_freguesias.groupby('ULS_norm').apply(
    lambda x: (x[x['critico']][pop_col].sum() / x[pop_col].sum() * 100)
    if x[pop_col].sum() > 0 else 0
).reset_index()
pop_crit.columns = ['ULS', 'pop_crit_pct']
pop_crit['pop_crit_pct'] = pop_crit['pop_crit_pct'].round(1)

kpi1 = kpi1.merge(pop_crit, on='ULS')

# Faixas etárias jovem/idoso (calculadas aqui em cima porque já são precisas
# para as colunas por freguesia, e mais abaixo para o KPI4 nacional)
faixas_jovem = ['[0-1[', '[1,5[', '[5,15[', '[15,25[']
faixas_idoso = ['[65,120[']
df_freguesias['pop_jovem'] = df_freguesias[faixas_jovem].sum(axis=1)
df_freguesias['pop_idoso'] = df_freguesias[faixas_idoso].sum(axis=1)

# Componentes "tempo x população" por freguesia (para agregares como quiseres,
# por ULS/Concelho/Distrito, numa tabela dinâmica: soma esta coluna e divide
# pela soma da população respetiva para obteres o tempo de acesso ponderado)
df_freguesias['TxPop_Jovem'] = df_freguesias['tempo_min'] * df_freguesias['pop_jovem']
df_freguesias['TxPop_Idoso'] = df_freguesias['tempo_min'] * df_freguesias['pop_idoso']
df_freguesias['TxPop_Total'] = df_freguesias['tempo_min'] * df_freguesias[pop_col]

# KPI 1B - Detalhe POR FREGUESIA (novo)
cols_detalhe = [c for c in [distrito_col, concelho_col, freguesia_col, 'ULS', hospital_col,
                             pop_col, 'distancia_km', 'tempo_min', 'critico',
                             'pop_jovem', 'pop_idoso',
                             'TxPop_Jovem', 'TxPop_Idoso', 'TxPop_Total'] if c is not None]
kpi_freguesia = df_freguesias[cols_detalhe].copy()
kpi_freguesia['distancia_km'] = kpi_freguesia['distancia_km'].round(1)
kpi_freguesia['tempo_min'] = kpi_freguesia['tempo_min'].round(1)
kpi_freguesia['critico'] = kpi_freguesia['critico'].map({True: 'Sim', False: 'Não'})
kpi_freguesia['TxPop_Jovem'] = kpi_freguesia['TxPop_Jovem'].round(1)
kpi_freguesia['TxPop_Idoso'] = kpi_freguesia['TxPop_Idoso'].round(1)
kpi_freguesia['TxPop_Total'] = kpi_freguesia['TxPop_Total'].round(1)
kpi_freguesia = kpi_freguesia.sort_values(
    [c for c in [distrito_col, concelho_col, freguesia_col] if c is not None]
)

# KPI 1C - % população coberta (<=60 min) por faixa etária, POR FREGUESIA (novo)
faixas = ['[0-1[', '[1,5[', '[5,15[', '[15,25[', '[25,45[', '[45,65[', '[65,120[']
id_cols = [c for c in [distrito_col, concelho_col, freguesia_col] if c is not None]
kpi_cobertura_freg = df_freguesias[id_cols + ['ULS', 'tempo_min']].copy()
for faixa in faixas:
    if faixa in df_freguesias.columns:
        kpi_cobertura_freg[f'Pop_{faixa}'] = df_freguesias[faixa]
        kpi_cobertura_freg[f'Cobertura_{faixa}_%'] = np.where(
            df_freguesias['tempo_min'] <= 60, 100.0, 0.0
        )
kpi_cobertura_freg['tempo_min'] = kpi_cobertura_freg['tempo_min'].round(1)
kpi_cobertura_freg = kpi_cobertura_freg.sort_values(id_cols)

# Nomes "bonitos" de ULS (para mostrar no output) e população por ULS
pop_por_uls = df_freguesias.groupby('ULS_norm')[pop_col].sum().reset_index()
pop_por_uls.columns = ['ULS_norm', 'Pop_ULS']

nome_uls_bonito = (
    df_freguesias.dropna(subset=['ULS'])
    .drop_duplicates('ULS_norm')
    .set_index('ULS_norm')['ULS']
)

# KPI 3
pop_nacional_total = df_freguesias[pop_col].sum()
kpi3_list = []
for faixa in faixas:
    if faixa in df_freguesias.columns:
        pop_tot = df_freguesias[faixa].sum()
        pop_ok = df_freguesias[df_freguesias['tempo_min'] <= 60][faixa].sum()
        pct_ok = (pop_ok / pop_tot * 100) if pop_tot > 0 else 0
        # Acesso ponderado (min) da faixa = Σ(Tempo_min × Pop_faixa) / Pop_total NACIONAL
        # (mesmo denominador em todas as faixas, para os valores serem diretamente comparáveis entre si)
        acesso_ponderado_faixa = (
            (df_freguesias['tempo_min'] * df_freguesias[faixa]).sum() / pop_nacional_total
        ) if pop_nacional_total > 0 else np.nan
        kpi3_list.append({
            'Faixa': faixa,
            'Pop_total': int(pop_tot),
            'Pop_cobertura': int(pop_ok),
            'Pct': round(pct_ok, 1),
            'Acesso_Ponderado_Min': round(acesso_ponderado_faixa, 2) if pd.notna(acesso_ponderado_faixa) else np.nan
        })
kpi3 = pd.DataFrame(kpi3_list)

# KPI 4 (agregado nacional - acesso ponderado por idade)
acesso_jovem = (df_freguesias['tempo_min'] * df_freguesias['pop_jovem']).sum() / df_freguesias['pop_jovem'].sum()
acesso_idoso = (df_freguesias['tempo_min'] * df_freguesias['pop_idoso']).sum() / df_freguesias['pop_idoso'].sum()

# Tempo médio de acesso PONDERADO POR POPULAÇÃO (nacional): cada freguesia pesa
# conforme a sua população -> responde a "quanto tempo espera, em média, um
# português". É a métrica correta para afirmações sobre acesso da população.
acesso_total = (df_freguesias['tempo_min'] * df_freguesias[pop_col]).sum() / df_freguesias[pop_col].sum()

# Tempo médio de acesso SIMPLES POR FREGUESIA (nacional): as 2.882 freguesias
# pesam todas o mesmo, independentemente da população -> responde a "qual é o
# tempo médio se tratarmos todo o território como igualmente relevante".
# NOTA: isto é DIFERENTE da antiga métrica 'kpi1[tempo_min].mean()' usada no
# dashboard, que era uma média das médias por ULS (39 ULS com peso igual),
# uma "dupla média" pouco interpretável. Esta aqui é a média direta sobre as
# 2.882 freguesias.
acesso_simples_freguesia = df_freguesias['tempo_min'].mean()
mediana_freguesia = df_freguesias['tempo_min'].median()

# Estatísticas rápidas sobre freguesias críticas
n_freguesias = len(df_freguesias)
n_criticas = int(df_freguesias['critico'].sum())
pct_freguesias_criticas = round(n_criticas / n_freguesias * 100, 1) if n_freguesias else 0

# ═══════════════════════════════════════════════════════════════════════════
# KPI 5 - COBERTURA REAL: CAMAS NECESSÁRIAS (procura) vs CAMAS (oferta)
# ═══════════════════════════════════════════════════════════════════════════
# Isto é diferente da "% população <=60min" (que só mede acesso geográfico).
# Aqui cruzamos, por ULS:
#   CAMAS NECESSÁRIAS = já vem calculada no Excel_Procura_Portugal.xlsx
#                        ('2. Agregado por Faixa'), por freguesia/faixa etária;
#                        agregamos por ULS através do mapeamento freguesia -> ULS.
#   CAMAS (OFERTA)     = nº de camas reais do ULS (Excel_Oferta_Portugal.xlsx).
# Diferença simples: Camas_Em_Falta = Camas_Necessarias - Camas (se positivo).
print("   Calculando cobertura Camas Necessárias x Camas (Oferta)...")

procura_uls = df_procura.dropna(subset=['ULS_norm']).groupby('ULS_norm').agg(
    Procura_Internamentos=(internamentos_esp_col, 'sum'),
    Camas_Necessarias=(camas_necessarias_col, 'sum'),
).reset_index()
procura_uls = procura_uls.merge(pop_por_uls, on='ULS_norm', how='left')

if camas_col:
    oferta_uls = df_oferta.groupby('ULS_norm')[camas_col].sum().reset_index()
    oferta_uls.columns = ['ULS_norm', 'Camas']
    kpi_cobertura_oferta = procura_uls.merge(oferta_uls, on='ULS_norm', how='left')
else:
    print("   ⚠️  Não encontrei coluna de camas em Excel_Oferta_Portugal.xlsx -> "
          "Cobertura Oferta ficará em branco. Confirma o nome da coluna e adiciona "
          "às keywords em 'camas_col' no início do script.")
    kpi_cobertura_oferta = procura_uls.copy()
    kpi_cobertura_oferta['Camas'] = np.nan

kpi_cobertura_oferta['ULS'] = kpi_cobertura_oferta['ULS_norm'].map(nome_uls_bonito)
kpi_cobertura_oferta['ULS'] = kpi_cobertura_oferta['ULS'].fillna(
    kpi_cobertura_oferta['ULS_norm'].str.title()
)

kpi_cobertura_oferta['Procura_Internamentos'] = kpi_cobertura_oferta['Procura_Internamentos'].round(0)
kpi_cobertura_oferta['Camas_Necessarias'] = kpi_cobertura_oferta['Camas_Necessarias'].round(1)

cobertura_raw = np.where(
    kpi_cobertura_oferta['Camas_Necessarias'] > 0,
    (kpi_cobertura_oferta['Camas'] / kpi_cobertura_oferta['Camas_Necessarias'] * 100),
    np.nan
)
kpi_cobertura_oferta['Cobertura_%'] = pd.Series(cobertura_raw, index=kpi_cobertura_oferta.index).clip(upper=100).round(1)
kpi_cobertura_oferta['Camas_Em_Falta'] = (
    kpi_cobertura_oferta['Camas_Necessarias'] - kpi_cobertura_oferta['Camas']
).clip(lower=0).round(1)

kpi_cobertura_oferta = kpi_cobertura_oferta[[
    'ULS', 'Pop_ULS', 'Procura_Internamentos', 'Camas_Necessarias', 'Camas',
    'Cobertura_%', 'Camas_Em_Falta'
]].sort_values('Cobertura_%', na_position='last')

# ═══════════════════════════════════════════════════════════════════════════
# EXCEL
# ═══════════════════════════════════════════════════════════════════════════
print("   Criando Excel...")

wb = Workbook()
wb.remove(wb.active)

# Placeholder do Dashboard (fica em 1º lugar); é preenchido no fim do script,
# depois de todos os KPIs de todas as outras folhas estarem calculados.
ws_dash = wb.create_sheet('📊 Dashboard', 0)

header_font = Font(bold=True, color="FFFFFF")
header_fill = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
critico_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")


def style_header(ws, ncols, row=1):
    for col in range(1, ncols + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")


def autofit(ws, ncols, min_width=10, max_width=45):
    for col in range(1, ncols + 1):
        letter = get_column_letter(col)
        max_len = min_width
        for cell in ws[letter]:
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[letter].width = min(max_len + 2, max_width)


# Folha nova: detalhe por freguesia
ws0 = wb.create_sheet('1. Por Freguesia', 1)
ws0['A1'] = ("Distancia_km = distância em linha reta (Haversine) freguesia -> hospital do ULS  |  "
             "Tempo_min = Distancia_km / 1 km/min + 15 min fixos (acesso/estacionamento/triagem)  |  "
             "Crítico = Tempo_min > 60  |  "
             "Produto_Tempo_x_Pop* = Tempo_min × População (NÃO é tempo médio — para obter o tempo médio "
             "ponderado de um conjunto de freguesias, soma esta coluna e divide pela soma da população respetiva)")
ws0['A1'].font = Font(italic=True, size=9, color="666666")
headers0 = []
if distrito_col: headers0.append('Distrito')
if concelho_col: headers0.append('Concelho')
if freguesia_col: headers0.append('Freguesia')
headers0.append('ULS')
if hospital_col: headers0.append('Hospital')
headers0 += ['População', 'Distancia_km', 'Tempo_min', 'Crítico (>60min)',
             'Pop_Jovem_0-25', 'Pop_Idoso_65+',
             'Produto_Tempo_x_PopJovem', 'Produto_Tempo_x_PopIdoso', 'Produto_Tempo_x_PopTotal']

for j, h in enumerate(headers0, 1):
    ws0.cell(row=2, column=j, value=h)

# Posições (letras de coluna) usadas para escrever FÓRMULAS reais do Excel
# em vez de valores estáticos - assim consegues auditar/recalcular no Excel.
col_tempo_letra = get_column_letter(headers0.index('Tempo_min') + 1)
col_popjov_letra = get_column_letter(headers0.index('Pop_Jovem_0-25') + 1)
col_popid_letra = get_column_letter(headers0.index('Pop_Idoso_65+') + 1)
col_pop_letra = get_column_letter(headers0.index('População') + 1)

for i, (_, data) in enumerate(kpi_freguesia.iterrows(), 3):
    col = 1
    if distrito_col:
        ws0.cell(row=i, column=col, value=data[distrito_col]); col += 1
    if concelho_col:
        ws0.cell(row=i, column=col, value=data[concelho_col]); col += 1
    if freguesia_col:
        ws0.cell(row=i, column=col, value=data[freguesia_col]); col += 1
    ws0.cell(row=i, column=col, value=data['ULS']); col += 1
    if hospital_col:
        ws0.cell(row=i, column=col, value=data[hospital_col]); col += 1
    ws0.cell(row=i, column=col, value=data[pop_col]); col += 1
    ws0.cell(row=i, column=col, value=data['distancia_km']); col += 1
    ws0.cell(row=i, column=col, value=data['tempo_min']); col += 1
    crit_col = col
    ws0.cell(row=i, column=col, value=data['critico']); col += 1
    ws0.cell(row=i, column=col, value=data['pop_jovem']); col += 1
    ws0.cell(row=i, column=col, value=data['pop_idoso']); col += 1
    # Produto_Tempo_x_Pop* = Tempo_min x População da faixa -> FÓRMULA viva no Excel
    ws0.cell(row=i, column=col, value=f"={col_tempo_letra}{i}*{col_popjov_letra}{i}"); col += 1
    ws0.cell(row=i, column=col, value=f"={col_tempo_letra}{i}*{col_popid_letra}{i}"); col += 1
    ws0.cell(row=i, column=col, value=f"={col_tempo_letra}{i}*{col_pop_letra}{i}")
    if data['critico'] == 'Sim':
        for c in range(1, col + 1):
            ws0.cell(row=i, column=c).fill = critico_fill

style_header(ws0, len(headers0), row=2)
ws0.freeze_panes = 'A3'
autofit(ws0, len(headers0))

# Folha nova: % população coberta (<=60 min) por faixa etária, uma linha por freguesia
ws0b = wb.create_sheet('2. Cobertura Etária (Freg)', 2)
ws0b['A1'] = ("Cobertura_[faixa]_% = 100 se Tempo_min <= 60, senão 0 (por freguesia)  |  "
              "Depende só do Tempo_min da freguesia, por isso é IGUAL em todas as faixas etárias da mesma linha "
              "— não há um limiar de tempo diferente por idade neste modelo.")
ws0b['A1'].font = Font(italic=True, size=9, color="666666")
headers0b = list(kpi_cobertura_freg.columns)
for j, h in enumerate(headers0b, 1):
    ws0b.cell(row=2, column=j, value=h)
for i, (_, data) in enumerate(kpi_cobertura_freg.iterrows(), 3):
    for j, h in enumerate(headers0b, 1):
        ws0b.cell(row=i, column=j, value=data[h])
style_header(ws0b, len(headers0b), row=2)
ws0b.freeze_panes = 'A3'
autofit(ws0b, len(headers0b))

# Folha 4 - COBERTURA: Camas Necessárias (procura) vs Camas (oferta)
ws4 = wb.create_sheet('3. Cobertura Procura x Oferta')
ws4['A1'] = (f"Camas_Necessarias = já vem calculado no Excel_Procura_Portugal.xlsx (folha '2. Agregado por Faixa'), "
             f"agregado por ULS via mapeamento freguesia -> ULS  |  "
             f"Cobertura% = min(100, Camas / Camas_Necessarias × 100)  |  "
             f"Camas_Em_Falta = max(0, Camas_Necessarias - Camas)")
ws4['A1'].font = Font(italic=True, size=9, color="666666")
headers4 = list(kpi_cobertura_oferta.columns)
for j, h in enumerate(headers4, 1):
    ws4.cell(row=2, column=j, value=h)
col_necessarias_letra = get_column_letter(headers4.index('Camas_Necessarias') + 1)
col_camas_letra = get_column_letter(headers4.index('Camas') + 1)
cobertura_fill_baixa = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
for i, row_data in enumerate(kpi_cobertura_oferta.iterrows(), 3):
    data = row_data[1]
    ws4.cell(row=i, column=headers4.index('ULS') + 1, value=data['ULS'])
    ws4.cell(row=i, column=headers4.index('Pop_ULS') + 1, value=data['Pop_ULS'] if pd.notna(data['Pop_ULS']) else 'N/D')
    ws4.cell(row=i, column=headers4.index('Procura_Internamentos') + 1, value=data['Procura_Internamentos'])
    ws4.cell(row=i, column=headers4.index('Camas_Necessarias') + 1, value=data['Camas_Necessarias'])
    ws4.cell(row=i, column=headers4.index('Camas') + 1, value=data['Camas'] if pd.notna(data['Camas']) else 'N/D')
    if pd.notna(data['Camas']):
        # Cobertura% = MIN(100, Camas/Camas_Necessarias*100) ; Camas_Em_Falta = MAX(0, Camas_Necessarias-Camas)
        ws4.cell(row=i, column=headers4.index('Cobertura_%') + 1,
                 value=f"=MIN(100,{col_camas_letra}{i}/{col_necessarias_letra}{i}*100)")
        ws4.cell(row=i, column=headers4.index('Camas_Em_Falta') + 1,
                 value=f"=MAX(0,{col_necessarias_letra}{i}-{col_camas_letra}{i})")
        cob = data['Cobertura_%']
        if pd.notna(cob) and cob < 100:
            for c in range(1, len(headers4) + 1):
                ws4.cell(row=i, column=c).fill = cobertura_fill_baixa
    else:
        ws4.cell(row=i, column=headers4.index('Cobertura_%') + 1, value='N/D')
        ws4.cell(row=i, column=headers4.index('Camas_Em_Falta') + 1, value='N/D')
style_header(ws4, len(headers4), row=2)
ws4.freeze_panes = 'A3'
autofit(ws4, len(headers4))


# Folha auxiliar OCULTA - guarda apenas os dados que alimentam os gráficos do
# Dashboard (Faixa Etária e Top 10 ULS). Não é uma página visível: substitui a
# antiga página '3. Cobertura Faixa Etária' e a tabela de números que antes
# ficava "escondida" dentro da própria folha do Dashboard.
ws_aux = wb.create_sheet('_ChartsData')
ws_aux.sheet_state = 'hidden'
ws_aux['A2'] = 'Faixa'
ws_aux['B2'] = 'Pop_total'
ws_aux['C2'] = 'Pop_cobertura'
ws_aux['D2'] = 'Pct'
ws_aux['E2'] = 'Acesso_Ponderado_Min'
for i, row_data in enumerate(kpi3.iterrows(), 3):
    data = row_data[1]
    ws_aux[f'A{i}'] = data['Faixa']
    ws_aux[f'B{i}'] = data['Pop_total']
    ws_aux[f'C{i}'] = data['Pop_cobertura']
    ws_aux[f'D{i}'] = data['Pct']
    ws_aux[f'E{i}'] = data['Acesso_Ponderado_Min']

# ═══════════════════════════════════════════════════════════════════════════
# DASHBOARD (folha '📊 Dashboard') - resumo visual, KPIs + gráficos
# ═══════════════════════════════════════════════════════════════════════════
ws_dash.sheet_view.showGridLines = False
ws_dash.column_dimensions['A'].width = 2
for col in 'BCDEFGHIJK':
    ws_dash.column_dimensions[col].width = 16

ws_dash['B2'] = 'DASHBOARD — COBERTURA HOSPITALAR'
ws_dash['B2'].font = Font(bold=True, size=16, color="2E75B6")

# --- Cartões de KPI (linhas 4-8) -------------------------------------------
# NOTA: o "Tempo médio de acesso" passou a usar a média PONDERADA POR
# POPULAÇÃO (acesso_total), e não a antiga 'kpi1[tempo_min].mean()' (média
# das médias por ULS, ~27.3 min) — essa métrica dava peso igual a ULS muito
# diferentes em tamanho/população e não é interpretável como "tempo médio
# nacional". A média simples por freguesia (peso igual, sem ponderação) fica
# disponível como card separado, para contraste/dispersão territorial.
kpi_cards = [
    ('Tempo médio de acesso (ponderado pop.)', f"{round(acesso_total, 1)} min", "2E75B6"),
    ('Tempo médio simples (por freguesia)', f"{round(acesso_simples_freguesia, 1)} min", "808080"),
    ('População coberta (<=60min)', f"{round(kpi3['Pct'].mean(), 1)}%", "2E75B6"),
    ('Freguesias críticas', f"{n_criticas}/{n_freguesias} ({pct_freguesias_criticas}%)", "C00000"),
    ('Cobertura Camas (nacional)', (f"{round(kpi_cobertura_oferta['Cobertura_%'].mean(), 1)}%"
                                     if kpi_cobertura_oferta['Cobertura_%'].notna().any() else 'N/D'), "2E75B6"),
]
card_cols = ['B', 'D', 'F', 'H', 'J']
for (titulo, valor, cor), col in zip(kpi_cards, card_cols):
    col2 = chr(ord(col) + 1)
    ws_dash.merge_cells(f'{col}4:{col2}4')
    ws_dash.merge_cells(f'{col}5:{col2}6')
    c_titulo = ws_dash[f'{col}4']
    c_titulo.value = titulo
    c_titulo.font = Font(bold=True, size=10, color="666666")
    c_titulo.alignment = Alignment(horizontal='center')
    c_valor = ws_dash[f'{col}5']
    c_valor.value = valor
    c_valor.font = Font(bold=True, size=20, color=cor)
    c_valor.alignment = Alignment(horizontal='center', vertical='center')
    for row in (4, 5, 6):
        for c in (col, col2):
            ws_dash[f'{c}{row}'].fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

# --- Gráfico 1: Acesso Ponderado por Faixa Etária (dados da folha auxiliar) -------
chart1 = BarChart()
chart1.type = "col"
chart1.title = "Acesso Ponderado (min) por Faixa Etária"
chart1.y_axis.title = "Minutos"
chart1.x_axis.title = "Faixa Etária"
chart1.style = 10
data1 = Reference(ws_aux, min_col=5, min_row=2, max_row=2 + len(kpi3))
cats1 = Reference(ws_aux, min_col=1, min_row=3, max_row=2 + len(kpi3))
chart1.add_data(data1, titles_from_data=True)
chart1.set_categories(cats1)
chart1.height, chart1.width = 8, 16
chart1.legend.position = 'b'
ws_dash.add_chart(chart1, 'B9')

# --- Gráfico 2: % Cobertura por Faixa Etária (dados da folha auxiliar) -----------
chart2 = BarChart()
chart2.type = "col"
chart2.title = "% População Coberta (<=60min) por Faixa Etária"
chart2.y_axis.title = "%"
chart2.style = 11
data2 = Reference(ws_aux, min_col=4, min_row=2, max_row=2 + len(kpi3))
cats2 = Reference(ws_aux, min_col=1, min_row=3, max_row=2 + len(kpi3))
chart2.add_data(data2, titles_from_data=True)
chart2.set_categories(cats2)
chart2.height, chart2.width = 8, 16
chart2.legend.position = 'b'
ws_dash.add_chart(chart2, 'B25')

# --- Gráfico 3: Top 10 piores ULS por Cobertura de Camas (folha 3) --------
_top10 = kpi_cobertura_oferta[kpi_cobertura_oferta['Cobertura_%'].notna()].nsmallest(10, 'Cobertura_%')
_n_top10 = len(_top10)
if _n_top10 > 0:
    _start_row_aux = 20  # zona da folha auxiliar oculta, só para alimentar o gráfico
    ws_aux.cell(row=_start_row_aux, column=1, value='ULS')
    ws_aux.cell(row=_start_row_aux, column=2, value='Cobertura_%')
    for _i, (_, _r) in enumerate(_top10.iterrows(), 1):
        ws_aux.cell(row=_start_row_aux + _i, column=1, value=_r['ULS'])
        ws_aux.cell(row=_start_row_aux + _i, column=2, value=_r['Cobertura_%'])
    chart3 = BarChart()
    chart3.type = "bar"  # barras horizontais (nomes de ULS mais legíveis)
    chart3.title = f"Top {_n_top10} ULS com Menor Cobertura de Camas (%)"
    chart3.x_axis.title = "Cobertura %"
    chart3.style = 12
    data3 = Reference(ws_aux, min_col=2, min_row=_start_row_aux, max_row=_start_row_aux + _n_top10)
    cats3 = Reference(ws_aux, min_col=1, min_row=_start_row_aux + 1, max_row=_start_row_aux + _n_top10)
    chart3.add_data(data3, titles_from_data=True)
    chart3.set_categories(cats3)
    chart3.height, chart3.width = 10, 16
    chart3.legend.position = 'b'
    ws_dash.add_chart(chart3, 'B41')

output = os.path.join(SCRIPT_DIR, 'KPI_Cobertura_Hospitalar_FINAL.xlsx')
wb.save(output)

# ═══════════════════════════════════════════════════════════════════════════
# FIM
# ═══════════════════════════════════════════════════════════════════════════
print("\n" + "=" * 70)
print("✅ CONCLUÍDO!")
print("=" * 70)
print(f"\n📊 KPIs:")
print(f"   Tempo médio (ponderado por população): {round(acesso_total, 1)} min")
print(f"   Tempo médio (simples, por freguesia):   {round(acesso_simples_freguesia, 1)} min")
print(f"   Mediana (por freguesia):                {round(mediana_freguesia, 1)} min")
print(f"   Jovens: {round(acesso_jovem, 1)} min | Idosos: {round(acesso_idoso, 1)} min")
print(f"   Cobertura: {round(kpi3['Pct'].mean(), 1)}%")
print(f"   Freguesias críticas: {n_criticas}/{n_freguesias} ({pct_freguesias_criticas}%)")
print(f"   Internamentos esperados (Procura): {int(kpi_cobertura_oferta['Procura_Internamentos'].sum())}")
print(f"   Camas necessárias (total nacional): {round(kpi_cobertura_oferta['Camas_Necessarias'].sum(), 1)}")
print(f"   Cobertura Camas Necessárias x Camas (Oferta) média: {round(kpi_cobertura_oferta['Cobertura_%'].mean(), 1) if kpi_cobertura_oferta['Cobertura_%'].notna().any() else 'N/D'}%")
print(f"\n📁 Ficheiro: {output}")
print("=" * 70)