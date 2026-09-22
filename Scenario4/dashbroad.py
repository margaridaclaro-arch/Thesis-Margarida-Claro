#!/usr/bin/env python3
"""
MODELO DE IMPACTO DE ESTRADAS CORTADAS NO ACESSO HOSPITALAR - REGIÃO DE LEIRIA
"""
import pandas as pd
import numpy as np
from math import radians, cos, sin, asin, sqrt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
import os
import json
import time
import requests

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
print(f"📁 Pasta: {SCRIPT_DIR}\n")

# ═══════════════════════════════════════════════════════════════════════════
# CONFIGURAÇÃO
# ═══════════════════════════════════════════════════════════════════════════
OSRM_BASE_URL = "https://router.project-osrm.org"
OSRM_BATCH_SIZE = 90
OSRM_SLEEP_SECONDS = 1.0
OSRM_CACHE_FILE_NORMAL = os.path.join(SCRIPT_DIR, '_osrm_cache.json')                 # partilhada com o modelo.py original
OSRM_CACHE_FILE_CORTADO = os.path.join(SCRIPT_DIR, '_osrm_cache_estradas_cortadas.json')

# Âmbito geográfico do cenário: por omissão, o distrito inteiro de Leiria
# (foi o mais afetado, segundo a Proteção Civil e a CIM de Leiria). Podes
# reduzir esta lista aos concelhos que te interessam especificamente.
DISTRITO_ALVO = 'Leiria'
CONCELHOS_ALVO = None 

LIMIAR_CRITICO_MIN = 60  # tempo (min) a partir do qual uma freguesia é "crítica"

# Penalizações manuais (minutos extra) para captar cortes em estradas
# secundárias que o OSRM não sabe que estavam cortadas (EN109, EN236, IC2,
# etc.). Chave = nome da Freguesia OU do Concelho (tal como aparece em
# freguesias_limpo.xlsx); se puseres um Concelho, aplica-se a todas as
# freguesias desse concelho.
# Estes valores NÃO são medições reais (a notícia não dá minutos de atraso,
# só descreve "cortada" ou "com dificuldades de circulação"). São uma
# ESTIMATIVA 
PENALIZACOES_MANUAIS = {
    # EN236 (Lousã <-> Castanheira de Pêra) "cortada ao trânsito" -- corte total.
    # Nota: esta estrada liga sobretudo a Lousã (distrito de Coimbra), pelo que
    # pode não ser a via principal de acesso ao Hospital de Leiria a partir de
    # Castanheira de Pêra -- mesmo assim, sinaliza-se por precaução.
    'Castanheira de Pêra': 30,

    # IC2, "zona da Roca" (Leiria) "fortes constrangimentos" -- via condicionada,
    # não um corte total. Afeta sobretudo o próprio concelho de Leiria, por onde
    # passa o IC2 a caminho do Hospital de Santo André.
    'Leiria': 20,

    # EN109 (Leiria <-> Figueira da Foz) "com dificuldades de circulação".
    # Marinha Grande fica entre Leiria e o litoral, na zona de influência desta via.
    'Marinha Grande': 20,
}

# ATRITO GERAL: além dos cortes específicos acima, houve MUITAS vias que
# nunca foram formalmente "cortadas" mas ficaram parcialmente obstruídas
# (árvores caídas, ramos, detritos, água na faixa de rodagem) -- dava para
# passar, mas mais devagar. A Proteção Civil falou em "queda de árvores, que
# criou obstáculos nas estradas" como um dos maiores problemas em todo o
# país, e a CIM de Leiria descreveu concretamente os concelhos de Ansião,
# Batalha, Leiria, Marinha Grande, Porto de Mós e Pombal como os "muito
# devastados" da região.
#
# Aplicam-se DOIS níveis (também são estimativas editoriais, não medições):
#   1) ATRITO_GERAL_DISTRITO_MIN -- um "imposto" mínimo a TODO o distrito de
#      Leiria (esteve todo sob aviso vermelho/laranja nesse dia).
#   2) CONCELHOS_MUITO_AFETADOS -- minutos extra SÓ nos concelhos que a CIM
#      de Leiria confirmou como mais devastados (fica em cima do nº 1).
ATRITO_GERAL_DISTRITO_MIN = 10  # aplica-se a TODAS as freguesias do distrito
CONCELHOS_MUITO_AFETADOS = {
    # concelho: minutos extra de atrito, além do ATRITO_GERAL_DISTRITO_MIN
    'Ansião': 10,
    'Batalha': 10,
    'Leiria': 10,
    'Marinha Grande': 10,
    'Porto de Mós': 10,
    'Pombal': 10,
}

# ═══════════════════════════════════════════════════════════════════════════
# CARREGAR FICHEIROS (mesmos ficheiros do modelo base)
# ═══════════════════════════════════════════════════════════════════════════
print("📂 Carregando Excel...")
try:
    df_freguesias_full = pd.read_excel(
        os.path.join(SCRIPT_DIR, 'freguesias_limpo.xlsx'), sheet_name=0
    )
    print("✅ Freguesias")
except Exception as e:
    print(f"❌ Freguesias: {e}")
    exit(1)

df_freguesias_full.columns = df_freguesias_full.columns.str.strip()


def find_column(df, keywords):
    cols_lower = {col.lower(): col for col in df.columns}
    for keyword in keywords:
        for col_lower, col_orig in cols_lower.items():
            if keyword.lower() in col_lower:
                return col_orig
    return None


def find_column_all(df, must_contain_all):
    for col in df.columns:
        cl = col.lower()
        if all(k.lower() in cl for k in must_contain_all):
            return col
    return None


lat_freq = find_column_all(df_freguesias_full, ['lat', 'freguesia']) or find_column(df_freguesias_full, ['lat', 'latitude'])
lon_freq = find_column_all(df_freguesias_full, ['long', 'freguesia']) or find_column(df_freguesias_full, ['long', 'longitude'])
lat_hosp = find_column_all(df_freguesias_full, ['lat', 'hospital'])
lon_hosp = find_column_all(df_freguesias_full, ['long', 'hospital'])
pop_col = find_column(df_freguesias_full, ['população', 'populacao', 'pop'])
distrito_col = find_column(df_freguesias_full, ['distrito'])
concelho_col = find_column(df_freguesias_full, ['concelho'])
freguesia_col = find_column(df_freguesias_full, ['freguesia e concelho', 'freguesia'])
hospital_col = find_column(df_freguesias_full, ['hosp'])

if not all([lat_freq, lon_freq, lat_hosp, lon_hosp, pop_col, distrito_col, concelho_col]):
    print("❌ Não consegui encontrar todas as colunas necessárias em freguesias_limpo.xlsx!")
    print(f"   Colunas disponíveis: {list(df_freguesias_full.columns)}")
    exit(1)

# Freguesia "pura" (sem sufixo de concelho), usada para aplicar penalizações manuais
freguesia_pura_col = None
for _c in df_freguesias_full.columns:
    _cl = _c.lower()
    if 'freguesia' in _cl and 'concelho' not in _cl:
        freguesia_pura_col = _c
        break
if freguesia_pura_col is None:
    freguesia_pura_col = freguesia_col

# ═══════════════════════════════════════════════════════════════════════════
# FILTRAR AO ÂMBITO DO CENÁRIO (distrito/concelhos alvo)
# ═══════════════════════════════════════════════════════════════════════════
df_freguesias = df_freguesias_full[
    df_freguesias_full[distrito_col].astype(str).str.strip().str.lower() == DISTRITO_ALVO.strip().lower()
].copy()
if CONCELHOS_ALVO:
    alvo_lower = [c.strip().lower() for c in CONCELHOS_ALVO]
    df_freguesias = df_freguesias[
        df_freguesias[concelho_col].astype(str).str.strip().str.lower().isin(alvo_lower)
    ].copy()

df_freguesias = df_freguesias.reset_index(drop=True)
print(f"   Âmbito: distrito '{DISTRITO_ALVO}'" +
      (f", concelhos {CONCELHOS_ALVO}" if CONCELHOS_ALVO else " (todos os concelhos)") +
      f" -> {len(df_freguesias)} freguesias\n")

if len(df_freguesias) == 0:
    print("❌ Nenhuma freguesia encontrada para o âmbito definido. Confirma DISTRITO_ALVO/CONCELHOS_ALVO.")
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


def _load_cache(path):
    if os.path.exists(path):
        try:
            with open(path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def _save_cache(path, cache):
    try:
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(cache, f)
    except Exception as e:
        print(f"   ⚠️  Não consegui gravar cache ({path}): {e}")


def _cache_key(lat_f, lon_f, lat_h, lon_h):
    return f"{round(lat_f, 5)},{round(lon_f, 5)}|{round(lat_h, 5)},{round(lon_h, 5)}"


def osrm_table_batch(freg_rows, lat_h, lon_h, cache, extra_params=""):
    """Igual à função do modelo.py, mas parametrizável com `extra_params`
    (ex.: '&exclude=motorway') para simular estradas cortadas, e com cache
    própria por cenário (para não misturar tempos normais com cortados)."""
    resultados = {}
    por_pedir = []
    for idx, lat_f, lon_f in freg_rows:
        key = _cache_key(lat_f, lon_f, lat_h, lon_h)
        if key in cache:
            cached = cache[key]
            if len(cached) == 2:  # cache antiga (de antes de rastrearmos o método)
                cached = [cached[0], cached[1], 'osrm']
            resultados[idx] = tuple(cached)
        else:
            por_pedir.append((idx, lat_f, lon_f, key))

    if not por_pedir:
        return resultados

    coords = [f"{lon_f},{lat_f}" for _, lat_f, lon_f, _ in por_pedir]
    coords.append(f"{lon_h},{lat_h}")
    n = len(por_pedir)
    sources = ";".join(str(i) for i in range(n))
    destination = str(n)

    url = (
        f"{OSRM_BASE_URL}/table/v1/driving/{';'.join(coords)}"
        f"?sources={sources}&destinations={destination}&annotations=distance,duration{extra_params}"
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
                # OSRM não encontrou rota para este par específico (ex.: só é
                # possível chegar por autoestrada e esta foi excluída) -> reserva
                # por linha reta. MARCADO como 'haversine' porque a linha reta é
                # sempre mais curta que a estrada real e NÃO deve ser comparada
                # diretamente ao tempo do cenário Normal (ver clamp mais abaixo).
                dist_km = haversine(lat_f, lon_f, lat_h, lon_h)
                tempo_min = dist_km / 1.0
                metodo = 'haversine'
            else:
                dist_km = dist_m / 1000.0
                tempo_min = dur_s / 60.0
                metodo = 'osrm'
            resultados[idx] = (dist_km, tempo_min, metodo)
            cache[key] = [dist_km, tempo_min, metodo]
    except Exception as e:
        print(f"   ⚠️  Erro OSRM neste lote (extra_params='{extra_params}', a usar haversine como reserva): {e}")
        for idx, lat_f, lon_f, key in por_pedir:
            dist_km = haversine(lat_f, lon_f, lat_h, lon_h)
            tempo_min = dist_km / 1.0
            resultados[idx] = (dist_km, tempo_min, 'haversine')
            cache[key] = [dist_km, tempo_min, 'haversine']

    time.sleep(OSRM_SLEEP_SECONDS)
    return resultados


def calcular_tempos(df, cache_path, extra_params, label):
    """Calcula distancia_km/tempo_min para todas as freguesias de `df`,
    agrupando por hospital (mesmo par lat/lon) para OSRM 'table'."""
    print(f"   Calculando cenário '{label}' (OSRM {extra_params or '(rede completa)'})...")
    cache = _load_cache(cache_path)
    df = df.copy()
    df['distancia_km'] = np.nan
    df['tempo_min'] = np.nan
    df['metodo'] = ''

    validos = df.dropna(subset=[lat_freq, lon_freq, lat_hosp, lon_hosp])
    hospitais_unicos = validos[[lat_hosp, lon_hosp]].drop_duplicates()
    n_hosp = len(hospitais_unicos)

    for h_i, (_, hrow) in enumerate(hospitais_unicos.iterrows(), 1):
        lat_h, lon_h = hrow[lat_hosp], hrow[lon_hosp]
        subset = validos[(validos[lat_hosp] == lat_h) & (validos[lon_hosp] == lon_h)]
        print(f"      Hospital {h_i}/{n_hosp}: {len(subset)} freguesia(s)...")
        freg_rows = list(zip(subset.index, subset[lat_freq], subset[lon_freq]))
        for b in range(0, len(freg_rows), OSRM_BATCH_SIZE):
            lote = freg_rows[b:b + OSRM_BATCH_SIZE]
            res = osrm_table_batch(lote, lat_h, lon_h, cache, extra_params=extra_params)
            for idx, (dist_km, tempo_min, metodo) in res.items():
                df.loc[idx, 'distancia_km'] = dist_km
                df.loc[idx, 'tempo_min'] = tempo_min
                df.loc[idx, 'metodo'] = metodo

    _save_cache(cache_path, cache)
    return df


def aplicar_penalizacoes_manuais(df):
    """Soma minutos extra a df['tempo_min']:
      1) PENALIZACOES_MANUAIS -- cortes/condicionamentos numa via nomeada específica.
      2) ATRITO_GERAL_DISTRITO_MIN -- imposto mínimo a todo o distrito (árvores/detritos
         dispersos, sem corte formal).
      3) CONCELHOS_MUITO_AFETADOS -- atrito extra nos concelhos confirmados como mais
         devastados pela CIM de Leiria.
    Guarda cada componente em colunas separadas, para auditares de onde vem cada minuto."""
    df = df.copy()

    df['penalizacao_corte_especifico_min'] = 0.0
    for chave, minutos in PENALIZACOES_MANUAIS.items():
        mask_freg = df[freguesia_pura_col].astype(str).str.strip().str.lower() == str(chave).strip().lower()
        mask_conc = df[concelho_col].astype(str).str.strip().str.lower() == str(chave).strip().lower()
        mask = mask_freg | mask_conc
        df.loc[mask, 'penalizacao_corte_especifico_min'] += minutos

    df['penalizacao_atrito_geral_min'] = float(ATRITO_GERAL_DISTRITO_MIN)
    for concelho, minutos in CONCELHOS_MUITO_AFETADOS.items():
        mask_conc = df[concelho_col].astype(str).str.strip().str.lower() == str(concelho).strip().lower()
        df.loc[mask_conc, 'penalizacao_atrito_geral_min'] += minutos

    df['penalizacao_manual_min'] = df['penalizacao_corte_especifico_min'] + df['penalizacao_atrito_geral_min']
    df['tempo_min'] = df['tempo_min'] + df['penalizacao_manual_min']
    return df


def corrigir_tempos_invalidos(df_normal, df_cortado):
    """Um corte/condicionamento de estradas NUNCA pode tornar um trajeto
    mais RÁPIDO do que em condições normais. Se isso acontecer no cálculo
    (normalmente porque o OSRM não encontrou rota sem autoestrada para
    aquele par e caiu para a reserva por linha reta -- ver 'metodo' ==
    'haversine' -- que é sempre mais curta que a estrada real), força o
    tempo cortado a ser, no mínimo, igual ao normal, e marca a linha como
    'Ajustado_Sem_Rota' para ficares a saber que aquele valor não veio de
    uma rota OSRM real e deve ser lido com reserva."""
    df_cortado = df_cortado.copy()
    tempo_normal = df_normal['tempo_min'].values
    invalido = (df_cortado['tempo_min'].values < tempo_normal) | (df_cortado['metodo'].values == 'haversine')
    df_cortado['ajustado_sem_rota'] = invalido
    df_cortado['tempo_min'] = np.maximum(df_cortado['tempo_min'].values, tempo_normal)
    df_cortado['distancia_km'] = np.maximum(df_cortado['distancia_km'].values, df_normal['distancia_km'].values)
    n_ajustados = int(invalido.sum())
    if n_ajustados:
        print(f"   ⚠️  {n_ajustados} freguesia(s) sem rota fiável para o cenário 'cortado' "
              f"(OSRM caiu para linha reta) -- tempo ajustado para não ficar abaixo do normal. "
              f"Ver coluna 'Ajustado_Sem_Rota' na folha de comparação.")
    return df_cortado


# ═══════════════════════════════════════════════════════════════════════════
# PROCESSAR: DOIS CENÁRIOS
# ═══════════════════════════════════════════════════════════════════════════
print("🔄 Processando cenários...")

df_normal = calcular_tempos(
    df_freguesias, OSRM_CACHE_FILE_NORMAL, extra_params="", label="NORMAL (rede completa)"
)

df_cortado = calcular_tempos(
    df_freguesias, OSRM_CACHE_FILE_CORTADO, extra_params="&exclude=motorway",
    label="ESTRADAS CORTADAS (sem autoestradas -> aproxima corte da A1/A23)"
)
df_cortado = corrigir_tempos_invalidos(df_normal, df_cortado)
df_cortado = aplicar_penalizacoes_manuais(df_cortado)

# Juntar os dois cenários numa só tabela de comparação
comp = df_normal[[distrito_col, concelho_col, freguesia_col, 'ULS', hospital_col, pop_col,
                   'distancia_km', 'tempo_min']].rename(
    columns={'distancia_km': 'Distancia_km_Normal', 'tempo_min': 'Tempo_min_Normal'}
)
comp['Distancia_km_Cortado'] = df_cortado['distancia_km'].values
comp['Tempo_min_Cortado'] = df_cortado['tempo_min'].values
comp['Penalizacao_Corte_Especifico_min'] = df_cortado['penalizacao_corte_especifico_min'].values
comp['Penalizacao_Atrito_Geral_min'] = df_cortado['penalizacao_atrito_geral_min'].values
comp['Penalizacao_Manual_min'] = df_cortado['penalizacao_manual_min'].values
comp['Ajustado_Sem_Rota'] = df_cortado['ajustado_sem_rota'].values

comp['Diferenca_min'] = (comp['Tempo_min_Cortado'] - comp['Tempo_min_Normal']).round(1)
comp['Diferenca_%'] = np.where(
    comp['Tempo_min_Normal'] > 0,
    (comp['Diferenca_min'] / comp['Tempo_min_Normal'] * 100).round(1),
    np.nan
)
comp['Distancia_km_Normal'] = comp['Distancia_km_Normal'].round(1)
comp['Distancia_km_Cortado'] = comp['Distancia_km_Cortado'].round(1)
comp['Tempo_min_Normal'] = comp['Tempo_min_Normal'].round(1)
comp['Tempo_min_Cortado'] = comp['Tempo_min_Cortado'].round(1)

comp['Critico_Normal'] = comp['Tempo_min_Normal'] > LIMIAR_CRITICO_MIN
comp['Critico_Cortado'] = comp['Tempo_min_Cortado'] > LIMIAR_CRITICO_MIN
comp['Passou_a_Critico'] = comp['Critico_Cortado'] & ~comp['Critico_Normal']

comp = comp.sort_values('Diferenca_min', ascending=False).reset_index(drop=True)

# ═══════════════════════════════════════════════════════════════════════════
# KPIs
# ═══════════════════════════════════════════════════════════════════════════
tempo_medio_normal = round(comp['Tempo_min_Normal'].mean(), 1)
tempo_medio_cortado = round(comp['Tempo_min_Cortado'].mean(), 1)
aumento_medio_min = round(tempo_medio_cortado - tempo_medio_normal, 1)
aumento_medio_pct = round((aumento_medio_min / tempo_medio_normal * 100), 1) if tempo_medio_normal else 0

n_freg = len(comp)
n_criticas_normal = int(comp['Critico_Normal'].sum())
n_criticas_cortado = int(comp['Critico_Cortado'].sum())
n_passaram_critico = int(comp['Passou_a_Critico'].sum())
pop_passaram_critico = int(comp.loc[comp['Passou_a_Critico'], pop_col].sum())
pop_total = int(comp[pop_col].sum())

# Última linha de dados na folha "1. Comparação por Freguesia" (cabeçalho na linha 2,
# dados a começar na linha 3) -- usada mais abaixo pelas fórmulas interativas do Dashboard.
last_row_ws1 = 2 + n_freg

# Resumo por concelho
resumo_concelho = comp.groupby(concelho_col).agg(
    Freguesias=(pop_col, 'count'),
    Populacao=(pop_col, 'sum'),
    Tempo_Normal_Medio=('Tempo_min_Normal', 'mean'),
    Tempo_Cortado_Medio=('Tempo_min_Cortado', 'mean'),
    Criticas_Normal=('Critico_Normal', 'sum'),
    Criticas_Cortado=('Critico_Cortado', 'sum'),
    Passaram_a_Critico=('Passou_a_Critico', 'sum'),
).reset_index()
resumo_concelho['Aumento_Medio_min'] = (resumo_concelho['Tempo_Cortado_Medio'] - resumo_concelho['Tempo_Normal_Medio']).round(1)
resumo_concelho['Tempo_Normal_Medio'] = resumo_concelho['Tempo_Normal_Medio'].round(1)
resumo_concelho['Tempo_Cortado_Medio'] = resumo_concelho['Tempo_Cortado_Medio'].round(1)
resumo_concelho = resumo_concelho.sort_values('Aumento_Medio_min', ascending=False)

# ═══════════════════════════════════════════════════════════════════════════
# EXCEL
# ═══════════════════════════════════════════════════════════════════════════
print("   Criando Excel...")

wb = Workbook()
wb.remove(wb.active)
ws_dash = wb.create_sheet('📊 Dashboard', 0)

header_font = Font(bold=True, color="FFFFFF")
header_fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
critico_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
piorou_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")


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


# --- Folha 0: Contexto e Metodologia ---------------------------------------
ws_ctx = wb.create_sheet('0. Contexto e Metodologia', 1)
contexto_txt = [
    "CENÁRIO: Depressão 'Kristin' (28 de janeiro de 2026) — a Proteção Civil e a CIM de Leiria",
    "identificaram o distrito de Leiria como o mais afetado do país, com estradas cortadas,",
    "falhas de energia e vários feridos/mortos. Fonte: Diário de Notícias,",
    "https://www.dn.pt/sociedade/mau-tempo-um-morto-em-vila-franca-de-xira-estradas-cortadas-e-escolas-fechadas",
    "",
    "Cortes de estrada reportados nesse dia na região de Leiria:",
    " - A1  cortada várias horas entre Pombal<->Leiria e Torres Novas<->Leiria",
    " - A23 cortada no nó de Torres Novas",
    " - IC2 com fortes constrangimentos na zona da Roca (Leiria)",
    " - EN109 (Leiria-Figueira da Foz) com dificuldades de circulação",
    " - EN236 (Lousã-Castanheira de Pêra) cortada ao trânsito",
    "",
    "METODOLOGIA:",
    " - Cenário NORMAL: tempo de acesso por estrada (OSRM/OpenStreetMap), rede viária completa.",
    " - Cenário CORTADO: mesmo cálculo, mas com 'exclude=motorway' no OSRM — força o",
    "   trajeto a evitar autoestradas/IP, aproximando o efeito do corte da A1/A23.",
    " - LIMITAÇÃO: o OSRM público só exclui CLASSES de via (autoestrada), não uma estrada",
    "   nomeada específica. Por isso os cortes na IC2, EN109 e EN236 foram ADICIONADOS",
    "   MANUALMENTE (ver tabela abaixo), em vez de calculados automaticamente pelo OSRM.",
    f" - Crítico = tempo de acesso > {LIMIAR_CRITICO_MIN} minutos.",
    "",
    "PENALIZAÇÕES MANUAIS APLICADAS (por concelho, somadas ao tempo 'Cortado' do OSRM):",
    "  Castanheira de Pêra   +30 min  ->  EN236 Lousã-Castanheira de Pêra 'cortada ao trânsito' (corte total)",
    "  Leiria                +20 min  ->  IC2 zona da Roca, 'fortes constrangimentos' (via condicionada)",
    "  Marinha Grande        +20 min  ->  EN109 Leiria-Figueira da Foz, 'dificuldades de circulação'",
    "",
    "*** AVISO: a notícia NÃO indica quantos minutos de atraso cada corte causou -- só descreve",
    "se a via estava cortada ou condicionada. Estes valores foram calibrados apenas pela",
    "gravidade da linguagem da notícia (cortada = +20min, condicionada = +10min) e são uma",
    "ESTIMATIVA EDITORIAL, não uma medição real. Substitui por dados reais (Waze/Google Maps",
    "nesse dia, ou confirmação da GNR/Câmara) assim que os tiveres -- edita o dicionário",
    "PENALIZACOES_MANUAIS no topo do script.",
    "",
    "ATRITO GERAL (árvores caídas/detritos SEM corte formal, via ainda transitável mas mais lenta):",
    f"  Todo o distrito de Leiria     +{ATRITO_GERAL_DISTRITO_MIN} min  ->  aviso vermelho/laranja em todo o distrito",
    "  + concelhos 'muito devastados' (CIM de Leiria), acima do valor de todo o distrito:",
] + [f"      {c:<20s} +{m} min" for c, m in CONCELHOS_MUITO_AFETADOS.items()] + [
    "  (também estimativa editorial -- ajusta em CONCELHOS_MUITO_AFETADOS / ATRITO_GERAL_DISTRITO_MIN)",
    "",
]
for i, linha in enumerate(contexto_txt, 1):
    c = ws_ctx.cell(row=i, column=1, value=linha)
    if i == 1:
        c.font = Font(bold=True, size=12)
    else:
        c.font = Font(size=10)
ws_ctx.column_dimensions['A'].width = 110

# --- Folha 1: Comparação por Freguesia --------------------------------------
ws1 = wb.create_sheet('1. Comparação por Freguesia', 2)
ws1['A1'] = ("Tempo_min_Normal/Cortado = tempo real por estrada (OSRM) ao hospital do ULS  |  "
             "Penalizacao_Corte_Especifico_min = EN236/IC2/EN109 (estimativa)  |  "
             "Penalizacao_Atrito_Geral_min = árvores/detritos dispersos, distrito todo + concelhos mais "
             "devastados (estimativa)  |  ambas JÁ INCLUÍDAS no Tempo_min_Cortado — ver folha "
             "'0. Contexto e Metodologia'  |  Ajustado_Sem_Rota = OSRM não achou rota sem autoestrada, "
             "tempo forçado a não ficar abaixo do Normal  |  "
             "Diferenca_min = Cortado - Normal  |  Crítico = tempo > 60min  |  "
             "Passou_a_Critico = ficou crítica só por causa do corte de estradas")
ws1['A1'].font = Font(italic=True, size=9, color="666666")

headers1 = [distrito_col, concelho_col, freguesia_col, 'ULS', hospital_col, pop_col,
            'Distancia_km_Normal', 'Tempo_min_Normal', 'Distancia_km_Cortado', 'Tempo_min_Cortado',
            'Penalizacao_Corte_Especifico_min', 'Penalizacao_Atrito_Geral_min', 'Penalizacao_Manual_min',
            'Ajustado_Sem_Rota', 'Diferenca_min', 'Diferenca_%',
            'Critico_Normal', 'Critico_Cortado', 'Passou_a_Critico']
for j, h in enumerate(headers1, 1):
    ws1.cell(row=2, column=j, value=h)

for i, (_, data) in enumerate(comp.iterrows(), 3):
    for j, h in enumerate(headers1, 1):
        val = data[h]
        if isinstance(val, (np.bool_, bool)):
            val = 'Sim' if val else 'Não'
        ws1.cell(row=i, column=j, value=val)
    if data['Passou_a_Critico']:
        for c in range(1, len(headers1) + 1):
            ws1.cell(row=i, column=c).fill = piorou_fill
    elif data['Critico_Cortado']:
        for c in range(1, len(headers1) + 1):
            ws1.cell(row=i, column=c).fill = critico_fill

style_header(ws1, len(headers1), row=2)
ws1.freeze_panes = 'A3'
autofit(ws1, len(headers1))

# --- Tabela Excel nativa: dá setas de filtro/ordenação em cada coluna -------
tab_range = f"A2:{get_column_letter(len(headers1))}{last_row_ws1}"
tabela1 = Table(displayName="TabelaComparacao", ref=tab_range)
tabela1.tableStyleInfo = TableStyleInfo(
    name="TableStyleMedium2", showRowStripes=True, showFirstColumn=False,
    showLastColumn=False, showColumnStripes=False
)
ws1.add_table(tabela1)

# --- Folha 2: Resumo por Concelho ------------------------------------------
ws2 = wb.create_sheet('2. Resumo por Concelho', 3)
headers2 = list(resumo_concelho.columns)
for j, h in enumerate(headers2, 1):
    ws2.cell(row=1, column=j, value=h)
for i, (_, data) in enumerate(resumo_concelho.iterrows(), 2):
    for j, h in enumerate(headers2, 1):
        val = data[h]
        if isinstance(val, (np.bool_, bool, np.integer)):
            val = int(val)
        ws2.cell(row=i, column=j, value=val)
style_header(ws2, len(headers2), row=1)
ws2.freeze_panes = 'A2'
autofit(ws2, len(headers2))

# --- Folha auxiliar oculta para os gráficos do Dashboard -------------------
ws_aux = wb.create_sheet('_ChartsData')
ws_aux.sheet_state = 'hidden'

# Top 10 freguesias mais penalizadas (maior Diferenca_min)
top10 = comp.nlargest(10, 'Diferenca_min')
ws_aux['A1'] = 'Freguesia'
ws_aux['B1'] = 'Diferenca_min'
for i, (_, r) in enumerate(top10.iterrows(), 2):
    ws_aux.cell(row=i, column=1, value=r[freguesia_col])
    ws_aux.cell(row=i, column=2, value=r['Diferenca_min'])

# Tempo médio Normal vs Cortado por concelho
ws_aux['D1'] = 'Concelho'
ws_aux['E1'] = 'Tempo_Normal_Medio'
ws_aux['F1'] = 'Tempo_Cortado_Medio'
for i, (_, r) in enumerate(resumo_concelho.iterrows(), 2):
    ws_aux.cell(row=i, column=4, value=r[concelho_col])
    ws_aux.cell(row=i, column=5, value=r['Tempo_Normal_Medio'])
    ws_aux.cell(row=i, column=6, value=r['Tempo_Cortado_Medio'])

# Lista para o dropdown "Filtrar por Concelho" do Dashboard: "Todos" + cada concelho
concelhos_unicos = sorted(resumo_concelho[concelho_col].astype(str).unique().tolist())
ws_aux['H1'] = 'Todos'
for i, c in enumerate(concelhos_unicos, 2):
    ws_aux.cell(row=i, column=8, value=c)
lista_filtro_range = f"'_ChartsData'!$H$1:$H${1 + len(concelhos_unicos)}"

# ═══════════════════════════════════════════════════════════════════════════
# DASHBOARD
# ═══════════════════════════════════════════════════════════════════════════
ws_dash.sheet_view.showGridLines = False
ws_dash.column_dimensions['A'].width = 2
for col in 'BCDEFGH':
    ws_dash.column_dimensions[col].width = 16

ws_dash['B2'] = 'IMPACTO DO CORTE DE ESTRADAS — REGIÃO DE LEIRIA (Depressão Kristin, 28 jan 2026)'
ws_dash['B2'].font = Font(bold=True, size=14, color="C00000")

# --- Dropdown interativo: filtra os 4 KPIs abaixo por Concelho -------------
ws_dash['B3'] = 'Filtrar por Concelho:'
ws_dash['B3'].font = Font(bold=True, size=10, color="333333")
ws_dash['C3'] = 'Todos'
ws_dash['C3'].font = Font(bold=True, size=11, color="C00000")
ws_dash['C3'].fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
dv_concelho = DataValidation(type="list", formula1=lista_filtro_range, allow_blank=False,
                              showDropDown=False)
dv_concelho.error = 'Escolhe um concelho da lista (ou "Todos").'
dv_concelho.errorTitle = 'Valor inválido'
dv_concelho.prompt = 'Escolhe um concelho para filtrar os KPIs abaixo, ou "Todos".'
dv_concelho.promptTitle = 'Filtro'
ws_dash.add_data_validation(dv_concelho)
dv_concelho.add(ws_dash['C3'])

# Referências reutilizadas pelas fórmulas dos KPIs (folha 1, linhas 3..last_row_ws1)
_s1 = "'1. Comparação por Freguesia'"
_rng_b = f"{_s1}!$B$3:$B${last_row_ws1}"   # Concelho
_rng_f = f"{_s1}!$F$3:$F${last_row_ws1}"   # População
_rng_h = f"{_s1}!$H$3:$H${last_row_ws1}"   # Tempo_min_Normal
_rng_j = f"{_s1}!$J$3:$J${last_row_ws1}"   # Tempo_min_Cortado
_rng_q = f"{_s1}!$Q$3:$Q${last_row_ws1}"   # Critico_Normal
_rng_r = f"{_s1}!$R$3:$R${last_row_ws1}"   # Critico_Cortado
_rng_s = f"{_s1}!$S$3:$S${last_row_ws1}"   # Passou_a_Critico
_crit = 'IF($C$3="Todos","*",$C$3)'        # "*" = sem filtro (todos os concelhos)

# Os valores dos cartões são FÓRMULAS (não números fixos), por isso recalculam
# sozinhos sempre que se muda o Concelho no dropdown C3.
kpi_formulas = [
    ('Tempo médio (normal)', f'=ROUND(AVERAGEIFS({_rng_h},{_rng_b},{_crit}),1)', '0.0" min"', "2E75B6"),
    ('Tempo médio (c/ cortes)', f'=ROUND(AVERAGEIFS({_rng_j},{_rng_b},{_crit}),1)', '0.0" min"', "C00000"),
    ('Aumento médio', None, None, "C00000"),                      # preenchido abaixo (depende dos 2 cartões anteriores)
    ('Freguesias que passaram a críticas', None, None, "C00000"),  # preenchido abaixo
]
card_cols = ['B', 'D', 'F', 'H']
for (titulo, formula, numfmt, cor), col in zip(kpi_formulas, card_cols):
    col2 = chr(ord(col) + 1)
    ws_dash.merge_cells(f'{col}4:{col2}4')
    ws_dash.merge_cells(f'{col}5:{col2}6')
    c_titulo = ws_dash[f'{col}4']
    c_titulo.value = titulo
    c_titulo.font = Font(bold=True, size=10, color="666666")
    c_titulo.alignment = Alignment(horizontal='center')
    c_valor = ws_dash[f'{col}5']
    if formula:
        c_valor.value = formula
        c_valor.number_format = numfmt
    c_valor.font = Font(bold=True, size=18, color=cor)
    c_valor.alignment = Alignment(horizontal='center', vertical='center')
    for row in (4, 5, 6):
        for c in (col, col2):
            ws_dash[f'{c}{row}'].fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

# 'Aumento médio' depende dos valores calculados nos 2 primeiros cartões (B5 e D5)
ws_dash['F5'] = '="+"&TEXT(D5-B5,"0.0")&" min ("&TEXT((D5-B5)/B5,"0.0%")&")"'

# 'Freguesias que passaram a críticas' -- contagem já filtrada pelo Concelho escolhido
ws_dash['H5'] = f'=COUNTIFS({_rng_s},"Sim",{_rng_b},{_crit})&"/"&COUNTIFS({_rng_b},{_crit})'

# Linha de contexto -- também fórmula, também respeita o filtro de Concelho
ws_dash['B8'] = (
    f'="Críticas (>60min) antes: "&COUNTIFS({_rng_b},{_crit},{_rng_q},"Sim")'
    f'&"  |  depois do corte: "&COUNTIFS({_rng_b},{_crit},{_rng_r},"Sim")'
    f'&"  |  população nas freguesias que pioraram: "&TEXT(SUMIFS({_rng_f},{_rng_b},{_crit},{_rng_s},"Sim"),"#,##0")'
    f'&" de "&TEXT(SUMIFS({_rng_f},{_rng_b},{_crit}),"#,##0")&" hab."'
)
ws_dash['B8'].font = Font(italic=True, size=10, color="666666")

# Gráfico 1: Top 10 freguesias mais penalizadas
chart1 = BarChart()
chart1.type = "bar"
chart1.title = "Top 10 Freguesias Mais Penalizadas (min a mais no acesso)"
chart1.x_axis.title = "Minutos a mais"
chart1.style = 12
data1 = Reference(ws_aux, min_col=2, min_row=1, max_row=1 + len(top10))
cats1 = Reference(ws_aux, min_col=1, min_row=2, max_row=1 + len(top10))
chart1.add_data(data1, titles_from_data=True)
chart1.set_categories(cats1)
chart1.height, chart1.width = 10, 16
chart1.legend = None
ws_dash.add_chart(chart1, 'B10')

# Gráfico 2: Tempo médio Normal vs Cortado por concelho
chart2 = BarChart()
chart2.type = "col"
chart2.title = "Tempo Médio de Acesso por Concelho: Normal vs Estradas Cortadas"
chart2.y_axis.title = "Minutos"
chart2.style = 10
n_conc = len(resumo_concelho)
data2 = Reference(ws_aux, min_col=5, max_col=6, min_row=1, max_row=1 + n_conc)
cats2 = Reference(ws_aux, min_col=4, min_row=2, max_row=1 + n_conc)
chart2.add_data(data2, titles_from_data=True)
chart2.set_categories(cats2)
chart2.height, chart2.width = 10, 20
chart2.legend.position = 'b'
ws_dash.add_chart(chart2, 'B27')

ws_dash['B48'] = ('Nota: os 4 cartões acima e a linha de contexto seguem o dropdown "Filtrar por Concelho" (C3). '
                   'Os 2 gráficos mostram sempre o distrito completo. Na folha "1. Comparação por Freguesia", '
                   'usa as setas de filtro no cabeçalho da tabela para ver só as freguesias que te interessam.')
ws_dash['B48'].font = Font(italic=True, size=9, color="999999")

output = os.path.join(SCRIPT_DIR, 'Impacto_Estradas_Cortadas_Leiria.xlsx')
wb.save(output)

# ═══════════════════════════════════════════════════════════════════════════
# FIM
# ═══════════════════════════════════════════════════════════════════════════
print("\n" + "=" * 70)
print("✅ CONCLUÍDO!")
print("=" * 70)
print(f"\n📊 Resumo ({DISTRITO_ALVO}, {n_freg} freguesias):")
print(f"   Tempo médio normal:  {tempo_medio_normal} min")
print(f"   Tempo médio cortado: {tempo_medio_cortado} min  (+{aumento_medio_min} min, {aumento_medio_pct}%)")
print(f"   Freguesias críticas (>60min): {n_criticas_normal} -> {n_criticas_cortado}")
print(f"   Freguesias que SÓ passaram a críticas por causa do corte: {n_passaram_critico}")
print(f"   População nessas freguesias: {pop_passaram_critico:,} hab.")
print(f"\n📁 Ficheiro: {output}")
print("=" * 70)