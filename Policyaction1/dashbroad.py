#!/usr/bin/env python3
"""
GERAR DASHBOARD INTERATIVO — MODELO 1 (2SFCA / Reafetação Hospitalar)
════════════════════════════════════════════════════════════════════
Lê o output já gerado pelo modelo.py (Modelo1_Reafetacao_2SFCA.xlsx) e
acrescenta uma folha "0. Dashboard" no INÍCIO do workbook, com:

  - Cartões KPI (ligados por fórmula à folha "1. Resumo")
  - Filtro interativo por Distrito (dropdown / validação de dados) que
    atualiza um mini-resumo via SUMPRODUCT / COUNTIFS
  - Tabela "Análise por Distrito" (18 linhas) com fórmulas SUMPRODUCT /
    COUNTIF / SUMIF sobre a folha "2. Reafetacao_Freguesias"
  - Tabela "Top 10 ULS por Ocupação (%)" com fórmulas LARGE + INDEX/MATCH
    sobre a folha "4. Capacidade_ULS"
  - Tabela "Top 15 freguesias com maior poupança de tempo" (a folha 2 já
    vem ordenada por poupança decrescente, por isso são apenas links
    diretos às primeiras 15 linhas)
  - 4 gráficos nativos do Excel (barras/pizza), todos ligados às tabelas
    de fórmulas acima — por isso são interativos: mudam sempre que os
    dados de origem mudam.

Não altera nenhuma das folhas originais do modelo (1, 2, 3, 4).
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
import os

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_FILE = os.path.join(SCRIPT_DIR, "Modelo1_Reafetacao_2SFCA.xlsx")
OUTPUT_FILE = os.path.join(SCRIPT_DIR, "Modelo1_Reafetacao_2SFCA_Dashboard.xlsx")

# ── Estilo ────────────────────────────────────────────────────────────────
AZUL       = "2E75B6"
AZUL_CLARO = "DCE6F1"
VERDE      = "2E7D32"
VERMELHO   = "C0392B"
CINZA      = "F2F2F2"
FONT_NAME  = "Arial"

HEADER_FILL = PatternFill(start_color=AZUL, end_color=AZUL, fill_type="solid")
HEADER_FONT = Font(name=FONT_NAME, bold=True, color="FFFFFF")
KPI_FILL    = PatternFill(start_color=AZUL_CLARO, end_color=AZUL_CLARO, fill_type="solid")
TITLE_FONT  = Font(name=FONT_NAME, bold=True, size=18, color=AZUL)
SUB_FONT    = Font(name=FONT_NAME, italic=True, size=10, color="666666")
LABEL_FONT  = Font(name=FONT_NAME, bold=True, size=9, color="666666")
KPI_FONT    = Font(name=FONT_NAME, bold=True, size=20, color=AZUL)
SECTION_FONT= Font(name=FONT_NAME, bold=True, size=12, color="FFFFFF")
SECTION_FILL= PatternFill(start_color=AZUL, end_color=AZUL, fill_type="solid")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

DISTRITOS = ['Aveiro', 'Beja', 'Braga', 'Bragança', 'Castelo Branco', 'Coimbra',
             'Faro', 'Guarda', 'Leiria', 'Lisboa', 'Portalegre', 'Porto',
             'Santarém', 'Setúbal', 'Viana do Castelo', 'Vila Real', 'Viseu', 'Évora']

N_FREG_LAST_ROW = 2883   # folha 2: dados em linhas 2..2883 (2882 freguesias)
N_ULS_LAST_ROW  = 40     # folha 4: dados em linhas 2..40 (39 ULS)

S2 = "'2. Reafetacao_Freguesias'"
S4 = "'4. Capacidade_ULS'"


def style_section_title(ws, cell_range, text):
    ws.merge_cells(cell_range)
    top_left = cell_range.split(":")[0]
    ws[top_left] = text
    ws[top_left].font = SECTION_FONT
    ws[top_left].fill = SECTION_FILL
    ws[top_left].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    row = int(''.join(filter(str.isdigit, top_left)))
    ws.row_dimensions[row].height = 22


def kpi_card(ws, cell_range, label, formula, font=KPI_FONT, num_fmt=None):
    ws.merge_cells(cell_range)
    top_left = cell_range.split(":")[0]
    col = ''.join(filter(str.isalpha, top_left))
    row = int(''.join(filter(str.isdigit, top_left)))
    # label acima do valor
    lbl_row = row - 1
    ws.cell(row=lbl_row, column=openpyxl.utils.column_index_from_string(col),
            value=label).font = LABEL_FONT
    c = ws[top_left]
    c.value = formula
    c.font = font
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = KPI_FILL
    if num_fmt:
        c.number_format = num_fmt
    for r in ws[cell_range]:
        for cell in r:
            cell.fill = KPI_FILL
            cell.border = BORDER


def main():
    print("A carregar workbook do modelo...")
    wb = openpyxl.load_workbook(INPUT_FILE)

    if "0. Dashboard" in wb.sheetnames:
        del wb["0. Dashboard"]
    ws = wb.create_sheet("0. Dashboard", 0)
    ws.sheet_view.showGridLines = False

    # larguras de coluna
    widths = {"A": 3, "B": 17, "C": 17, "D": 17, "E": 17, "F": 17, "G": 17,
              "H": 17, "I": 17, "J": 17, "K": 4}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    # ── Título ──────────────────────────────────────────────────────────
    ws.merge_cells("B2:J2")
    ws["B2"] = "DASHBOARD INTERATIVO — REAFETAÇÃO HOSPITALAR (MODELO 2SFCA)"
    ws["B2"].font = TITLE_FONT
    ws.merge_cells("B3:J3")
    ws["B3"] = ("Comparação hospital de referência ATUAL vs. atribuição ÓTIMA (minimização do tempo de "
                "trajeto, sujeita à capacidade real de camas) — 2882 freguesias, 39 ULS/hospitais")
    ws["B3"].font = SUB_FONT
    ws.row_dimensions[2].height = 26

    # ── KPIs nacionais (ligados à folha 1. Resumo) ─────────────────────
    row_lbl, row_val = 5, 6
    kpi_card(ws, f"B{row_val}:B{row_val}", "Nº FREGUESIAS", "='1. Resumo'!C5", num_fmt="#,##0")
    kpi_card(ws, f"C{row_val}:C{row_val}", "Nº HOSPITAIS (ULS)", "='1. Resumo'!C6", num_fmt="#,##0")
    kpi_card(ws, f"D{row_val}:D{row_val}", "TEMPO MÉDIO ATUAL", "='1. Resumo'!C8", num_fmt='0.0" min"')
    kpi_card(ws, f"E{row_val}:E{row_val}", "TEMPO MÉDIO ÓTIMO", "='1. Resumo'!C9", num_fmt='0.0" min"')
    kpi_card(ws, f"F{row_val}:F{row_val}", "REDUÇÃO DE TEMPO", "='1. Resumo'!C10", num_fmt='0.0" min"')
    kpi_card(ws, f"G{row_val}:G{row_val}", "% FREG. MUDAM DE ULS",
             f"=COUNTIF({S2}!J2:J{N_FREG_LAST_ROW},TRUE)/COUNTA({S2}!J2:J{N_FREG_LAST_ROW})",
             num_fmt="0.0%")
    kpi_card(ws, f"H{row_val}:H{row_val}", "% CRÍTICAS (>60min) ATUAL", "='1. Resumo'!C13/100", num_fmt="0.0%")
    kpi_card(ws, f"I{row_val}:I{row_val}", "% CRÍTICAS (>60min) ÓTIMO", "='1. Resumo'!C14/100", num_fmt="0.0%")
    kpi_card(ws, f"J{row_val}:J{row_val}", "DÉFICE ESTRUTURAL", "='1. Resumo'!C17", num_fmt='0.0" camas"')
    for c in "BCDEFGHIJ":
        ws.row_dimensions[row_val].height = 30

    # ── Filtro interativo por Distrito ─────────────────────────────────
    style_section_title(ws, "B9:J9", "🔎  FILTRO INTERATIVO — RESUMO POR DISTRITO SELECIONADO")

    ws["B11"] = "Escolher Distrito:"
    ws["B11"].font = Font(name=FONT_NAME, bold=True, size=11)
    ws["C11"] = "Lisboa"
    ws["C11"].font = Font(name=FONT_NAME, bold=True, size=11, color=VERMELHO)
    ws["C11"].fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    ws["C11"].alignment = Alignment(horizontal="center")
    ws["C11"].border = BORDER

    # lista auxiliar de distritos (fora da área visível principal, coluna N)
    ws["N1"] = "lista_distritos"
    for i, d in enumerate(DISTRITOS, start=2):
        ws.cell(row=i, column=14, value=d)  # coluna N
    ws.column_dimensions["N"].width = 2
    ws.column_dimensions["N"].hidden = False  # mantido visível mas estreito

    dv = DataValidation(type="list", formula1="=$N$2:$N$19", allow_blank=False,
                         showDropDown=False)
    dv.error = "Escolha um distrito da lista."
    dv.errorTitle = "Distrito inválido"
    ws.add_data_validation(dv)
    dv.add(ws["C11"])

    filtro_labels = ["Nº Freguesias", "População total", "Tempo médio ATUAL (pond.)",
                      "Tempo médio ÓTIMO (pond.)", "Poupança média", "% que muda de ULS"]
    filtro_formulas = [
        f'=COUNTIF({S2}!A2:A{N_FREG_LAST_ROW},$C$11)',
        f'=SUMIF({S2}!A2:A{N_FREG_LAST_ROW},$C$11,{S2}!D2:D{N_FREG_LAST_ROW})',
        f'=SUMPRODUCT(({S2}!$A2:A{N_FREG_LAST_ROW}=$C$11)*{S2}!D2:D{N_FREG_LAST_ROW}*{S2}!F2:F{N_FREG_LAST_ROW})'
        f'/SUMPRODUCT(({S2}!$A2:A{N_FREG_LAST_ROW}=$C$11)*{S2}!D2:D{N_FREG_LAST_ROW})',
        f'=SUMPRODUCT(({S2}!$A2:A{N_FREG_LAST_ROW}=$C$11)*{S2}!D2:D{N_FREG_LAST_ROW}*{S2}!H2:H{N_FREG_LAST_ROW})'
        f'/SUMPRODUCT(({S2}!$A2:A{N_FREG_LAST_ROW}=$C$11)*{S2}!D2:D{N_FREG_LAST_ROW})',
        f'=SUMIF({S2}!A2:A{N_FREG_LAST_ROW},$C$11,{S2}!I2:I{N_FREG_LAST_ROW})/COUNTIF({S2}!A2:A{N_FREG_LAST_ROW},$C$11)',
        f'=COUNTIFS({S2}!A2:A{N_FREG_LAST_ROW},$C$11,{S2}!J2:J{N_FREG_LAST_ROW},TRUE)'
        f'/COUNTIF({S2}!A2:A{N_FREG_LAST_ROW},$C$11)',
    ]
    fmts = ["#,##0", "#,##0", '0.0" min"', '0.0" min"', '0.0" min"', "0.0%"]
    col0 = 4  # coluna D
    for i, (lbl, formula, fmt) in enumerate(zip(filtro_labels, filtro_formulas, fmts)):
        col = col0 + i
        letter = get_column_letter(col)
        ws.cell(row=10, column=col, value=lbl).font = LABEL_FONT
        cell = ws.cell(row=11, column=col, value=formula)
        cell.font = Font(name=FONT_NAME, bold=True, size=12, color=VERDE)
        cell.number_format = fmt
        cell.fill = KPI_FILL
        cell.alignment = Alignment(horizontal="center")
        cell.border = BORDER
    ws.row_dimensions[11].height = 24

    # ── Tabela: Análise por Distrito (fórmulas) ─────────────────────────
    style_section_title(ws, "B14:F14", "📊  ANÁLISE POR DISTRITO")
    hdr_row = 15
    headers = ["Distrito", "Nº Freguesias", "Tempo Médio ATUAL (min)",
               "Tempo Médio ÓTIMO (min)", "% Muda de ULS"]
    for j, h in enumerate(headers, 2):
        c = ws.cell(row=hdr_row, column=j, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[hdr_row].height = 30

    dist_start_row = hdr_row + 1
    for i, d in enumerate(DISTRITOS):
        r = dist_start_row + i
        ws.cell(row=r, column=2, value=d).border = BORDER
        ws.cell(row=r, column=3,
                value=f'=COUNTIF({S2}!$A$2:$A${N_FREG_LAST_ROW},$B{r})').border = BORDER
        ws.cell(row=r, column=4,
                value=(f'=ROUND(SUMPRODUCT(({S2}!$A$2:$A${N_FREG_LAST_ROW}=$B{r})*{S2}!$D$2:$D${N_FREG_LAST_ROW}*{S2}!$F$2:$F${N_FREG_LAST_ROW})'
                       f'/SUMPRODUCT(({S2}!$A$2:$A${N_FREG_LAST_ROW}=$B{r})*{S2}!$D$2:$D${N_FREG_LAST_ROW}),1)')
                ).border = BORDER
        ws.cell(row=r, column=5,
                value=(f'=ROUND(SUMPRODUCT(({S2}!$A$2:$A${N_FREG_LAST_ROW}=$B{r})*{S2}!$D$2:$D${N_FREG_LAST_ROW}*{S2}!$H$2:$H${N_FREG_LAST_ROW})'
                       f'/SUMPRODUCT(({S2}!$A$2:$A${N_FREG_LAST_ROW}=$B{r})*{S2}!$D$2:$D${N_FREG_LAST_ROW}),1)')
                ).border = BORDER
        ws.cell(row=r, column=6,
                value=(f'=COUNTIFS({S2}!$A$2:$A${N_FREG_LAST_ROW},$B{r},{S2}!$J$2:$J${N_FREG_LAST_ROW},TRUE)'
                       f'/COUNTIF({S2}!$A$2:$A${N_FREG_LAST_ROW},$B{r})')
                ).number_format = "0.0%"
        ws.cell(row=r, column=6).border = BORDER
        for col in (2, 3, 4, 5):
            ws.cell(row=r, column=col).alignment = Alignment(horizontal="center")
            if (i % 2) == 1:
                ws.cell(row=r, column=col).fill = PatternFill(start_color=CINZA, end_color=CINZA, fill_type="solid")
        ws.cell(row=r, column=6).alignment = Alignment(horizontal="center")
        if (i % 2) == 1:
            ws.cell(row=r, column=6).fill = PatternFill(start_color=CINZA, end_color=CINZA, fill_type="solid")
    dist_end_row = dist_start_row + len(DISTRITOS) - 1

    # ── Tabela: Top 10 ULS por Ocupação (%) ─────────────────────────────
    style_section_title(ws, "H14:J14", "🏥  TOP 10 ULS POR OCUPAÇÃO (%)")
    ws.cell(row=hdr_row, column=8, value="ULS").font = HEADER_FONT
    ws.cell(row=hdr_row, column=8).fill = HEADER_FILL
    ws.cell(row=hdr_row, column=8).border = BORDER
    ws.cell(row=hdr_row, column=9, value="Ocupação (%)").font = HEADER_FONT
    ws.cell(row=hdr_row, column=9).fill = HEADER_FILL
    ws.cell(row=hdr_row, column=9).alignment = Alignment(horizontal="center", wrap_text=True)
    ws.cell(row=hdr_row, column=9).border = BORDER
    ws.merge_cells(start_row=hdr_row, start_column=9, end_row=hdr_row, end_column=10)

    # helper column (única, sem empates) na folha 4 para permitir LARGE + MATCH exatos
    ws4 = wb["4. Capacidade_ULS"]
    ws4["I1"] = "Rank_Helper"
    ws4["I1"].font = HEADER_FONT
    ws4["I1"].fill = HEADER_FILL
    for r in range(2, N_ULS_LAST_ROW + 1):
        ws4.cell(row=r, column=9, value=f"=D{r}+ROW()/100000000")

    top_uls_start = hdr_row + 1
    for i in range(10):
        r = top_uls_start + i
        rank_formula = f'=LARGE({S4}!$I$2:$I${N_ULS_LAST_ROW},{i + 1})'
        helper_cell = f"$L{r}"
        ws.cell(row=r, column=12, value=rank_formula)  # coluna L auxiliar (oculta)
        ws.cell(row=r, column=8,
                value=(f'=INDEX({S4}!$A$2:$A${N_ULS_LAST_ROW},MATCH({helper_cell},{S4}!$I$2:$I${N_ULS_LAST_ROW},0))')
                ).border = BORDER
        ocup_cell = ws.cell(row=r, column=9,
                value=(f'=INDEX({S4}!$D$2:$D${N_ULS_LAST_ROW},MATCH({helper_cell},{S4}!$I$2:$I${N_ULS_LAST_ROW},0))'))
        ws.merge_cells(start_row=r, start_column=9, end_row=r, end_column=10)
        ocup_cell.number_format = '0.0"%"'
        ocup_cell.alignment = Alignment(horizontal="center")
        ocup_cell.border = BORDER
        ws.cell(row=r, column=8).alignment = Alignment(horizontal="left")
        if i % 2 == 1:
            for col in (8, 9, 10):
                ws.cell(row=r, column=col).fill = PatternFill(start_color=CINZA, end_color=CINZA, fill_type="solid")
    ws.column_dimensions["L"].width = 2
    top_uls_end = top_uls_start + 9

    # ── Tabela: Top 15 freguesias com maior poupança de tempo ──────────
    top15_title_row = dist_end_row + 3
    style_section_title(ws, f"B{top15_title_row}:F{top15_title_row}",
                         "⏱️  TOP 15 FREGUESIAS COM MAIOR POUPANÇA DE TEMPO (folha 2, já ordenada)")
    top15_hdr = top15_title_row + 1
    headers15 = ["Freguesia", "Distrito", "ULS Atual → ULS Ótimo", "Tempo Atual (min)", "Poupança (min)"]
    for j, h in enumerate(headers15, 2):
        c = ws.cell(row=top15_hdr, column=j, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[top15_hdr].height = 28
    top15_start = top15_hdr + 1
    for i in range(15):
        src_row = 2 + i  # folha 2 já vem ordenada por Poupanca_Tempo_min desc
        r = top15_start + i
        ws.cell(row=r, column=2, value=f"={S2}!C{src_row}").border = BORDER
        ws.cell(row=r, column=3, value=f"={S2}!A{src_row}").border = BORDER
        ws.cell(row=r, column=4,
                value=f'={S2}!E{src_row}&" → "&{S2}!G{src_row}').border = BORDER
        ws.cell(row=r, column=5, value=f"={S2}!F{src_row}").border = BORDER
        ws.cell(row=r, column=5).number_format = "0.0"
        ws.cell(row=r, column=6, value=f"={S2}!I{src_row}").border = BORDER
        ws.cell(row=r, column=6).number_format = "0.0"
        for col in (2, 3, 4, 5, 6):
            ws.cell(row=r, column=col).alignment = Alignment(horizontal="left" if col in (2, 3, 4) else "center")
            if i % 2 == 1:
                ws.cell(row=r, column=col).fill = PatternFill(start_color=CINZA, end_color=CINZA, fill_type="solid")
    top15_end = top15_start + 14

    # ── Contagem nacional Mudou vs Não Mudou (para o gráfico de pizza) ──
    pie_row = top15_end + 3
    ws.cell(row=pie_row, column=8, value="Mudou de ULS").font = LABEL_FONT
    ws.cell(row=pie_row, column=9,
            value=f"=COUNTIF({S2}!J2:J{N_FREG_LAST_ROW},TRUE)")
    ws.cell(row=pie_row + 1, column=8, value="Manteve ULS").font = LABEL_FONT
    ws.cell(row=pie_row + 1, column=9,
            value=f"=COUNTIF({S2}!J2:J{N_FREG_LAST_ROW},FALSE)")

    # ══════════════════════════════════════════════════════════════════
    # GRÁFICOS — dispostos em pilha vertical, largura inteira, sem
    # sobreposições. Cada gráfico fica logo abaixo da sua tabela de
    # origem e reserva um bloco de linhas em branco calculado a partir
    # da sua altura real (evita que o gráfico seguinte lhe caia em cima).
    # ══════════════════════════════════════════════════════════════════
    CHART_W_WIDE = 24     # cm
    CHART_H      = 11     # cm
    ROWS_PER_CHART = 24   # linhas reservadas por gráfico de altura "normal" (11 cm)

    def legend_bottom(chart):
        chart.legend.position = "b"
        chart.legend.overlay = False

    next_row = max(dist_end_row, top_uls_end) + 3

    # 1) Barras: Tempo médio Atual vs Ótimo por Distrito (larga, 18 categorias)
    style_section_title(ws, f"B{next_row}:J{next_row}",
                         "📈  GRÁFICO — Tempo médio de acesso por Distrito (Atual vs. Ótimo)")
    chart1_row = next_row + 1
    chart1 = BarChart()
    chart1.type = "col"
    chart1.grouping = "clustered"
    chart1.title = "Tempo médio ponderado por Distrito (min)"
    chart1.y_axis.title = "Minutos"
    chart1.x_axis.title = "Distrito"
    chart1.style = 10
    data1 = Reference(ws, min_col=4, max_col=5, min_row=hdr_row, max_row=dist_end_row)
    cats1 = Reference(ws, min_col=2, min_row=dist_start_row, max_row=dist_end_row)
    chart1.add_data(data1, titles_from_data=True)
    chart1.set_categories(cats1)
    chart1.width = CHART_W_WIDE
    chart1.height = CHART_H
    legend_bottom(chart1)
    ws.add_chart(chart1, f"B{chart1_row}")
    next_row = chart1_row + ROWS_PER_CHART

    # 2) Barras: Top 10 ULS por Ocupação
    style_section_title(ws, f"B{next_row}:J{next_row}",
                         "📈  GRÁFICO — Top 10 ULS por Ocupação (%)")
    chart2_row = next_row + 1
    chart2 = BarChart()
    chart2.type = "bar"
    chart2.title = "Top 10 ULS por Ocupação (%)"
    chart2.y_axis.title = "ULS"
    chart2.x_axis.title = "Ocupação (%)"
    chart2.style = 11
    data2 = Reference(ws, min_col=9, max_col=9, min_row=hdr_row, max_row=top_uls_end)
    cats2 = Reference(ws, min_col=8, min_row=top_uls_start, max_row=top_uls_end)
    chart2.add_data(data2, titles_from_data=True)
    chart2.set_categories(cats2)
    chart2.width = CHART_W_WIDE
    chart2.height = CHART_H
    legend_bottom(chart2)
    ws.add_chart(chart2, f"B{chart2_row}")
    next_row = chart2_row + ROWS_PER_CHART

    # 4) Barras horizontais: Top 15 freguesias com maior poupança
    style_section_title(ws, f"B{next_row}:J{next_row}",
                         "📈  GRÁFICO — Top 15 freguesias com maior poupança de tempo (min)")
    chart4_row = next_row + 1
    chart4 = BarChart()
    chart4.type = "bar"
    chart4.title = "Top 15 freguesias — poupança de tempo (min)"
    chart4.x_axis.title = "Minutos poupados"
    chart4.style = 12
    data4 = Reference(ws, min_col=6, max_col=6, min_row=top15_hdr, max_row=top15_end)
    cats4 = Reference(ws, min_col=2, min_row=top15_start, max_row=top15_end)
    chart4.add_data(data4, titles_from_data=True)
    chart4.set_categories(cats4)
    chart4.width = CHART_W_WIDE
    chart4.height = CHART_H + 2
    legend_bottom(chart4)
    ws.add_chart(chart4, f"B{chart4_row}")
    next_row = chart4_row + ROWS_PER_CHART + 2

    # 3) Pizza: Mudou vs Manteve ULS
    style_section_title(ws, f"B{next_row}:J{next_row}",
                         "📈  GRÁFICO — Freguesias que mudam de ULS (nível nacional)")
    chart3_row = next_row + 1
    chart3 = PieChart()
    chart3.title = "Freguesias que mudam de ULS"
    data3 = Reference(ws, min_col=9, min_row=pie_row, max_row=pie_row + 1)
    cats3 = Reference(ws, min_col=8, min_row=pie_row, max_row=pie_row + 1)
    chart3.add_data(data3, titles_from_data=False)
    chart3.set_categories(cats3)
    chart3.dataLabels = DataLabelList()
    chart3.dataLabels.showPercent = True
    chart3.dataLabels.showVal = True
    chart3.width = 15
    chart3.height = CHART_H
    legend_bottom(chart3)
    ws.add_chart(chart3, f"B{chart3_row}")
    next_row = chart3_row + ROWS_PER_CHART

    # esconder colunas auxiliares (N = lista distritos, L = rank helper)
    ws.column_dimensions["L"].hidden = True

    # congelar painel e mover dashboard para o separador ativo
    ws.freeze_panes = "B5"
    ws.sheet_view.zoomScale = 90
    wb.active = 0

    print("A guardar workbook...")
    wb.save(OUTPUT_FILE)
    print(f"✅ Guardado em: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()