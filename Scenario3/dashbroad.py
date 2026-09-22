#!/usr/bin/env python3
"""
GERAR DASHBOARD INTERATIVO — MODELO 3 (Impacto de +10% na População Idosa)
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import BarChart, Reference
import os

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_FILE = os.path.join(SCRIPT_DIR, "Modelo3_Impacto_Pop_Idosa.xlsx")
OUTPUT_FILE = os.path.join(SCRIPT_DIR, "Modelo3_Impacto_Pop_Idosa_Dashboard.xlsx")

# ── Estilo (vermelho — cor de destaque do Modelo 3) ─────────────────────────
VERMELHO       = "C00000"
VERMELHO_ESC   = "8C0000"
VERMELHO_CLARO = "F7DADA"
LARANJA        = "E07A2C"
AZUL           = "2E75B6"
VERDE          = "2E9E6B"
CINZA          = "F2F2F2"
FONT_NAME      = "Arial"

HEADER_FILL = PatternFill(start_color=VERMELHO, end_color=VERMELHO, fill_type="solid")
HEADER_FONT = Font(name=FONT_NAME, bold=True, color="FFFFFF")
KPI_FILL    = PatternFill(start_color=VERMELHO_CLARO, end_color=VERMELHO_CLARO, fill_type="solid")
TITLE_FONT  = Font(name=FONT_NAME, bold=True, size=18, color=VERMELHO_ESC)
SUB_FONT    = Font(name=FONT_NAME, italic=True, size=10, color="666666")
LABEL_FONT  = Font(name=FONT_NAME, bold=True, size=9, color="666666")
KPI_FONT    = Font(name=FONT_NAME, bold=True, size=18, color=VERMELHO_ESC)
SECTION_FONT= Font(name=FONT_NAME, bold=True, size=12, color="FFFFFF")
SECTION_FILL= PatternFill(start_color=VERMELHO, end_color=VERMELHO, fill_type="solid")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

DISTRITOS = ['Aveiro', 'Beja', 'Braga', 'Bragança', 'Castelo Branco', 'Coimbra',
             'Faro', 'Guarda', 'Leiria', 'Lisboa', 'Portalegre', 'Porto',
             'Santarém', 'Setúbal', 'Viana do Castelo', 'Vila Real', 'Viseu', 'Évora']

N_FREG_LAST_ROW = 2884   # folha 3: dados em linhas 3..2884 (2882 freguesias)
N_ULS_LAST_ROW  = 40     # folha 2: dados em linhas 2..40 (39 ULS)

S3 = "'3. Por_Freguesia'"     # A=Distrito B=Concelho C=Freguesia D=ULS
                               # K(11)=Pop65_Antes L(12)=Pop65_Depois M(13)=Camas_Antes N(14)=Camas_Depois
S2 = "'2. Cobertura_ULS'"     # A=ULS B=Camas_Antes C=Camas_Depois D=Capacidade
                               # E=Cobertura_Antes F=Cobertura_Depois G=Falta_Antes H=Falta_Depois I=Delta


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

    widths = {"A": 3, "B": 19, "C": 19, "D": 19, "E": 19, "F": 19, "G": 19,
              "H": 19, "I": 19, "J": 19, "K": 4}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    # ── Título ──────────────────────────────────────────────────────────
    ws.merge_cells("B2:J2")
    ws["B2"] = "DASHBOARD INTERATIVO — IMPACTO DE +10% NA POPULAÇÃO IDOSA (MODELO 3)"
    ws["B2"].font = TITLE_FONT
    ws.merge_cells("B3:J3")
    ws["B3"] = ("O que acontece à procura e à cobertura hospitalar do país se a população 65+ crescer 10%, "
                "mantendo a oferta de camas constante? — 2882 freguesias, 39 ULS, mesma rede geográfica")
    ws["B3"].font = SUB_FONT
    ws.row_dimensions[2].height = 26

    # ── KPIs nacionais (ligados à folha 1. Resumo) ─────────────────────
    row_val = 6
    kpi_card(ws, f"B{row_val}:B{row_val}", "CAMAS NECESSÁRIAS (ANTES)", "='1. Resumo'!C5", num_fmt="#,##0")
    kpi_card(ws, f"C{row_val}:C{row_val}", "CAMAS NECESSÁRIAS (DEPOIS)", "='1. Resumo'!C6", num_fmt="#,##0")
    kpi_card(ws, f"D{row_val}:D{row_val}", "ACRÉSCIMO DE CAMAS", "='1. Resumo'!C7", font=Font(name=FONT_NAME, bold=True, size=18, color=VERMELHO),
             num_fmt="+#,##0")
    kpi_card(ws, f"E{row_val}:E{row_val}", "COBERTURA MÉDIA (ANTES)", "='1. Resumo'!C9/100", num_fmt="0.0%")
    kpi_card(ws, f"F{row_val}:F{row_val}", "COBERTURA MÉDIA (DEPOIS)", "='1. Resumo'!C10/100", num_fmt="0.0%")
    kpi_card(ws, f"G{row_val}:G{row_val}", "QUEDA DE COBERTURA",
             "=('1. Resumo'!C10-'1. Resumo'!C9)/100", num_fmt="+0.0%")
    ws.row_dimensions[row_val].height = 30

    ws.merge_cells("B7:J7")
    ws["B7"] = ('="Nota: a procura de camas cresce proporcionalmente à população idosa (+10%); a oferta "'
                '&"mantém-se constante, pelo que a cobertura oferta/necessidade por ULS piora na mesma proporção."')
    ws["B7"].font = SUB_FONT

    # ── Filtro interativo por Distrito ─────────────────────────────────
    style_section_title(ws, "B9:J9", "🔎  FILTRO INTERATIVO — RESUMO POR DISTRITO SELECIONADO")

    ws["B11"] = "Escolher Distrito:"
    ws["B11"].font = Font(name=FONT_NAME, bold=True, size=11)
    ws["C11"] = "Lisboa"
    ws["C11"].font = Font(name=FONT_NAME, bold=True, size=11, color=VERMELHO_ESC)
    ws["C11"].fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    ws["C11"].alignment = Alignment(horizontal="center")
    ws["C11"].border = BORDER

    ws["N1"] = "lista_distritos"
    for i, d in enumerate(DISTRITOS, start=2):
        ws.cell(row=i, column=14, value=d)
    ws.column_dimensions["N"].width = 2

    dv = DataValidation(type="list", formula1="=$N$2:$N$19", allow_blank=False, showDropDown=False)
    dv.error = "Escolha um distrito da lista."
    dv.errorTitle = "Distrito inválido"
    ws.add_data_validation(dv)
    dv.add(ws["C11"])

    filtro_labels = ["Nº Freguesias", "Pop. 65+ (ANTES)", "Pop. 65+ (DEPOIS)",
                      "Camas Necess. (ANTES)", "Camas Necess. (DEPOIS)", "Acréscimo de Camas"]
    filtro_formulas = [
        f'=COUNTIF({S3}!A3:A{N_FREG_LAST_ROW},$C$11)',
        f'=ROUND(SUMIF({S3}!A3:A{N_FREG_LAST_ROW},$C$11,{S3}!K3:K{N_FREG_LAST_ROW}),0)',
        f'=ROUND(SUMIF({S3}!A3:A{N_FREG_LAST_ROW},$C$11,{S3}!L3:L{N_FREG_LAST_ROW}),0)',
        f'=ROUND(SUMIF({S3}!A3:A{N_FREG_LAST_ROW},$C$11,{S3}!M3:M{N_FREG_LAST_ROW}),1)',
        f'=ROUND(SUMIF({S3}!A3:A{N_FREG_LAST_ROW},$C$11,{S3}!N3:N{N_FREG_LAST_ROW}),1)',
        f'=ROUND(SUMIF({S3}!A3:A{N_FREG_LAST_ROW},$C$11,{S3}!N3:N{N_FREG_LAST_ROW})'
        f'-SUMIF({S3}!A3:A{N_FREG_LAST_ROW},$C$11,{S3}!M3:M{N_FREG_LAST_ROW}),1)',
    ]
    fmts = ["#,##0", "#,##0", "#,##0", "#,##0.0", "#,##0.0", "+#,##0.0"]
    col0 = 4
    for i, (lbl, formula, fmt) in enumerate(zip(filtro_labels, filtro_formulas, fmts)):
        col = col0 + i
        ws.cell(row=10, column=col, value=lbl).font = LABEL_FONT
        cell = ws.cell(row=11, column=col, value=formula)
        cell.font = Font(name=FONT_NAME, bold=True, size=12, color=VERMELHO_ESC)
        cell.number_format = fmt
        cell.fill = KPI_FILL
        cell.alignment = Alignment(horizontal="center")
        cell.border = BORDER
    ws.row_dimensions[11].height = 24

    # ── Tabela: Análise por Distrito (fórmulas) ─────────────────────────
    style_section_title(ws, "B14:F14", "📊  ANÁLISE POR DISTRITO — CAMAS NECESSÁRIAS E POPULAÇÃO IDOSA")
    hdr_row = 15
    headers = ["Distrito", "Camas Necess. ANTES", "Camas Necess. DEPOIS",
               "Pop. 65+ ANTES", "Pop. 65+ DEPOIS"]
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
                value=f'=ROUND(SUMIF({S3}!$A$3:$A${N_FREG_LAST_ROW},$B{r},{S3}!$M$3:$M${N_FREG_LAST_ROW}),1)').border = BORDER
        ws.cell(row=r, column=4,
                value=f'=ROUND(SUMIF({S3}!$A$3:$A${N_FREG_LAST_ROW},$B{r},{S3}!$N$3:$N${N_FREG_LAST_ROW}),1)').border = BORDER
        ws.cell(row=r, column=5,
                value=f'=ROUND(SUMIF({S3}!$A$3:$A${N_FREG_LAST_ROW},$B{r},{S3}!$K$3:$K${N_FREG_LAST_ROW}),0)').border = BORDER
        ws.cell(row=r, column=6,
                value=f'=ROUND(SUMIF({S3}!$A$3:$A${N_FREG_LAST_ROW},$B{r},{S3}!$L$3:$L${N_FREG_LAST_ROW}),0)').border = BORDER
        for col in (2, 3, 4, 5, 6):
            ws.cell(row=r, column=col).alignment = Alignment(horizontal="center")
            if (i % 2) == 1:
                ws.cell(row=r, column=col).fill = PatternFill(start_color=CINZA, end_color=CINZA, fill_type="solid")
    dist_end_row = dist_start_row + len(DISTRITOS) - 1

    # ── Tabela: Top 15 ULS com maior agravamento do défice ──────────────
    style_section_title(ws, "H14:J14", "🏥  TOP 15 ULS — MAIOR AGRAVAMENTO DO DÉFICE")
    ws.cell(row=hdr_row, column=8, value="ULS").font = HEADER_FONT
    ws.cell(row=hdr_row, column=8).fill = HEADER_FILL
    ws.cell(row=hdr_row, column=8).border = BORDER
    ws.cell(row=hdr_row, column=9, value="Agravamento (camas)").font = HEADER_FONT
    ws.cell(row=hdr_row, column=9).fill = HEADER_FILL
    ws.cell(row=hdr_row, column=9).alignment = Alignment(horizontal="center", wrap_text=True)
    ws.cell(row=hdr_row, column=9).border = BORDER
    ws.merge_cells(start_row=hdr_row, start_column=9, end_row=hdr_row, end_column=10)

    top_uls_start = hdr_row + 1
    for i in range(15):
        src_row = 2 + i   # folha 2 já vem ordenada por Delta_Camas_Em_Falta desc
        r = top_uls_start + i
        ws.cell(row=r, column=8, value=f'={S2}!A{src_row}').border = BORDER
        val_cell = ws.cell(row=r, column=9, value=f'={S2}!I{src_row}')
        ws.merge_cells(start_row=r, start_column=9, end_row=r, end_column=10)
        val_cell.number_format = "+0.0"
        val_cell.alignment = Alignment(horizontal="center")
        val_cell.border = BORDER
        ws.cell(row=r, column=8).alignment = Alignment(horizontal="left")
        if i % 2 == 1:
            for col in (8, 9, 10):
                ws.cell(row=r, column=col).fill = PatternFill(start_color=CINZA, end_color=CINZA, fill_type="solid")
    top_uls_end = top_uls_start + 14

    # ── Tabela: Top 15 ULS por pior cobertura DEPOIS ────────────────────
    cov_title_row = max(dist_end_row, top_uls_end) + 3
    style_section_title(ws, f"B{cov_title_row}:F{cov_title_row}",
                         "📉  TOP 15 ULS COM PIOR COBERTURA DEPOIS DO CHOQUE DEMOGRÁFICO")
    cov_hdr = cov_title_row + 1
    headers_cov = ["ULS", "Cobertura % ANTES", "Cobertura % DEPOIS", "Camas em Falta ANTES", "Camas em Falta DEPOIS"]
    for j, h in enumerate(headers_cov, 2):
        c = ws.cell(row=cov_hdr, column=j, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[cov_hdr].height = 30
    cov_start = cov_hdr + 1

    # ranking auxiliar (única, sem empates) para obter as 15 piores coberturas DEPOIS
    ws2 = wb["2. Cobertura_ULS"]
    ws2["K1"] = "Rank_Cobertura_Depois"
    ws2["K1"].font = HEADER_FONT
    ws2["K1"].fill = HEADER_FILL
    for r in range(2, N_ULS_LAST_ROW + 1):
        ws2.cell(row=r, column=11, value=f"=F{r}-ROW()/100000000")   # -offset p/ desempate único, ordem ascendente

    for i in range(15):
        r = cov_start + i
        rank_formula = f'=SMALL({S2}!$K$2:$K${N_ULS_LAST_ROW},{i + 1})'
        helper_cell = f"$L{r}"
        ws.cell(row=r, column=12, value=rank_formula)
        ws.cell(row=r, column=2,
                value=f'=INDEX({S2}!$A$2:$A${N_ULS_LAST_ROW},MATCH({helper_cell},{S2}!$K$2:$K${N_ULS_LAST_ROW},0))').border = BORDER
        ws.cell(row=r, column=3,
                value=f'=INDEX({S2}!$E$2:$E${N_ULS_LAST_ROW},MATCH({helper_cell},{S2}!$K$2:$K${N_ULS_LAST_ROW},0))').border = BORDER
        ws.cell(row=r, column=3).number_format = '0.0"%"'
        ws.cell(row=r, column=4,
                value=f'=INDEX({S2}!$F$2:$F${N_ULS_LAST_ROW},MATCH({helper_cell},{S2}!$K$2:$K${N_ULS_LAST_ROW},0))').border = BORDER
        ws.cell(row=r, column=4).number_format = '0.0"%"'
        ws.cell(row=r, column=5,
                value=f'=INDEX({S2}!$G$2:$G${N_ULS_LAST_ROW},MATCH({helper_cell},{S2}!$K$2:$K${N_ULS_LAST_ROW},0))').border = BORDER
        ws.cell(row=r, column=6,
                value=f'=INDEX({S2}!$H$2:$H${N_ULS_LAST_ROW},MATCH({helper_cell},{S2}!$K$2:$K${N_ULS_LAST_ROW},0))').border = BORDER
        for col in (2, 3, 4, 5, 6):
            ws.cell(row=r, column=col).alignment = Alignment(horizontal="center" if col != 2 else "left")
            if i % 2 == 1:
                ws.cell(row=r, column=col).fill = PatternFill(start_color=CINZA, end_color=CINZA, fill_type="solid")
    cov_end = cov_start + 14
    ws.column_dimensions["L"].width = 2

    # ══════════════════════════════════════════════════════════════════
    # GRÁFICOS — pilha vertical, largura inteira, sem sobreposições
    # ══════════════════════════════════════════════════════════════════
    CHART_W_WIDE = 24
    CHART_H = 11
    ROWS_PER_CHART = 24

    def legend_bottom(chart):
        chart.legend.position = "b"
        chart.legend.overlay = False

    next_row = cov_end + 3

    # 1) Barras horizontais: Camas Necessárias por Distrito — Antes vs Depois
    style_section_title(ws, f"B{next_row}:J{next_row}",
                         "📈  GRÁFICO — Camas necessárias por Distrito (Antes vs. Depois)")
    chart1_row = next_row + 1
    chart1 = BarChart()
    chart1.type = "bar"
    chart1.grouping = "clustered"
    chart1.title = "Camas necessárias por Distrito"
    chart1.x_axis.title = "Camas"
    chart1.y_axis.title = "Distrito"
    chart1.style = 10
    data1 = Reference(ws, min_col=3, max_col=4, min_row=hdr_row, max_row=dist_end_row)
    cats1 = Reference(ws, min_col=2, min_row=dist_start_row, max_row=dist_end_row)
    chart1.add_data(data1, titles_from_data=True)
    chart1.set_categories(cats1)
    chart1.width = CHART_W_WIDE
    chart1.height = CHART_H + 4
    legend_bottom(chart1)
    ws.add_chart(chart1, f"B{chart1_row}")
    next_row = chart1_row + ROWS_PER_CHART + 6

    # 2) Barras horizontais: População 65+ por Distrito — Antes vs Depois
    style_section_title(ws, f"B{next_row}:J{next_row}",
                         "📈  GRÁFICO — População 65+ por Distrito (Antes vs. Depois)")
    chart2_row = next_row + 1
    chart2 = BarChart()
    chart2.type = "bar"
    chart2.grouping = "clustered"
    chart2.title = "População idosa (65+) por Distrito"
    chart2.x_axis.title = "Habitantes"
    chart2.y_axis.title = "Distrito"
    chart2.style = 11
    data2 = Reference(ws, min_col=5, max_col=6, min_row=hdr_row, max_row=dist_end_row)
    cats2 = Reference(ws, min_col=2, min_row=dist_start_row, max_row=dist_end_row)
    chart2.add_data(data2, titles_from_data=True)
    chart2.set_categories(cats2)
    chart2.width = CHART_W_WIDE
    chart2.height = CHART_H + 4
    legend_bottom(chart2)
    ws.add_chart(chart2, f"B{chart2_row}")
    next_row = chart2_row + ROWS_PER_CHART + 6

    # 3) Barras: Top 15 ULS com maior agravamento do défice
    style_section_title(ws, f"B{next_row}:J{next_row}",
                         "📈  GRÁFICO — Top 15 ULS com maior agravamento do défice (camas)")
    chart3_row = next_row + 1
    chart3 = BarChart()
    chart3.type = "bar"
    chart3.title = "Agravamento do défice de camas por ULS"
    chart3.x_axis.title = "Camas adicionais em falta"
    chart3.style = 12
    data3 = Reference(ws, min_col=9, max_col=9, min_row=hdr_row, max_row=top_uls_end)
    cats3 = Reference(ws, min_col=8, min_row=top_uls_start, max_row=top_uls_end)
    chart3.add_data(data3, titles_from_data=True)
    chart3.set_categories(cats3)
    chart3.width = CHART_W_WIDE
    chart3.height = CHART_H + 4
    legend_bottom(chart3)
    ws.add_chart(chart3, f"B{chart3_row}")
    next_row = chart3_row + ROWS_PER_CHART + 6

    # 4) Barras: Cobertura % Antes vs Depois — top 15 piores
    style_section_title(ws, f"B{next_row}:J{next_row}",
                         "📈  GRÁFICO — Cobertura (%) das 15 ULS mais críticas (Antes vs. Depois)")
    chart4_row = next_row + 1
    chart4 = BarChart()
    chart4.type = "bar"
    chart4.grouping = "clustered"
    chart4.title = "Cobertura de camas (%) — ULS mais críticas"
    chart4.x_axis.title = "Cobertura (%)"
    chart4.style = 13
    data4 = Reference(ws, min_col=3, max_col=4, min_row=cov_hdr, max_row=cov_end)
    cats4 = Reference(ws, min_col=2, min_row=cov_start, max_row=cov_end)
    chart4.add_data(data4, titles_from_data=True)
    chart4.set_categories(cats4)
    chart4.width = CHART_W_WIDE
    chart4.height = CHART_H + 4
    legend_bottom(chart4)
    ws.add_chart(chart4, f"B{chart4_row}")

    # esconder colunas auxiliares (N = lista distritos, L = rank helper)
    ws.column_dimensions["L"].hidden = True

    ws.freeze_panes = "B6"
    ws.sheet_view.zoomScale = 90
    wb.active = 0

    print("A guardar workbook...")
    wb.save(OUTPUT_FILE)
    print(f"✅ Guardado em: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()