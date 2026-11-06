#!/usr/bin/env python3
"""
Gera a planilha Organizador de Declaração de Imposto de Renda
Desafio DIO - Formação Análise de Dados com Excel e IA
"""

import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side
)
from openpyxl.worksheet.datavalidation import DataValidation

# ── Cores e estilos ──────────────────────────────────────────────
COR_PRIMARIA  = "1F4E79"
COR_SECUNDARIA = "D6E4F0"
COR_DESTAQUE  = "E8B84B"
COR_ENTRADA   = "FFF2CC"
COR_CALC      = "F2F2F2"
COR_TEXTO     = "333333"

FONT_TITLE  = Font(name="Calibri", size=20, bold=True, color="FFFFFF")
FONT_SUB    = Font(name="Calibri", size=14, bold=True, color=COR_PRIMARIA)
FONT_HEADER = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
FONT_BODY   = Font(name="Calibri", size=11, color=COR_TEXTO)
FONT_LINK   = Font(name="Calibri", size=11, color="0563C1", underline="single")
FONT_VALOR  = Font(name="Calibri", size=11, bold=True, color=COR_PRIMARIA)
FONT_VALOR_GRANDE = Font(name="Calibri", size=18, bold=True, color=COR_PRIMARIA)

FILL_PRIMARIA   = PatternFill("solid", fgColor=COR_PRIMARIA)
FILL_SECUNDARIA = PatternFill("solid", fgColor=COR_SECUNDARIA)
FILL_DESTAQUE   = PatternFill("solid", fgColor=COR_DESTAQUE)
FILL_ENTRADA    = PatternFill("solid", fgColor=COR_ENTRADA)
FILL_CALC       = PatternFill("solid", fgColor=COR_CALC)
FILL_VERDE      = PatternFill("solid", fgColor="C6EFCE")
FILL_VERMELHO   = PatternFill("solid", fgColor="FFC7CE")

ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT   = Alignment(horizontal="left", vertical="center", wrap_text=True)
ALIGN_RIGHT  = Alignment(horizontal="right", vertical="center")

THIN_BORDER = Border(
    left=Side(style="thin", color="CCCCCC"),
    right=Side(style="thin", color="CCCCCC"),
    top=Side(style="thin", color="CCCCCC"),
    bottom=Side(style="thin", color="CCCCCC"),
)

FMT_MOEDA = '#,##0.00'
FMT_DATA  = 'DD/MM/AAAA'
FMT_PCT   = '0%'

# ── Helpers ──────────────────────────────────────────────────────

def style_header_row(ws, row, max_col, fill=FILL_PRIMARIA, font=FONT_HEADER):
    for c in range(1, max_col + 1):
        cell = ws.cell(row, c)
        cell.font = font
        cell.fill = fill
        cell.alignment = ALIGN_CENTER
        cell.border = THIN_BORDER

def style_input_cell(cell, fmt=None):
    cell.fill = FILL_ENTRADA
    cell.font = FONT_BODY
    cell.border = THIN_BORDER
    cell.alignment = ALIGN_LEFT
    if fmt:
        cell.number_format = fmt

def style_calc_cell(cell, fmt=FMT_MOEDA):
    cell.fill = FILL_CALC
    cell.font = FONT_VALOR
    cell.number_format = fmt
    cell.border = THIN_BORDER
    cell.alignment = ALIGN_RIGHT

def add_title(ws, title, max_col, start_row=1):
    ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=max_col)
    cell = ws.cell(start_row, 1, title)
    cell.font = FONT_TITLE
    cell.fill = FILL_PRIMARIA
    cell.alignment = ALIGN_CENTER
    ws.row_dimensions[start_row].height = 45
    return start_row + 1

def add_subtitle(ws, subtitle, max_col, row):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=max_col)
    cell = ws.cell(row, 1, subtitle)
    cell.font = FONT_SUB
    cell.fill = FILL_SECUNDARIA
    cell.alignment = ALIGN_LEFT
    ws.row_dimensions[row].height = 28
    return row + 1

def add_nav_bar(ws, row, max_col, links):
    c = 1
    for nome, ref in links:
        cell = ws.cell(row, c, nome)
        cell.hyperlink = ref
        cell.font = FONT_LINK
        cell.fill = FILL_SECUNDARIA
        cell.alignment = ALIGN_CENTER
        cell.border = THIN_BORDER
        c += 1
    for cc in range(c, max_col + 1):
        cell = ws.cell(row, cc)
        cell.fill = FILL_SECUNDARIA
        cell.border = THIN_BORDER
    ws.row_dimensions[row].height = 25
    return row + 1


# ═══════════════════════════════════════════════════════════════════
#  CONSTRUÇÃO
# ═══════════════════════════════════════════════════════════════════

wb = openpyxl.Workbook()

# ── ABA 1: MENU ──────────────────────────────────────────────────
ws_menu = wb.active
ws_menu.title = "Menu"
ws_menu.sheet_view.showGridLines = False

max_col = 8
r = 1
r = add_title(ws_menu, "📋 Organizador de Declaração de Imposto de Renda", max_col, r)
r = add_nav_bar(ws_menu, r, max_col, [
    ("🏠 Menu", "#Menu!A1"),
    ("💰 Rendimentos", "#Rendimentos!A1"),
    ("💳 Despesas", "#Despesas!A1"),
    ("🏡 Bens", "#Bens!A1"),
    ("📊 Resumo", "#Resumo!A1"),
    ("📖 Apoio", "#Apoio!A1"),
])
r += 1

# Cards de resumo (cada card = 2 cells lado a lado, label em cima, valor embaixo)
cards = [
    ("💰 Rendimentos", "=Resumo!B2", FILL_VERDE),
    ("💳 Despesas", "=Resumo!B8", FILL_VERMELHO),
    ("🏡 Bens", "=Resumo!B12", FILL_SECUNDARIA),
    ("📊 Base Cálculo", "=Resumo!B16", FILL_DESTAQUE),
]

for idx, (label, formula, fill) in enumerate(cards):
    col = 1 + idx * 2
    # Label
    c_lbl = ws_menu.cell(r, col, label)
    c_lbl.font = Font(name="Calibri", size=12, bold=True, color=COR_TEXTO)
    c_lbl.fill = fill
    c_lbl.alignment = ALIGN_CENTER
    c_lbl.border = THIN_BORDER
    ws_menu.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col+1)

    # Valor (linha abaixo)
    c_val = ws_menu.cell(r+1, col, formula)
    c_val.font = FONT_VALOR_GRANDE
    c_val.fill = fill
    c_val.alignment = ALIGN_CENTER
    c_val.number_format = FMT_MOEDA
    c_val.border = THIN_BORDER
    ws_menu.merge_cells(start_row=r+1, start_column=col, end_row=r+1, end_column=col+1)

    # Fills pros merged cells
    for rr in range(r, r+2):
        for cc in range(col, col+2):
            mc = ws_menu.cell(rr, cc)
            mc.fill = fill
            mc.border = THIN_BORDER

ws_menu.row_dimensions[r].height = 28
ws_menu.row_dimensions[r+1].height = 40

r += 3

# Instruções
ws_menu.merge_cells(start_row=r, start_column=1, end_row=r, end_column=max_col)
c = ws_menu.cell(r, 1, "📌 Como usar:")
c.font = FONT_SUB
r += 1

instrucoes = [
    "1. Preencha seus rendimentos na aba 💰 Rendimentos (salário, freelas, aluguéis)",
    "2. Registre despesas dedutíveis na aba 💳 Despesas (médicas, educação, etc.)",
    "3. Liste seus bens e direitos na aba 🏡 Bens",
    "4. Acompanhe o consolidado automático na aba 📊 Resumo",
    "5. Consulte tabelas e limites na aba 📖 Apoio",
]
for inst in instrucoes:
    ws_menu.merge_cells(start_row=r, start_column=1, end_row=r, end_column=max_col)
    c = ws_menu.cell(r, 1, inst)
    c.font = FONT_BODY
    c.alignment = ALIGN_LEFT
    r += 1

r += 1
ws_menu.merge_cells(start_row=r, start_column=1, end_row=r, end_column=max_col)
c = ws_menu.cell(r, 1, "🔗 Links Rápidos")
c.font = FONT_SUB
r += 1

links_rapidos = [
    ("Programa IRPF 2026 - Receita Federal", "https://www.gov.br/receitafederal/pt-br/assuntos/imposto-de-renda"),
    ("Tabela Progressiva IRPF", "https://www.gov.br/receitafederal/pt-br/acesso-a-informacao/perguntas-frequentes/tabela-incidencia"),
    ("Consulta Restituição", "https://www.gov.br/receitafederal/pt-br/assuntos/imposto-de-renda/pagamento-e-restituicao"),
    ("e-CAC - Centro de Atendimento Virtual", "https://www.gov.br/receitafederal/pt-br/assuntos/meu-imposto-de-renda/e-cac"),
]
for nome_link, url in links_rapidos:
    c = ws_menu.cell(r, 2, f"  ► {nome_link}")
    c.hyperlink = url
    c.font = FONT_LINK
    c.alignment = ALIGN_LEFT
    r += 1

r += 1
ws_menu.merge_cells(start_row=r, start_column=1, end_row=r, end_column=max_col)
c = ws_menu.cell(r, 1, "Células em amarelo = preenchimento manual | Cinza = calculadas automaticamente")
c.font = Font(name="Calibri", size=9, italic=True, color="888888")

# Larguras
for c_letra in 'ABCDEFGH':
    ws_menu.column_dimensions[c_letra].width = 20

# ── ABA 2: RENDIMENTOS ──────────────────────────────────────────
ws_rend = wb.create_sheet("Rendimentos")
r = 1
r = add_title(ws_rend, "💰 Rendimentos", max_col, r)
r = add_nav_bar(ws_rend, r, max_col, [
    ("🏠 Menu", "#Menu!A1"),
    ("💰 Rendimentos", "#Rendimentos!A1"),
    ("💳 Despesas", "#Despesas!A1"),
    ("🏡 Bens", "#Bens!A1"),
    ("📊 Resumo", "#Resumo!A1"),
    ("📖 Apoio", "#Apoio!A1"),
])
r += 1

headers_rend = ["Data", "Descrição", "Tipo", "Valor (R$)", "IR Retido (R$)", "Fonte Pagadora"]
for c, h in enumerate(headers_rend, 1):
    ws_rend.cell(r, c, h)
style_header_row(ws_rend, r, len(headers_rend))
r_header_rend = r
r += 1

dv_tipo_rend = DataValidation(
    type="list",
    formula1='"Tributável,Isento,Tributável Exclusivo Definitivo,Rendimento de Aluguel,Pensão Alimentícia"',
    allow_blank=True
)
dv_tipo_rend.error = "Selecione um tipo válido"
dv_tipo_rend.errorTitle = "Tipo inválido"
ws_rend.add_data_validation(dv_tipo_rend)

exemplos_rend = [
    ("01/01/2026", "Salário - Ctrl+Play", "Tributável", 5500.00, 687.50, "Ctrl+Play Escola de Programação LTDA"),
    ("15/01/2026", "Aula Particular - Freelance", "Tributável", 1200.00, 0, "Autônomo"),
    ("20/01/2026", "Dividendos - Investimentos", "Isento", 350.00, 0, "B3 S.A."),
    ("05/02/2026", "Salário - Ctrl+Play", "Tributável", 5500.00, 687.50, "Ctrl+Play Escola de Programação LTDA"),
    ("10/02/2026", "Aluguel Recebido", "Rendimento de Aluguel", 1800.00, 0, "Inquilino - Maria Santos"),
    ("28/02/2026", "Rendimento Poupança", "Isento", 45.00, 0, "Banco do Brasil"),
    ("05/03/2026", "Salário - Ctrl+Play", "Tributável", 5500.00, 687.50, "Ctrl+Play Escola de Programação LTDA"),
    ("15/03/2026", "Curso Ministrado - Workshop IA", "Tributável", 2500.00, 0, "Autônomo"),
    ("01/04/2026", "Salário - Ctrl+Play", "Tributável", 5500.00, 687.50, "Ctrl+Play Escola de Programação LTDA"),
    ("05/04/2026", "Bolsa de Estudo - DIO", "Isento", 0, 0, "Digital Innovation One"),
]

for row_data in exemplos_rend:
    for c, val in enumerate(row_data, 1):
        cell = ws_rend.cell(r, c)
        if c == 1:
            cell.value = val
            style_input_cell(cell, FMT_DATA)
        elif c in (4, 5):
            cell.value = val
            style_input_cell(cell, FMT_MOEDA)
        else:
            cell.value = val
            style_input_cell(cell)
    dv_tipo_rend.add(ws_rend.cell(r, 3))
    r += 1

# Linha de totais
ws_rend.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
cell = ws_rend.cell(r, 1, "TOTAIS:")
cell.font = FONT_HEADER
cell.fill = FILL_PRIMARIA
cell.alignment = ALIGN_RIGHT
cell.border = THIN_BORDER
ws_rend.cell(r, 2).fill = FILL_PRIMARIA
ws_rend.cell(r, 2).border = THIN_BORDER

# Soma valores
cell_val = ws_rend.cell(r, 4)
cell_val.value = f"=SUM(D{r_header_rend+1}:D{r-1})"
style_calc_cell(cell_val)

# Soma IR retido
cell_ir = ws_rend.cell(r, 5)
cell_ir.value = f"=SUM(E{r_header_rend+1}:E{r-1})"
style_calc_cell(cell_ir)

# Células vazias
for cc in [3, 6]:
    ws_rend.cell(r, cc).fill = FILL_CALC
    ws_rend.cell(r, cc).border = THIN_BORDER

r_total_rend = r

# Larguras
ws_rend.column_dimensions['A'].width = 14
ws_rend.column_dimensions['B'].width = 32
ws_rend.column_dimensions['C'].width = 24
ws_rend.column_dimensions['D'].width = 16
ws_rend.column_dimensions['E'].width = 16
ws_rend.column_dimensions['F'].width = 32

# ── ABA 3: DESPESAS ──────────────────────────────────────────────
ws_desp = wb.create_sheet("Despesas")
r = 1
r = add_title(ws_desp, "💳 Despesas Dedutíveis", max_col, r)
r = add_nav_bar(ws_desp, r, max_col, [
    ("🏠 Menu", "#Menu!A1"),
    ("💰 Rendimentos", "#Rendimentos!A1"),
    ("💳 Despesas", "#Despesas!A1"),
    ("🏡 Bens", "#Bens!A1"),
    ("📊 Resumo", "#Resumo!A1"),
    ("📖 Apoio", "#Apoio!A1"),
])
r += 1

headers_desp = ["Data", "Descrição", "Categoria", "Valor (R$)", "CNPJ/CPF Prestador", "Observações"]
for c, h in enumerate(headers_desp, 1):
    ws_desp.cell(r, c, h)
style_header_row(ws_desp, r, len(headers_desp))
r_header_desp = r
r += 1

dv_cat_desp = DataValidation(
    type="list",
    formula1='"Médica,Educação,Dependente,Previdência Privada,Pensão Alimentícia,Doação,Outros"',
    allow_blank=True
)
ws_desp.add_data_validation(dv_cat_desp)

exemplos_desp = [
    ("10/01/2026", "Plano de Saúde Unimed", "Médica", 450.00, "00.000.000/0001-91", "Plano empresarial"),
    ("15/01/2026", "Curso de Especialização - IA", "Educação", 1890.00, "00.000.000/0002-72", "Pós-graduação"),
    ("20/01/2026", "Consulta Oftalmológica", "Médica", 250.00, "000.000.000-00", "Dra. Ana Oliveira"),
    ("05/02/2026", "Material Didático", "Educação", 180.00, "", "Livros técnicos"),
    ("10/02/2026", "Pensão Alimentícia", "Pensão Alimentícia", 1200.00, "", "Conforme acordo judicial"),
    ("15/02/2026", "Exames Laboratoriais", "Médica", 320.00, "00.000.000/0003-53", "Lab Saúde"),
    ("05/03/2026", "Curso Online - Data Science", "Educação", 2500.00, "00.000.000/0004-34", "DIO"),
    ("10/03/2026", "Plano de Saúde Unimed", "Médica", 450.00, "00.000.000/0001-91", "Mensalidade"),
    ("20/03/2026", "Previdência Privada", "Previdência Privada", 500.00, "00.000.000/0005-15", "BrasilPrev"),
    ("05/04/2026", "Consulta Psicóloga", "Médica", 200.00, "000.000.000-01", "Dr. Carlos Mendes"),
]

for row_data in exemplos_desp:
    for c, val in enumerate(row_data, 1):
        cell = ws_desp.cell(r, c)
        if c == 1:
            cell.value = val
            style_input_cell(cell, FMT_DATA)
        elif c == 4:
            cell.value = val
            style_input_cell(cell, FMT_MOEDA)
        else:
            cell.value = val
            style_input_cell(cell)
    dv_cat_desp.add(ws_desp.cell(r, 3))
    r += 1

# Totais
ws_desp.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
cell = ws_desp.cell(r, 1, "TOTAIS:")
cell.font = FONT_HEADER
cell.fill = FILL_PRIMARIA
cell.alignment = ALIGN_RIGHT
cell.border = THIN_BORDER
ws_desp.cell(r, 2).fill = FILL_PRIMARIA
ws_desp.cell(r, 2).border = THIN_BORDER

cell_cat = ws_desp.cell(r, 3, "Total Geral")
cell_cat.font = FONT_HEADER
cell_cat.fill = FILL_PRIMARIA
cell_cat.alignment = ALIGN_CENTER
cell_cat.border = THIN_BORDER

cell_val = ws_desp.cell(r, 4)
cell_val.value = f"=SUM(D{r_header_desp+1}:D{r-1})"
style_calc_cell(cell_val)

ws_desp.cell(r, 5).fill = FILL_CALC
ws_desp.cell(r, 5).border = THIN_BORDER
ws_desp.cell(r, 6).fill = FILL_CALC
ws_desp.cell(r, 6).border = THIN_BORDER

r_total_desp = r

# Larguras
ws_desp.column_dimensions['A'].width = 14
ws_desp.column_dimensions['B'].width = 30
ws_desp.column_dimensions['C'].width = 20
ws_desp.column_dimensions['D'].width = 16
ws_desp.column_dimensions['E'].width = 22
ws_desp.column_dimensions['F'].width = 25

# ── ABA 4: BENS ──────────────────────────────────────────────────
ws_bens = wb.create_sheet("Bens")
r = 1
r = add_title(ws_bens, "🏡 Bens e Direitos", max_col, r)
r = add_nav_bar(ws_bens, r, max_col, [
    ("🏠 Menu", "#Menu!A1"),
    ("💰 Rendimentos", "#Rendimentos!A1"),
    ("💳 Despesas", "#Despesas!A1"),
    ("🏡 Bens", "#Bens!A1"),
    ("📊 Resumo", "#Resumo!A1"),
    ("📖 Apoio", "#Apoio!A1"),
])
r += 1

headers_bens = ["Tipo", "Descrição", "Valor Aquisição (R$)", "Valor Atual (R$)", "Data Aquisição", "% Patrimônio"]
for c, h in enumerate(headers_bens, 1):
    ws_bens.cell(r, c, h)
style_header_row(ws_bens, r, len(headers_bens))
r_header_bens = r
r += 1

dv_tipo_bem = DataValidation(
    type="list",
    formula1='"Imóvel,Veículo,Investimento,Título Público,Fundos de Investimento,Poupança,Ação,Previdência Privada,Outros"',
    allow_blank=True
)
ws_bens.add_data_validation(dv_tipo_bem)

exemplos_bens = [
    ("Im��vel", "Apartamento 2 quartos - Vila Mariana", 350000.00, 420000.00, "15/03/2020", 0.50),
    ("Veículo", "SUV Chevrolet Tracker 2022", 125000.00, 95000.00, "10/01/2022", 0.15),
    ("Investimento", "Tesouro Selic 2029", 50000.00, 53200.00, "05/06/2023", 0.08),
    ("Poupança", "Poupança Banco do Brasil", 25000.00, 26250.00, "01/01/2020", 0.04),
    ("Fundos de Investimento", "Fundo Multimercado XP", 35000.00, 38500.00, "20/08/2021", 0.06),
    ("Previdência Privada", "PGBL BrasilPrev", 30000.00, 32500.00, "10/03/2022", 0.05),
    ("Ação", "PETR4 - Petrobras (500 ações)", 15000.00, 18500.00, "15/01/2024", 0.03),
    ("Título Público", "NTN-B Principal 2035", 40000.00, 43800.00, "01/11/2023", 0.07),
    ("Imóvel", "Terreno - Interior SP", 80000.00, 95000.00, "20/12/2021", 0.12),
]

for row_data in exemplos_bens:
    for c, val in enumerate(row_data, 1):
        cell = ws_bens.cell(r, c)
        if c in (3, 4):
            cell.value = val
            style_input_cell(cell, FMT_MOEDA)
        elif c == 5:
            cell.value = val
            style_input_cell(cell, FMT_DATA)
        elif c == 6:
            cell.value = val
            style_input_cell(cell, FMT_PCT)
        else:
            cell.value = val
            style_input_cell(cell)
    dv_tipo_bem.add(ws_bens.cell(r, 1))
    r += 1

# Totais
ws_bens.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
cell = ws_bens.cell(r, 1, "TOTAIS:")
cell.font = FONT_HEADER
cell.fill = FILL_PRIMARIA
cell.alignment = ALIGN_RIGHT
cell.border = THIN_BORDER
ws_bens.cell(r, 2).fill = FILL_PRIMARIA
ws_bens.cell(r, 2).border = THIN_BORDER

cell_adq = ws_bens.cell(r, 3)
cell_adq.value = f"=SUM(C{r_header_bens+1}:C{r-1})"
style_calc_cell(cell_adq)

cell_atual = ws_bens.cell(r, 4)
cell_atual.value = f"=SUM(D{r_header_bens+1}:D{r-1})"
style_calc_cell(cell_atual)

for cc in [5, 6]:
    ws_bens.cell(r, cc).fill = FILL_CALC
    ws_bens.cell(r, cc).border = THIN_BORDER

r_total_bens = r

# Larguras
ws_bens.column_dimensions['A'].width = 22
ws_bens.column_dimensions['B'].width = 32
ws_bens.column_dimensions['C'].width = 20
ws_bens.column_dimensions['D'].width = 18
ws_bens.column_dimensions['E'].width = 16
ws_bens.column_dimensions['F'].width = 14

# ── ABA 5: RESUMO ────────────────────────────────────────────────
ws_resumo = wb.create_sheet("Resumo")
r = 1
r = add_title(ws_resumo, "📊 Resumo Consolidado", max_col, r)
r = add_nav_bar(ws_resumo, r, max_col, [
    ("🏠 Menu", "#Menu!A1"),
    ("💰 Rendimentos", "#Rendimentos!A1"),
    ("💳 Despesas", "#Despesas!A1"),
    ("🏡 Bens", "#Bens!A1"),
    ("📊 Resumo", "#Resumo!A1"),
    ("📖 Apoio", "#Apoio!A1"),
])
r += 2

# Cabeçalho da tabela
for c in range(1, max_col + 1):
    ws_resumo.cell(r, c).border = Border()
    ws_resumo.cell(r, c).fill = PatternFill()
ws_resumo.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
c = ws_resumo.cell(r, 1, "Indicador")
c.font = FONT_HEADER
c.fill = FILL_PRIMARIA
c.alignment = ALIGN_CENTER
c.border = THIN_BORDER
ws_resumo.cell(r, 2).fill = FILL_PRIMARIA
ws_resumo.cell(r, 2).border = THIN_BORDER

ws_resumo.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
c = ws_resumo.cell(r, 3, "Valor (R$)")
c.font = FONT_HEADER
c.fill = FILL_PRIMARIA
c.alignment = ALIGN_CENTER
c.border = THIN_BORDER
ws_resumo.cell(r, 4).fill = FILL_PRIMARIA
ws_resumo.cell(r, 4).border = THIN_BORDER
r += 1

# Linhas do resumo
resumo_items = [
    ("💰 RENDIMENTOS", None, None, FILL_SECUNDARIA, True),
    ("   Total de Rendimentos", f"=Rendimentos!D{r_total_rend}", FMT_MOEDA, FILL_VERDE, False),
    ("   IR Retido na Fonte", f"=Rendimentos!E{r_total_rend}", FMT_MOEDA, FILL_CALC, False),
    ("", None, None, None, False),
    ("💳 DESPESAS", None, None, FILL_SECUNDARIA, True),
    ("   Total de Despesas Dedutíveis", f"=Despesas!D{r_total_desp}", FMT_MOEDA, FILL_VERMELHO, False),
    ("", None, None, None, False),
    ("🏡 BENS E DIREITOS", None, None, FILL_SECUNDARIA, True),
    ("   Total de Bens (Valor Atual)", f"=Bens!D{r_total_bens}", FMT_MOEDA, FILL_SECUNDARIA, False),
    ("   Total de Bens (Aquisição)", f"=Bens!C{r_total_bens}", FMT_MOEDA, FILL_CALC, False),
    ("", None, None, None, False),
    ("📊 BASE DE CÁLCULO ESTIMADA", f"=B{6}-B{10}", FMT_MOEDA, FILL_DESTAQUE, True),
]

for i, (label, formula, fmt, fill, is_header) in enumerate(resumo_items):
    row = r + i

    if not label and not formula:
        ws_resumo.row_dimensions[row].height = 8
        continue

    if is_header and label:
        ws_resumo.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
        ws_resumo.merge_cells(start_row=row, start_column=3, end_row=row, end_column=4)
        c = ws_resumo.cell(row, 1, label)
        c.font = Font(name="Calibri", size=11, bold=True, color=COR_TEXTO)
        c.fill = fill
        c.alignment = ALIGN_LEFT
        c.border = THIN_BORDER
        for cc in [2, 3, 4]:
            ws_resumo.cell(row, cc).fill = fill
            ws_resumo.cell(row, cc).border = THIN_BORDER
        continue

    # Linha normal
    ws_resumo.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
    c_lbl = ws_resumo.cell(row, 1, label)
    c_lbl.font = Font(name="Calibri", size=10, color="666666" if fill == FILL_CALC else COR_TEXTO)
    c_lbl.alignment = ALIGN_LEFT
    c_lbl.border = Border(bottom=Side(style="thin", color="EEEEEE"))
    ws_resumo.cell(row, 2).border = Border(bottom=Side(style="thin", color="EEEEEE"))

    if formula:
        c_val = ws_resumo.cell(row, 3, formula)
        c_val.number_format = fmt
        c_val.font = Font(name="Calibri", size=10, color="666666" if fill == FILL_CALC else COR_TEXTO)
        c_val.alignment = ALIGN_RIGHT
        c_val.border = Border(bottom=Side(style="thin", color="EEEEEE"))
        ws_resumo.merge_cells(start_row=row, start_column=3, end_row=row, end_column=4)
        ws_resumo.cell(row, 4).border = Border(bottom=Side(style="thin", color="EEEEEE"))

ws_resumo.column_dimensions['A'].width = 4
ws_resumo.column_dimensions['B'].width = 36
ws_resumo.column_dimensions['C'].width = 22
ws_resumo.column_dimensions['D'].width = 4
for c_letra in 'EFGH':
    ws_resumo.column_dimensions[c_letra].width = 15

# ── ABA 6: APOIO ─────────────────────────────────────────────────
ws_apoio = wb.create_sheet("Apoio")
r = 1
r = add_title(ws_apoio, "📖 Tabelas de Apoio", max_col, r)
r = add_nav_bar(ws_apoio, r, max_col, [
    ("🏠 Menu", "#Menu!A1"),
    ("💰 Rendimentos", "#Rendimentos!A1"),
    ("💳 Despesas", "#Despesas!A1"),
    ("🏡 Bens", "#Bens!A1"),
    ("📊 Resumo", "#Resumo!A1"),
    ("📖 Apoio", "#Apoio!A1"),
])
r += 1

r = add_subtitle(ws_apoio, "Tabela Progressiva do IRPF - 2026", max_col, r)
r += 1

headers_ir = ["Faixa", "Base de Cálculo (R$)", "Alíquota", "Parcela a Deduzir (R$)"]
for c, h in enumerate(headers_ir, 1):
    ws_apoio.cell(r, c, h)
style_header_row(ws_apoio, r, len(headers_ir))
r += 1

faixas_ir = [
    ("1ª Faixa", "Até 2.259,20", "Isento", "0,00"),
    ("2ª Faixa", "De 2.259,21 até 2.828,65", "7,5%", "169,44"),
    ("3ª Faixa", "De 2.828,66 até 3.751,05", "15%", "381,44"),
    ("4ª Faixa", "De 3.751,06 até 4.664,68", "22,5%", "662,77"),
    ("5ª Faixa", "Acima de 4.664,68", "27,5%", "896,00"),
]

for row_data in faixas_ir:
    for c, val in enumerate(row_data, 1):
        cell = ws_apoio.cell(r, c, val)
        cell.font = FONT_BODY
        cell.border = THIN_BORDER
        cell.alignment = ALIGN_CENTER
    r += 1

r += 2
r = add_subtitle(ws_apoio, "Limites de Dedução por Categoria", max_col, r)
r += 1

headers_lim = ["Categoria", "Limite Anual (R$)", "Observação"]
for c, h in enumerate(headers_lim, 1):
    ws_apoio.cell(r, c, h)
style_header_row(ws_apoio, r, len(headers_lim))
r += 1

limites = [
    ("Despesas Médicas", "Sem limite", "Todas as despesas comprovadas são dedutíveis"),
    ("Despesas com Educação", "3.561,50", "Por dependente (cursos, faculdade, pós, técnico)"),
    ("Previdência Privada (PGBL)", "12% da renda bruta", "Apenas PGBL, não VGBL"),
    ("Pensão Alimentícia", "Sem limite", "Conforme decisão judicial"),
    ("Dependentes", "2.275,08", "Por dependente na declaração"),
]

for row_data in limites:
    for c, val in enumerate(row_data, 1):
        cell = ws_apoio.cell(r, c, val)
        cell.font = FONT_BODY
        cell.border = THIN_BORDER
        cell.alignment = ALIGN_RIGHT if c == 2 else ALIGN_LEFT
    r += 1

r += 2
r = add_subtitle(ws_apoio, "Links Úteis", max_col, r)
r += 1

links_apoio = [
    ("Programa Gerador da Declaração (PGD) IRPF", "https://www.gov.br/receitafederal/pt-br/centrais-de-conteudo/download/pgd"),
    ("Consulta de Malha Fiscal", "https://www.gov.br/receitafederal/pt-br/assuntos/meu-imposto-de-renda/malha-fiscal"),
    ("e-CAC Centro de Atendimento Virtual", "https://www.gov.br/receitafederal/pt-br/assuntos/meu-imposto-de-renda/e-cac"),
    ("Tabela de Incidência do IRRF", "https://www.gov.br/receitafederal/pt-br/acesso-a-informacao/perguntas-frequentes/tabela-incidencia"),
    ("Cartilha do Contribuinte - IRPF", "https://www.gov.br/receitafederal/pt-br/assuntos/imposto-de-renda"),
]
for nome_link, url in links_apoio:
    c = ws_apoio.cell(r, 1, f"  ► {nome_link}")
    c.hyperlink = url
    c.font = FONT_LINK
    c.alignment = ALIGN_LEFT
    r += 1

# Larguras
ws_apoio.column_dimensions['A'].width = 30
ws_apoio.column_dimensions['B'].width = 28
ws_apoio.column_dimensions['C'].width = 45
for c_letra in 'DEFGH':
    ws_apoio.column_dimensions[c_letra].width = 15

# ── Configuração de impressão ────────────────────────────────────
for ws in wb.worksheets:
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_setup.paperSize = ws.PAPERSIZE_A4

# ── Salvar ────────────────────────────────────────────────────────
output_path = "/mnt/c/Users/conta/DIO/declaracao-imposto-renda/organizador-irpf.xlsx"
wb.save(output_path)
print(f"✅ Planilha salva em: {output_path}")
print(f"   Abas: {wb.sheetnames}")
