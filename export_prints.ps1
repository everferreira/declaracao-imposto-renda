# Exporta prints reais da planilha via Excel COM (Windows)
$ErrorActionPreference = 'Stop'
$dir = 'C:\Users\conta\DIO\declaracao-imposto-renda'
$wbPath = "$dir\organizador-irpf.xlsx"
$inv = [System.Globalization.CultureInfo]::InvariantCulture

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($wbPath)
$xl.CalculateFull()

# --- Print 1: Menu ---
$ws = $wb.Worksheets.Item('Menu')
$ws.ExportAsFixedFormat(0, "$dir\prints\print_menu.pdf")

# --- Print 2: Rendimentos ---
$ws = $wb.Worksheets.Item('Rendimentos')
$ws.ExportAsFixedFormat(0, "$dir\prints\print_rendimentos.pdf")

# --- Print 3: Despesas ---
$ws = $wb.Worksheets.Item('Despesas')
$ws.ExportAsFixedFormat(0, "$dir\prints\print_despesas.pdf")

# --- Print 4: Bens ---
$ws = $wb.Worksheets.Item('Bens')
$ws.ExportAsFixedFormat(0, "$dir\prints\print_bens.pdf")

# --- Print 5: Resumo ---
$ws = $wb.Worksheets.Item('Resumo')
$ws.ExportAsFixedFormat(0, "$dir\prints\print_resumo.pdf")

# --- Print 6: Apoio ---
$ws = $wb.Worksheets.Item('Apoio')
$ws.ExportAsFixedFormat(0, "$dir\prints\print_apoio.pdf")

# --- Salva com os valores cacheados ---
$wb.Save()
$wb.Close($false)
$xl.Quit()
Write-Output 'EXPORT_OK'
