# Prepara dados\CONVATEC.csv para o painel.
# 1) Se o CONVATEC.csv da pasta CONVATEC do SharePoint (SIT - Gestao Operacional) for mais novo que a planilha, usa ele.
# 2) Senao, exporta a aba "Base por NF Oficial" da planilha (versao SALVA, aberta so para leitura).
$ErrorActionPreference = 'Stop'
$raiz = Split-Path -Parent $PSScriptRoot
$src = Join-Path (Split-Path -Parent $raiz) 'Base CONVATEC - com tracking.xlsm'
$out = Join-Path $raiz 'dados\CONVATEC.csv'
$share = Get-ChildItem (Join-Path $env:USERPROFILE 'OneDrive - CEVA Logistics') -Directory -Filter 'SIT - Gest*Operacional*' -ErrorAction SilentlyContinue | Select-Object -First 1
$csvShare = $null
if ($share) { $csvShare = Get-Item (Join-Path $share.FullName 'CONVATEC\CONVATEC.csv') -ErrorAction SilentlyContinue }
$xlsm = Get-Item $src -ErrorAction SilentlyContinue
if ($csvShare -and (-not $xlsm -or $csvShare.LastWriteTime -gt $xlsm.LastWriteTime)) {
  Copy-Item $csvShare.FullName $out -Force
  (Get-Item $out).LastWriteTime = $csvShare.LastWriteTime
  Write-Output ('Base do SharePoint (' + $csvShare.LastWriteTime.ToString('dd/MM/yyyy HH:mm') + ') -> ' + $out)
  exit 0
}
if (-not $xlsm) { throw ('Planilha nao encontrada: ' + $src) }
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.AutomationSecurity = 3
$linhas = New-Object System.Collections.Generic.List[string]
try {
  $wb = $xl.Workbooks.Open($src, 0, $true)
  $d = $wb.Worksheets.Item('Base por NF Oficial')
  $last = $d.Cells.Item($d.Rows.Count, 1).End(-4162).Row
  $lc = $d.Cells.Item(1, $d.Columns.Count).End(-4159).Column
  for ($r = 1; $r -le $last; $r++) {
    $c = @(); for ($k = 1; $k -le $lc; $k++) { $c += (($d.Cells.Item($r, $k).Text) -replace '[;\r\n]', ' ') }
    $linhas.Add(($c -join ';'))
  }
  $wb.Close($false)
} finally { $xl.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null }
[System.IO.File]::WriteAllLines($out, $linhas, (New-Object System.Text.UTF8Encoding($true)))
Write-Output ('Base exportada da planilha: ' + ($last - 1) + ' notas, ' + $lc + ' colunas -> ' + $out)
