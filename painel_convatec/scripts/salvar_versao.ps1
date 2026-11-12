# Salva a ultima versao do CODIGO (macros, painel, scripts) no repositorio Git e faz commit.
# Nao copia dados: planilha, CSV e paineis com notas ficam fora (o .gitignore tambem bloqueia).
# Uso: SALVAR VERSAO NO GIT.bat            -> commit local
#      SALVAR VERSAO NO GIT.bat enviar     -> commit e envio ao GitHub
param([string]$modo = '')
$ErrorActionPreference = 'Stop'
$raiz = Split-Path -Parent $PSScriptRoot
$cov = Split-Path -Parent $raiz
$cfg = Get-Content -Raw -Encoding UTF8 (Join-Path $raiz 'config.json') | ConvertFrom-Json
$repo = $cfg.repo_publico
if (-not $repo) { $repo = 'C:\Users\tannete\repos\ceva-convatec' }
if (-not (Test-Path (Join-Path $repo '.git'))) { throw ('Repositorio nao encontrado: ' + $repo) }
$mapa = @(
  @('modBaseWV_v12_CONVATEC.bas', 'macros'),
  @('modTrackingCONVATEC.bas', 'macros'),
  @('evento_BasePorNF_Oficial.bas.txt', 'macros')
)
foreach ($m in $mapa) { Copy-Item (Join-Path $cov $m[0]) (Join-Path $repo $m[1]) -Force }
$pc = Join-Path $repo 'painel_convatec'
foreach ($f in @('config.json', 'README.md', 'CLAUDE.md', 'ATUALIZAR PAINEL.bat', 'SALVAR VERSAO NO GIT.bat')) { if (Test-Path (Join-Path $raiz $f)) { Copy-Item (Join-Path $raiz $f) $pc -Force } }
Copy-Item (Join-Path $raiz 'painel\template.html') (Join-Path $pc 'painel') -Force
Get-ChildItem (Join-Path $raiz 'scripts') -File | Where-Object { $_.Extension -in '.py', '.ps1' } | ForEach-Object { Copy-Item $_.FullName (Join-Path $pc 'scripts') -Force }
$pub = Join-Path $raiz $cfg.publico
if (Test-Path $pub) { Copy-Item $pub (Join-Path $repo 'index.html') -Force }
# trava: nada de e-mail real nem notas no que vai para o Git (este script fica fora da busca: o padrao abaixo casaria com ele mesmo)
$ruim = git -C $repo grep -I -l -E '@cevalogistics|@tecmar|"rows":\[\[' -- . ':!painel_convatec/scripts/salvar_versao.ps1' 2>$null
if ($ruim) { throw ('Parei: dado sensivel em ' + ($ruim -join ', ')) }
git -C $repo add -A
$mud = git -C $repo status --porcelain
if (-not $mud) { Write-Output 'Nada mudou desde a ultima versao salva.'; exit 0 }
Write-Output $mud
$msg = 'ConvaTec: versao de ' + (Get-Date -Format 'dd/MM/yyyy HH:mm')
git -C $repo -c user.name="Emilio Tannet" -c user.email="EXT.Emilio.Tannet@cevalogistics.com" commit -q -m $msg
if ($LASTEXITCODE) { throw 'commit falhou' }
Write-Output ('Commit: ' + (git -C $repo log --oneline -1))
if ($modo -eq 'enviar') { git -C $repo push -q origin HEAD; if ($LASTEXITCODE) { throw 'envio falhou' }; Write-Output 'Enviado ao GitHub.' }
else { Write-Output 'Salvo so neste computador. Para mandar ao GitHub: SALVAR VERSAO NO GIT.bat enviar' }
