# Script para corrigir erros no winreset.ps1
$scriptPath = "c:\Users\jao\Desktop\apps\winreset.ps1"

if (Test-Path $scriptPath) {
    # Fazer backup
    $backupPath = "$scriptPath.backup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    Copy-Item $scriptPath $backupPath
    Write-Host "Backup criado: $backupPath" -ForegroundColor Green
    
    # Ler conteúdo
    $content = Get-Content $scriptPath -Raw
    
    # Correção 1: Escapar caracteres & nas ChoiceDescription
    $content = $content -replace '"&Sim"', '"S&im"'
    $content = $content -replace '"&Nao"', '"N&ao"'
    
    # Correção 2: Corrigir regex problemática
    $content = $content -replace '\(\$padrao\[\\w\\\-\]\+\)', '"($padrao" + "[\w\-]+)"'
    
    # Salvar arquivo corrigido
    $content | Set-Content $scriptPath -Encoding UTF8
    
    Write-Host "Correções aplicadas com sucesso!" -ForegroundColor Green
    Write-Host "Execute novamente: iex (irm 'sua-url-do-script')" -ForegroundColor Cyan
}
else {
    Write-Host "Arquivo não encontrado: $scriptPath" -ForegroundColor Red
}
