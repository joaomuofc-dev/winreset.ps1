# WinReset.ps1 - Reset Total de Impressoras e Spooler
# Criado por João Victor - Ferramenta Real e Expansível

[Console]::Title = "🖨️ WinReset - Ferramenta de Reset de Impressoras"

function Show-Text { param($txt, $color="White"); Write-Host $txt -ForegroundColor $color }
function Separator { Show-Text ("=" * 60) Yellow }
function Pause { Read-Host "`n⏸ Pressione ENTER para continuar..." }

function Listar-Impressoras {
    Show-Text "`n📃 Impressoras instaladas no sistema:" Cyan
    $global:impressoras = Get-Printer | Select-Object Name, DriverName, PortName
    if (!$impressoras -or $impressoras.Count -eq 0) {
        Show-Text "❌ Nenhuma impressora foi encontrada." Red
        Pause
        return
    }

    $i = 0
    foreach ($imp in $impressoras) {
        Write-Host "[$i] $($imp.Name)  |  Driver: $($imp.DriverName)  |  Porta: $($imp.PortName)"
        $i++
    }
}

function Resetar-Impressora {
    Listar-Impressoras
    $index = Read-Host "`nDigite o número da impressora que deseja resetar"
    if ($index -notmatch '^\d+$' -or $index -ge $impressoras.Count) {
        Show-Text "❌ Índice inválido. Tente novamente." Red; Pause; return
    }

    $nome = $impressoras[$index].Name
    Show-Text "`n🔄 Resetando impressora '$nome'..." Cyan

    try {
        Stop-Service spooler -Force
        Remove-Item "C:\Windows\System32\spool\PRINTERS\*" -Force -ErrorAction SilentlyContinue
        Start-Service spooler
        Show-Text "✅ Spooler reiniciado e fila da impressora limpa." Green
    } catch {
        Show-Text "❌ Erro ao limpar spooler: $_" Red
    }

    Pause
}

function Resetar-Tudo {
    Show-Text "`n♻ Resetando todos os serviços e filas de impressão..." Cyan
    try {
        Stop-Service spooler -Force
        Remove-Item "C:\Windows\System32\spool\PRINTERS\*" -Force -ErrorAction SilentlyContinue
        Start-Service spooler
        Show-Text "✅ Todos os serviços de impressão foram resetados." Green
    } catch {
        Show-Text "❌ Falha ao reiniciar os serviços de impressão." Red
    }
    Pause
}

function Menu-WinReset {
    do {
        Clear-Host
        Separator
        Show-Text "🖨️ WINRESET - Ferramenta de Reset de Impressoras 🧼" Magenta
        Separator
        Show-Text "[1] 📋 Listar impressoras instaladas"
        Show-Text "[2] 🔁 Resetar uma impressora específica"
        Show-Text "[3] ♻ Resetar todos os serviços de impressão"
        Show-Text "[0] ❌ Sair"
        Separator
        $op = Read-Host "`nEscolha uma opção"
        switch ($op) {
            '1' { Listar-Impressoras; Pause }
            '2' { Resetar-Impressora }
            '3' { Resetar-Tudo }
            '0' { break }
            default { Show-Text "❌ Opção inválida. Tente novamente." Red; Pause }
        }
    } while ($true)
}

Menu-WinReset
