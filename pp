# WinReset.ps1 - Reset Total e Forçado de Qualquer Impressora - Versão Aprimorada
# Criado por João Victor

[Console]::Title = "🖨️ WinReset - Reset Total e Forçado de Impressoras"

# Caminho para arquivo de log
$logFile = "$env:USERPROFILE\WinReset_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"

function Log-Write {
    param([string]$msg)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$timestamp] $msg"
    Add-Content -Path $logFile -Value $line
}

function Show-Text {
    param(
        [string]$txt,
        [ConsoleColor]$color = "White"
    )
    Write-Host $txt -ForegroundColor $color
    Log-Write $txt
}

function Separator {
    Show-Text ("=" * 60) Yellow
}

function Pause {
    Read-Host "`n⏸ Pressione ENTER para continuar..."
}

function Testar-Admin {
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
    if (-not $isAdmin) {
        Show-Text "❌ Execute o PowerShell como ADMINISTRADOR para usar esta ferramenta." Red
        Pause
        exit
    }
}

function Verificar-ServicoSpooler {
    try {
        $status = (Get-Service -Name spooler).Status
        Show-Text "🛠️ Serviço Spooler está atualmente: $status" Cyan
        Log-Write "Status do serviço spooler: $status"
    }
    catch {
        Show-Text "❌ Não foi possível obter status do serviço spooler: $_" Red
        Log-Write "Erro ao obter status do serviço spooler: $_"
    }
}

function Listar-Impressoras {
    Clear-Host
    Separator
    Show-Text "📃 Impressoras instaladas no sistema:" Cyan
    Separator

    try {
        $global:impressoras = Get-Printer | Select-Object Name, DriverName, PortName
    }
    catch {
        Show-Text "❌ Erro ao listar impressoras: $_" Red
        Log-Write "Erro ao listar impressoras: $_"
        Pause
        return $false
    }

    if (!$impressoras -or $impressoras.Count -eq 0) {
        Show-Text "❌ Nenhuma impressora foi encontrada." Red
        Pause
        return $false
    }

    for ($i=0; $i -lt $impressoras.Count; $i++) {
        $imp = $impressoras[$i]
        Show-Text "[$i] $($imp.Name)  |  Driver: $($imp.DriverName)  |  Porta: $($imp.PortName)"
    }

    return $true
}

function Limpar-FilasImpressora {
    param(
        [string]$printerName
    )
    try {
        Show-Text "⏳ Limpando filas da impressora '$printerName'..." Yellow
        $jobs = Get-CimInstance -ClassName Win32_PrintJob | Where-Object { $_.Name -like "$printerName,*" }
        foreach ($job in $jobs) {
            $job | Invoke-CimMethod -MethodName Delete | Out-Null
        }
        Show-Text "✅ Filas da impressora '$printerName' limpas." Green
    }
    catch {
        Show-Text "❌ Erro ao limpar filas da impressora: $_" Red
    }
}

function Limpar-FilesSpooler {
    Show-Text "⏳ Limpando arquivos da fila do spooler..." Yellow
    $spoolPath = "C:\Windows\System32\spool\PRINTERS\*"
    if (Test-Path $spoolPath) {
        try {
            Remove-Item $spoolPath -Force -Recurse -ErrorAction Stop
            Show-Text "✅ Arquivos de spool limpos." Green
        }
        catch {
            Show-Text "❌ Erro ao limpar arquivos de spooler: $_" Red
        }
    }
    else {
        Show-Text "⚠️ Pasta de spooler não encontrada." Yellow
    }
}

function Resetar-Impressora-Bruta {
    if (-not (Listar-Impressoras)) { return }

    $index = Read-Host "`nDigite o número da impressora que deseja resetar"
    if ($index -notmatch '^\d+$' -or [int]$index -ge $impressoras.Count) {
        Show-Text "❌ Índice inválido. Tente novamente." Red
        Pause
        return
    }

    $nome = $impressoras[$index].Name
    Show-Text "`n🔄 Iniciando reset brutal da impressora '$nome'..." Cyan

    try {
        Show-Text "⏳ Parando serviço spooler..." Yellow
        Stop-Service spooler -Force -ErrorAction Stop

        Limpar-FilasImpressora -printerName $nome
        Limpar-FilesSpooler

        # Opcional: remover driver da impressora
        $driverName = $impressoras[$index].DriverName
        if ($driverName) {
            try {
                Show-Text "⏳ Removendo driver '$driverName'..." Yellow
                Remove-PrinterDriver -Name $driverName -ErrorAction SilentlyContinue
                Show-Text "✅ Driver removido." Green
            }
            catch {
                Show-Text "⚠️ Não foi possível remover o driver: $_" Yellow
            }
        }

        Show-Text "⏳ Iniciando serviço spooler..." Yellow
        Start-Service spooler -ErrorAction Stop

        Show-Text "✅ Impressora '$nome' resetada com sucesso!" Green
        Verificar-ServicoSpooler
    }
    catch {
        Show-Text "❌ Erro no reset brutal: $_" Red
        Log-Write "Erro no reset brutal: $_"
    }

    Pause
}

function Resetar-Tudo-Bruto {
    Clear-Host
    Separator
    Show-Text "⚠️ ATENÇÃO: Você irá resetar TODOS os serviços, filas e drivers de impressão!" Yellow
    $confirm = Read-Host "Deseja continuar? (S/N)"
    if ($confirm.ToUpper() -ne 'S') {
        Show-Text "Operação cancelada pelo usuário." Red
        Pause
        return
    }

    Show-Text "`n♻ Resetando todos os serviços e filas de impressão (bruto)..." Cyan

    try {
        Show-Text "⏳ Parando serviço spooler..." Yellow
        Stop-Service spooler -Force -ErrorAction Stop

        Show-Text "⏳ Limpando todas as filas..." Yellow
        $jobs = Get-CimInstance -ClassName Win32_PrintJob
        foreach ($job in $jobs) {
            $job | Invoke-CimMethod -MethodName Delete | Out-Null
        }
        Show-Text "✅ Todas as filas de impressão foram limpas." Green

        Limpar-FilesSpooler

        # Opcional: remover todos os drivers (descomente se quiser)
        # $drivers = Get-PrinterDriver
        # foreach ($drv in $drivers) {
        #     Remove-PrinterDriver -Name $drv.Name -ErrorAction SilentlyContinue
        # }
        # Show-Text "✅ Todos os drivers de impressoras removidos." Green

        Show-Text "⏳ Iniciando serviço spooler..." Yellow
        Start-Service spooler -ErrorAction Stop

        Show-Text "✅ Serviço spooler reiniciado com sucesso." Green
        Verificar-ServicoSpooler
    }
    catch {
        Show-Text "❌ Falha ao reiniciar os serviços de impressão: $_" Red
        Log-Write "Falha ao reiniciar serviços: $_"
    }

    Pause
}

function Menu-WinReset {
    Testar-Admin

    do {
        Clear-Host
        Separator
        Show-Text "🖨️ WINRESET BRUTO - Ferramenta de Reset Completo de Impressoras" Magenta
        Separator

        Verificar-ServicoSpooler

        Show-Text "[1] 📋 Listar impressoras instaladas"
        Show-Text "[2] 🔁 Resetar impressora específica (bruto)"
        Show-Text "[3] ♻ Resetar todos os serviços e filas de impressão (bruto)"
        Show-Text "[0] ❌ Sair"
        Separator

        $op = Read-Host "`nEscolha uma opção"
        switch ($op) {
            '1' { 
                Listar-Impressoras
                Pause
            }
            '2' { Resetar-Impressora-Bruta }
            '3' { Resetar-Tudo-Bruto }
            '0' { break }
            default { 
                Show-Text "❌ Opção inválida. Tente novamente." Red
                Pause
            }
        }
    } while ($true)
}

# Executa o menu
Menu-WinReset
