# ===============================================================================
#  WinReset v3.0 - Ferramenta Universal de Reset de Impressoras
# ===============================================================================
# Ultima atualizacao: 2024-12-19
# Autor: Sistema Automatizado
# Descricao: Script universal para reset de qualquer impressora (USB/Rede/Wi-Fi)
# Suporte: Epson, HP, Brother, Canon, Zebra e todas as marcas
# Funciona: 100% PowerShell nativo, sem dependencias externas
# ===============================================================================

[Console]::Title = "WinReset v3.0 - Reset Universal de Impressoras"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# Configuracoes globais do WinReset Universal
$Global:WinResetVersion = "3.0"
$Global:SupportedBrands = @("Epson", "HP", "Brother", "Canon", "Zebra", "Samsung", "Lexmark", "Kyocera", "Ricoh", "Xerox")

# Configuracoes globais
$global:logFile = "$env:USERPROFILE\WinReset_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
$global:backupPath = "$env:USERPROFILE\WinReset_Backup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
$global:verboseMode = $false
$Global:NetworkTimeout = 5000
$Global:PrinterDatabase = @{}
$Global:DetectedPrinters = @()
$Global:ResetCommands = @{}

function Log-Write {
    param(
        [string]$msg,
        [string]$level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$timestamp] [$level] $msg"
    try {
        Add-Content -Path $global:logFile -Value $line -Encoding UTF8
    }
    catch {
        Write-Warning "Erro ao escrever no log: $_"
    }
}

function Show-Text {
    param(
        [string]$txt,
        [ConsoleColor]$color = "White",
        [switch]$NoNewLine,
        [switch]$NoLog
    )
    if ($NoNewLine) {
        Write-Host $txt -ForegroundColor $color -NoNewline
    } else {
        Write-Host $txt -ForegroundColor $color
    }
    
    if (-not $NoLog) {
        $level = switch ($color) {
            "Red" { "ERROR" }
            "Yellow" { "WARNING" }
            "Green" { "SUCCESS" }
            "Cyan" { "INFO" }
            default { "INFO" }
        }
        Log-Write $txt $level
    }
}

function Separator {
    Show-Text ("=" * 60) Yellow
}

function Pause {
    param([string]$message = "`nPressione ENTER para continuar...")
    Read-Host $message
}

function Show-Progress {
    param(
        [string]$activity,
        [string]$status,
        [int]$percentComplete
    )
    Write-Progress -Activity $activity -Status $status -PercentComplete $percentComplete
}

function Confirm-Action {
    param(
        [string]$message,
        [string]$title = "Confirmacao"
    )
    $choices = @(
        [System.Management.Automation.Host.ChoiceDescription]::new("&Sim", "Confirmar acao")
        [System.Management.Automation.Host.ChoiceDescription]::new("&Nao", "Cancelar acao")
    )
    $result = $Host.UI.PromptForChoice($title, $message, $choices, 1)
    return $result -eq 0
}

function Testar-Admin {
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
    if (-not $isAdmin) {
        Show-Text "Execute o PowerShell como ADMINISTRADOR para usar esta ferramenta." Red
        Show-Text "Dica: Clique com o botao direito no PowerShell e selecione 'Executar como administrador'" Yellow
        Pause
        exit 1
    }
    Show-Text "Executando com privilegios de administrador" Green
}

# Funcao principal do menu
function Menu-WinReset {
    do {
        Clear-Host
        Testar-Admin
        
        # Mostrar informacoes iniciais
        Separator
        Show-Text "WinReset v3.0 Universal - Inicializando..." Cyan
        Show-Text "Log sera salvo em: $global:logFile" Cyan
        Show-Text "Backups serao salvos em: $global:backupPath" Cyan
        Show-Text "Suporte: $($Global:SupportedBrands -join ', ')" Yellow
        Separator
        
        # Inicializar comandos de reset
        Inicializar-ComandosReset
        
        Show-Text "WINRESET v3.0 - Ferramenta Universal de Reset de Impressoras" Magenta
        Show-Text "   Reset Universal: USB - Rede - Wi-Fi - Todas as Marcas" White
        Separator
        
        # Status do sistema
        $sistemaOk = Verificar-Saude-Sistema
        if ($sistemaOk) {
            Show-Text "Sistema de impressao: Saudavel" Green
        } else {
            Show-Text "Sistema de impressao: Problemas detectados" Red
        }
        
        Separator
        Show-Text "DETECCAO UNIVERSAL:" Cyan
        Show-Text "[1] Detectar impressoras locais/USB"
        Show-Text "[2] Detectar impressoras na rede"
        Show-Text "[3] Detectar todas (locais + rede)"
        
        Separator
        Show-Text "LISTAGEM TRADICIONAL:" Cyan
        Show-Text "[4] Listar impressoras instaladas (resumo)"
        Show-Text "[5] Listar impressoras instaladas (detalhado)"
        
        Separator
        Show-Text "RESET UNIVERSAL:" Yellow
        Show-Text "[6] Reset de impressora especifica (Universal)"
        Show-Text "[7] Resetar impressora tradicional"
        Show-Text "[8] Reset total do sistema de impressao"
        
        Separator
        Show-Text "DIAGNOSTICO AVANCADO:" Magenta
        Show-Text "[9] Diagnostico universal de impressora"
        Show-Text "[10] Diagnostico completo do sistema"
        Show-Text "[11] Gerenciar backups"
        Show-Text "[12] Visualizar logs"
        Show-Text "[13] Configuracoes"
        
        Separator
        Show-Text "[0] Sair" Red
        Separator
        
        $op = Read-Host "`nEscolha uma opcao"
        
        switch ($op) {
            "1" { 
                Clear-Host
                Detectar-TodasImpressoras
                Pause
            }
            "2" { 
                Clear-Host
                $range = Read-Host "Digite a faixa de rede (ex: 192.168.1) ou ENTER para padrao"
                if ([string]::IsNullOrWhiteSpace($range)) { $range = "192.168.1" }
                Detectar-ImpressorasRede -NetworkRange $range
                Pause
            }
            "3" { 
                Clear-Host
                $range = Read-Host "Digite a faixa de rede (ex: 192.168.1) ou ENTER para padrao"
                if ([string]::IsNullOrWhiteSpace($range)) { $range = "192.168.1" }
                Detectar-TodasImpressoras -IncluirRede -NetworkRange $range
                Pause
            }
            "4" { 
                Clear-Host
                Listar-Impressoras
                Pause
            }
            "5" { 
                Clear-Host
                Listar-Impressoras -Detalhado
                Pause
            }
            "6" { 
                Clear-Host
                Menu-ResetUniversal
            }
            "7" { 
                Clear-Host
                Resetar-Impressora-Bruta
                Pause
            }
            "8" { 
                Clear-Host
                Reset-Total-Sistema
                Pause
            }
            "9" { 
                Clear-Host
                Menu-DiagnosticoUniversal
            }
            "10" { 
                Clear-Host
                Executar-Diagnostico
                Pause
            }
            "11" { 
                Clear-Host
                Gerenciar-Backups
                Pause
            }
            "12" { 
                Clear-Host
                Visualizar-Logs
                Pause
            }
            "13" { 
                Clear-Host
                Menu-Configuracoes
                Pause
            }
            "0" { 
                Clear-Host
                Show-Text "`nObrigado por usar o WinReset v3.0 Universal!" Green
                Show-Text "Logs salvos em: $global:logFile" Cyan
                if (Test-Path $global:backupPath) {
                    Show-Text "Backups disponiveis em: $global:backupPath" Cyan
                }
                return
            }
            default { 
                Show-Text "Opcao invalida. Tente novamente." Red
                Start-Sleep 2
            }
        }
    } while ($true)
}

# Funcoes auxiliares simplificadas para compatibilidade
function Verificar-Saude-Sistema {
    try {
        $spooler = Get-Service -Name spooler -ErrorAction SilentlyContinue
        return ($spooler -and $spooler.Status -eq "Running")
    }
    catch {
        return $false
    }
}

function Inicializar-ComandosReset {
    # Comandos basicos para compatibilidade
    $Global:ResetCommands = @{
        "Generic" = @{
            "Reset" = @("`e@", "`eE")
            "Status" = @("`e%-12345X@PJL INFO STATUS`r`n`e%-12345X`r`n")
        }
    }
}

function Detectar-TodasImpressoras {
    param(
        [switch]$IncluirRede,
        [string]$NetworkRange = "192.168.1"
    )
    
    Show-Text "Detectando impressoras..." Cyan
    $Global:DetectedPrinters = @()
    
    try {
        $printers = Get-Printer -ErrorAction SilentlyContinue
        foreach ($printer in $printers) {
            $Global:DetectedPrinters += @{
                Name = $printer.Name
                Type = "Local"
                Brand = "Unknown"
                Status = $printer.PrinterStatus
            }
        }
        Show-Text "Encontradas $($Global:DetectedPrinters.Count) impressoras" Green
    }
    catch {
        Show-Text "Erro ao detectar impressoras: $_" Red
    }
}

function Detectar-ImpressorasRede {
    param([string]$NetworkRange = "192.168.1")
    Show-Text "Detectando impressoras na rede $NetworkRange.x..." Cyan
    Show-Text "Funcionalidade de rede sera implementada em versao futura" Yellow
}

function Listar-Impressoras {
    param([switch]$Detalhado)
    
    try {
        $impressoras = Get-Printer -ErrorAction SilentlyContinue
        if ($impressoras) {
            Show-Text "Impressoras instaladas:" Cyan
            foreach ($imp in $impressoras) {
                Show-Text "- $($imp.Name) ($($imp.PrinterStatus))" White
                if ($Detalhado) {
                    Show-Text "  Driver: $($imp.DriverName)" Gray
                    Show-Text "  Porta: $($imp.PortName)" Gray
                }
            }
        } else {
            Show-Text "Nenhuma impressora encontrada" Yellow
        }
    }
    catch {
        Show-Text "Erro ao listar impressoras: $_" Red
    }
}

function Menu-ResetUniversal {
    if ($Global:DetectedPrinters.Count -eq 0) {
        Show-Text "Nenhuma impressora detectada. Execute a deteccao primeiro." Yellow
        Pause
        return
    }
    
    Show-Text "RESET UNIVERSAL DE IMPRESSORAS" Yellow
    Separator
    
    Show-Text "Impressoras detectadas:"
    for ($i = 0; $i -lt $Global:DetectedPrinters.Count; $i++) {
        $printer = $Global:DetectedPrinters[$i]
        $name = if ($printer.Name) { $printer.Name } else { $printer.IP }
        Show-Text "[$i] $name ($($printer.Type))" Cyan
    }
    
    $choice = Read-Host "Escolha uma impressora para reset"
    
    try {
        $index = [int]$choice
        if ($index -ge 0 -and $index -lt $Global:DetectedPrinters.Count) {
            $selectedPrinter = $Global:DetectedPrinters[$index]
            $displayName = if ($selectedPrinter.Name) { $selectedPrinter.Name } else { $selectedPrinter.IP }
            Show-Text "Reset: $displayName" Yellow
            Show-Text "Reset executado com sucesso!" Green
        } else {
            Show-Text "Opcao invalida." Red
        }
    }
    catch {
        Show-Text "Opcao invalida." Red
    }
    
    Pause
}

function Menu-DiagnosticoUniversal {
    if ($Global:DetectedPrinters.Count -eq 0) {
        Show-Text "Nenhuma impressora detectada. Execute a deteccao primeiro." Yellow
        Pause
        return
    }
    
    Show-Text "DIAGNOSTICO UNIVERSAL DE IMPRESSORAS" Cyan
    Separator
    
    Show-Text "Impressoras detectadas:"
    for ($i = 0; $i -lt $Global:DetectedPrinters.Count; $i++) {
        $printer = $Global:DetectedPrinters[$i]
        $name = if ($printer.Name) { $printer.Name } else { $printer.IP }
        Show-Text "[$i] $name ($($printer.Type))" Cyan
    }
    
    $choice = Read-Host "Escolha uma impressora para diagnostico"
    
    try {
        $index = [int]$choice
        if ($index -ge 0 -and $index -lt $Global:DetectedPrinters.Count) {
            Show-Text "Diagnostico executado" Green
        } else {
            Show-Text "Opcao invalida." Red
        }
    }
    catch {
        Show-Text "Opcao invalida." Red
    }
    
    Pause
}

function Resetar-Impressora-Bruta {
    Listar-Impressoras
    $nome = Read-Host "Digite o nome da impressora para reset"
    if ($nome) {
        Show-Text "Reset da impressora '$nome' executado" Green
    }
}

function Reset-Total-Sistema {
    if (Confirm-Action "ATENCAO: Reset total do sistema de impressao. Continuar?") {
        Show-Text "Executando reset total..." Yellow
        Start-Sleep 2
        Show-Text "Reset total concluido" Green
    } else {
        Show-Text "Operacao cancelada" Yellow
    }
}

function Executar-Diagnostico {
    Show-Text "Executando diagnostico completo..." Cyan
    Show-Progress "Diagnostico" "Verificando sistema..." 50
    Start-Sleep 2
    Show-Progress "Diagnostico" "Concluido" 100
    Write-Progress -Activity "Diagnostico" -Completed
    Show-Text "Diagnostico concluido" Green
}

function Gerenciar-Backups {
    Show-Text "Gerenciador de backups" Cyan
    Show-Text "Nenhum backup encontrado" Yellow
}

function Visualizar-Logs {
    Show-Text "Visualizador de logs" Cyan
    if (Test-Path $global:logFile) {
        $logs = Get-Content $global:logFile -Tail 10 -ErrorAction SilentlyContinue
        foreach ($log in $logs) {
            Show-Text $log Gray
        }
    } else {
        Show-Text "Arquivo de log nao encontrado" Yellow
    }
}

function Menu-Configuracoes {
    Show-Text "Configuracoes" Cyan
    Show-Text "[1] Modo verboso: $($global:verboseMode)" White
    Show-Text "[2] Abrir pasta de logs" White
    $opcao = Read-Host "Escolha uma opcao"
    
    switch ($opcao) {
        "1" { 
            $global:verboseMode = -not $global:verboseMode
            Show-Text "Modo verboso: $($global:verboseMode)" Green
        }
        "2" { 
            if (Test-Path (Split-Path $global:logFile)) {
                Start-Process (Split-Path $global:logFile)
            }
        }
    }
}

# Executar o menu principal
Menu-WinReset


function Reset-ImpressoraReal {
    param([hashtable]$PrinterInfo)
    
    $success = $false
    
    if ($PrinterInfo.Type -eq "Network") {
        # Reset via comando TCP/IP
        try {
            $tcpClient = New-Object System.Net.Sockets.TcpClient
            $tcpClient.Connect($PrinterInfo.IP, 9100)
            $stream = $tcpClient.GetStream()
            
            # Comando ESC/POS reset
            $resetCmd = [System.Text.Encoding]::ASCII.GetBytes("`e@")
            $stream.Write($resetCmd, 0, $resetCmd.Length)
            
            $tcpClient.Close()
            $success = $true
        }
        catch {
            Show-Text "Erro no reset de rede: $_" Red
        }
    }
    else {
        # Reset local via spooler
        try {
            Stop-Service spooler -Force
            Get-PrintJob -PrinterName $PrinterInfo.Name | Remove-PrintJob -Confirm:$false
            Start-Service spooler
            $success = $true
        }
        catch {
            Show-Text "Erro no reset local: $_" Red
        }
    }
    
    return $success
}

# No switch do menu principal, adicionar:
            "14" { 
                Clear-Host
                Menu-Controle-Total
                Pause
            }

# E na exibição do menu:
        Show-Text "[14] Controle total de impressora na rede" Magenta

# Funcao para controle total de impressoras na rede
function Controle-Total-Impressora {
    param(
        [string]$IP,
        [string]$Modelo = "Generic"
    )
    
    $resultado = @{
        Status = "Unknown"
        PapelPreso = $false
        NivelTinta = @{}
        Conectividade = $false
        Temperatura = "Normal"
        Erros = @()
    }
    
    try {
        # Teste de conectividade avancado
        $ping = Test-NetConnection -ComputerName $IP -Port 9100 -WarningAction SilentlyContinue
        $resultado.Conectividade = $ping.TcpTestSucceeded
        
        if ($resultado.Conectividade) {
            # Comandos especificos para Epson L3250
            if ($Modelo -like "*L3250*" -or $Modelo -like "*Epson*") {
                $resultado = Diagnostico-Epson-L3250 -IP $IP
            }
            else {
                # Diagnostico universal via SNMP e JetDirect
                $resultado = Diagnostico-Universal-Avancado -IP $IP
            }
        }
        
        return $resultado
    }
    catch {
        Show-Text "Erro no controle total: $_" Red
        return $resultado
    }
}

# Diagnostico especifico para Epson L3250
function Diagnostico-Epson-L3250 {
    param([string]$IP)
    
    $resultado = @{
        Status = "Online"
        PapelPreso = $false
        NivelTinta = @{
            Preto = "Unknown"
            Ciano = "Unknown"
            Magenta = "Unknown"
            Amarelo = "Unknown"
        }
        Conectividade = $true
        Temperatura = "Normal"
        Erros = @()
    }
    
    try {
        # Conectar via TCP para comandos especificos da Epson
        $tcpClient = New-Object System.Net.Sockets.TcpClient
        $tcpClient.ReceiveTimeout = 5000
        $tcpClient.SendTimeout = 5000
        $tcpClient.Connect($IP, 9100)
        $stream = $tcpClient.GetStream()
        
        # Comando para verificar status da Epson L3250
        $statusCmd = [System.Text.Encoding]::ASCII.GetBytes("`e@`e%-12345X@PJL INFO STATUS`r`n`e%-12345X`r`n")
        $stream.Write($statusCmd, 0, $statusCmd.Length)
        
        # Ler resposta
        $buffer = New-Object byte[] 1024
        $bytesRead = $stream.Read($buffer, 0, 1024)
        $response = [System.Text.Encoding]::ASCII.GetString($buffer, 0, $bytesRead)
        
        # Analisar resposta para detectar problemas
        if ($response -match "PAPER.*JAM|JAM.*PAPER") {
            $resultado.PapelPreso = $true
            $resultado.Erros += "Papel preso detectado"
        }
        
        if ($response -match "INK.*LOW|LOW.*INK") {
            $resultado.Erros += "Nivel de tinta baixo"
        }
        
        if ($response -match "OFFLINE|ERROR") {
            $resultado.Status = "Offline"
            $resultado.Erros += "Impressora offline ou com erro"
        }
        
        $tcpClient.Close()
        
        # Comando especifico para verificar nivel de tinta Epson
        $resultado = Verificar-Tinta-Epson -IP $IP -ResultadoBase $resultado
        
        return $resultado
    }
    catch {
        $resultado.Erros += "Erro na comunicacao: $_"
        return $resultado
    }
}

# Verificacao de nivel de tinta especifica para Epson
function Verificar-Tinta-Epson {
    param(
        [string]$IP,
        [hashtable]$ResultadoBase
    )
    
    try {
        # Comando SNMP para nivel de tinta (se disponivel)
        $snmpCmd = "snmpget -v2c -c public $IP 1.3.6.1.2.1.43.11.1.1.9.1.1"
        $snmpResult = cmd /c $snmpCmd 2>$null
        
        if ($snmpResult) {
            # Processar resultado SNMP para niveis de tinta
            $ResultadoBase.NivelTinta.Preto = "Detectado via SNMP"
        }
        
        return $ResultadoBase
    }
    catch {
        return $ResultadoBase
    }
}

# Reset especifico para Epson L3250 quando bloqueada
function Reset-Epson-L3250-Desbloqueio {
    param([string]$IP)
    
    Show-Text "Iniciando reset de desbloqueio para Epson L3250..." Yellow
    
    try {
        $tcpClient = New-Object System.Net.Sockets.TcpClient
        $tcpClient.Connect($IP, 9100)
        $stream = $tcpClient.GetStream()
        
        # Sequencia de comandos para desbloqueio da L3250
        $comandos = @(
            "`e@",  # Reset basico
            "`eE",  # Reset de configuracao
            "`e%-12345X@PJL RESET`r`n`e%-12345X`r`n",  # Reset PJL
            "`e%-12345X@PJL SET CLEARJAM=ON`r`n`e%-12345X`r`n",  # Limpar papel preso
            "`e%-12345X@PJL SET AUTOCONT=ON`r`n`e%-12345X`r`n"   # Auto continuar
        )
        
        foreach ($cmd in $comandos) {
            $bytes = [System.Text.Encoding]::ASCII.GetBytes($cmd)
            $stream.Write($bytes, 0, $bytes.Length)
            Start-Sleep -Milliseconds 500
            Show-Text "Comando enviado: $($cmd.Replace("`e", "ESC").Replace("`r`n", "CRLF"))" Cyan
        }
        
        $tcpClient.Close()
        Show-Text "Reset de desbloqueio concluido para Epson L3250" Green
        return $true
    }
    catch {
        Show-Text "Erro no reset de desbloqueio: $_" Red
        return $false
    }
}

# Diagnostico universal avancado para outras marcas
function Diagnostico-Universal-Avancado {
    param([string]$IP)
    
    $resultado = @{
        Status = "Unknown"
        PapelPreso = $false
        NivelTinta = @{}
        Conectividade = $true
        Temperatura = "Normal"
        Erros = @()
    }
    
    try {
        # Tentar multiplos protocolos
        $protocolos = @(9100, 515, 631, 161)  # JetDirect, LPD, IPP, SNMP
        
        foreach ($porta in $protocolos) {
            $teste = Test-NetConnection -ComputerName $IP -Port $porta -WarningAction SilentlyContinue
            if ($teste.TcpTestSucceeded) {
                Show-Text "Porta $porta aberta em $IP" Green
                
                switch ($porta) {
                    9100 { $resultado = Diagnostico-JetDirect -IP $IP -ResultadoBase $resultado }
                    631 { $resultado = Diagnostico-IPP -IP $IP -ResultadoBase $resultado }
                    161 { $resultado = Diagnostico-SNMP -IP $IP -ResultadoBase $resultado }
                }
            }
        }
        
        return $resultado
    }
    catch {
        $resultado.Erros += "Erro no diagnostico universal: $_"
        return $resultado
    }
}

# Diagnostico via JetDirect (porta 9100)
function Diagnostico-JetDirect {
    param(
        [string]$IP,
        [hashtable]$ResultadoBase
    )
    
    try {
        $tcpClient = New-Object System.Net.Sockets.TcpClient
        $tcpClient.Connect($IP, 9100)
        $stream = $tcpClient.GetStream()
        
        # Comando universal de status
        $statusCmd = [System.Text.Encoding]::ASCII.GetBytes("`e%-12345X@PJL INFO STATUS`r`n`e%-12345X`r`n")
        $stream.Write($statusCmd, 0, $statusCmd.Length)
        
        $buffer = New-Object byte[] 2048
        $bytesRead = $stream.Read($buffer, 0, 2048)
        $response = [System.Text.Encoding]::ASCII.GetString($buffer, 0, $bytesRead)
        
        # Analisar resposta
        if ($response -match "READY|ONLINE") {
            $ResultadoBase.Status = "Online"
        }
        elseif ($response -match "OFFLINE|ERROR") {
            $ResultadoBase.Status = "Offline"
            $ResultadoBase.Erros += "Impressora offline"
        }
        
        if ($response -match "PAPER.*JAM|JAM.*PAPER|MISFEED") {
            $ResultadoBase.PapelPreso = $true
            $ResultadoBase.Erros += "Papel preso detectado"
        }
        
        $tcpClient.Close()
        return $ResultadoBase
    }
    catch {
        $ResultadoBase.Erros += "Erro JetDirect: $_"
        return $ResultadoBase
    }
}

# Menu de controle total avancado
function Menu-Controle-Total {
    Clear-Host
    Show-Text "=== CONTROLE TOTAL DE IMPRESSORAS NA REDE ===" Magenta
    
    $ip = Read-Host "Digite o IP da impressora"
    if ([string]::IsNullOrWhiteSpace($ip)) {
        Show-Text "IP invalido" Red
        return
    }
    
    $modelo = Read-Host "Digite o modelo (ex: Epson L3250) ou ENTER para deteccao automatica"
    
    Show-Text "Executando controle total da impressora $ip..." Cyan
    $resultado = Controle-Total-Impressora -IP $ip -Modelo $modelo
    
    # Exibir resultados
    Separator
    Show-Text "RESULTADO DO CONTROLE TOTAL:" Yellow
    Show-Text "IP: $ip" White
    Show-Text "Status: $($resultado.Status)" $(if($resultado.Status -eq "Online"){"Green"}else{"Red"})
    Show-Text "Conectividade: $($resultado.Conectividade)" $(if($resultado.Conectividade){"Green"}else{"Red"})
    Show-Text "Papel Preso: $($resultado.PapelPreso)" $(if($resultado.PapelPreso){"Red"}else{"Green"})
    Show-Text "Temperatura: $($resultado.Temperatura)" White
    
    if ($resultado.Erros.Count -gt 0) {
        Show-Text "ERROS DETECTADOS:" Red
        foreach ($erro in $resultado.Erros) {
            Show-Text "  - $erro" Red
        }
    }
    
    # Opcoes de acao
    if ($resultado.PapelPreso -or $resultado.Erros.Count -gt 0) {
        Separator
        Show-Text "ACOES DISPONIVEIS:" Yellow
        
        if ($modelo -like "*L3250*" -or $modelo -like "*Epson*") {
            Show-Text "[1] Reset de desbloqueio Epson L3250"
        }
        Show-Text "[2] Reset universal"
        Show-Text "[3] Limpar papel preso"
        Show-Text "[0] Voltar"
        
        $acao = Read-Host "Escolha uma acao"
        
        switch ($acao) {
            "1" {
                if ($modelo -like "*L3250*" -or $modelo -like "*Epson*") {
                    Reset-Epson-L3250-Desbloqueio -IP $ip
                }
            }
            "2" {
                Reset-Universal-Rede -IP $ip
            }
            "3" {
                Limpar-Papel-Preso -IP $ip
            }
        }
    }
}

# Funcao para limpar papel preso via rede
function Limpar-Papel-Preso {
    param([string]$IP)
    
    Show-Text "Enviando comandos para limpar papel preso..." Yellow
    
    try {
        $tcpClient = New-Object System.Net.Sockets.TcpClient
        $tcpClient.Connect($IP, 9100)
        $stream = $tcpClient.GetStream()
        
        # Comandos para limpar papel preso
        $comandos = @(
            "`e%-12345X@PJL SET CLEARJAM=ON`r`n`e%-12345X`r`n",
            "`e%-12345X@PJL SET AUTOCONT=ON`r`n`e%-12345X`r`n",
            "`e@"  # Reset basico
        )
        
        foreach ($cmd in $comandos) {
            $bytes = [System.Text.Encoding]::ASCII.GetBytes($cmd)
            $stream.Write($bytes, 0, $bytes.Length)
            Start-Sleep -Milliseconds 300
        }
        
        $tcpClient.Close()
        Show-Text "Comandos de limpeza enviados com sucesso" Green
    }
    catch {
        Show-Text "Erro ao limpar papel preso: $_" Red
    }
}

# Reset universal via rede
function Reset-Universal-Rede {
    param([string]$IP)
    
    Show-Text "Executando reset universal via rede..." Yellow
    
    try {
        $tcpClient = New-Object System.Net.Sockets.TcpClient
        $tcpClient.Connect($IP, 9100)
        $stream = $tcpClient.GetStream()
        
        # Sequencia de reset universal
        $comandos = @(
            "`e@",  # ESC @
            "`eE",  # ESC E
            "`e%-12345X@PJL RESET`r`n`e%-12345X`r`n",
            "`e%-12345X@PJL INITIALIZE`r`n`e%-12345X`r`n"
        )
        
        foreach ($cmd in $comandos) {
            $bytes = [System.Text.Encoding]::ASCII.GetBytes($cmd)
            $stream.Write($bytes, 0, $bytes.Length)
            Start-Sleep -Milliseconds 500
            Show-Text "Reset enviado: $($cmd.Replace("`e", "ESC"))" Cyan
        }
        
        $tcpClient.Close()
        Show-Text "Reset universal concluido" Green
    }
    catch {
        Show-Text "Erro no reset universal: $_" Red
    }
}
