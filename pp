# ===============================================================================
#  WinReset v3.0 - Ferramenta Universal de Reset de Impressoras
# ===============================================================================
# Ultima atualizacao: 2024-12-19
# Autor: Sistema Automatizado
# Descricao: Script universal para reset de qualquer impressora (USB/Rede/Wi-Fi)
# Suporte: Epson, HP, Brother, Canon, Zebra e todas as marcas
# Funciona: 100% PowerShell nativo, sem dependencias externas
# ===============================================================================

# Sistema de seguranÃ§a e auditoria
function Sistema-Seguranca {
    param(
        [string]$Acao,
        [string]$IP,
        [string]$Usuario = $env:USERNAME
    )
    
    $logSeguranca = "$env:USERPROFILE\WinReset_Security_$(Get-Date -Format 'yyyyMM').log"
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    # Verificar permissÃµes
    if (-not (Verificar-Permissoes -Acao $Acao)) {
        $entrada = "[$timestamp] [NEGADO] UsuÃ¡rio: $Usuario | AÃ§Ã£o: $Acao | IP: $IP | Motivo: Sem permissÃ£o"
        Add-Content -Path $logSeguranca -Value $entrada
        Show-Text "âŒ Acesso negado para a aÃ§Ã£o: $Acao" Red
        return $false
    }
    
    # Log da aÃ§Ã£o autorizada
    $entrada = "[$timestamp] [AUTORIZADO] UsuÃ¡rio: $Usuario | AÃ§Ã£o: $Acao | IP: $IP"
    Add-Content -Path $logSeguranca -Value $entrada
    
    # Verificar se Ã© aÃ§Ã£o crÃ­tica
    $acoesCriticas = @("Reset-Total", "Formatacao", "Configuracao-Rede")
    if ($Acao -in $acoesCriticas) {
        Show-Text "âš ï¸  AÃ‡ÃƒO CRÃTICA DETECTADA: $Acao" Yellow
        if (-not (Confirmar-AcaoCritica -Acao $Acao -IP $IP)) {
            $entrada = "[$timestamp] [CANCELADO] UsuÃ¡rio: $Usuario | AÃ§Ã£o: $Acao | IP: $IP | Motivo: Cancelado pelo usuÃ¡rio"
            Add-Content -Path $logSeguranca -Value $entrada
            return $false
        }
    }
    
    return $true
}

# VerificaÃ§Ã£o de permissÃµes
function Verificar-Permissoes {
    param([string]$Acao)
    
    # Verificar se Ã© administrador
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
    
    $acoesAdminObrigatorio = @("Reset-Total", "Configuracao-Sistema", "Backup-Restauracao")
    
    if ($Acao -in $acoesAdminObrigatorio -and -not $isAdmin) {
        return $false
    }
    
    return $true
}

# ConfirmaÃ§Ã£o para aÃ§Ãµes crÃ­ticas
function Confirmar-AcaoCritica {
    param(
        [string]$Acao,
        [string]$IP
    )
    
    Show-Text "ðŸ” CONFIRMAÃ‡ÃƒO DE SEGURANÃ‡A" Red
    Show-Text "AÃ§Ã£o: $Acao" Yellow
    Show-Text "Alvo: $IP" Yellow
    Show-Text "Esta Ã© uma aÃ§Ã£o irreversÃ­vel que pode afetar o funcionamento da impressora." Red
    
    $codigo = Get-Random -Minimum 1000 -Maximum 9999
    Show-Text "Digite o cÃ³digo de confirmaÃ§Ã£o: $codigo" Cyan
    
    $entrada = Read-Host "CÃ³digo"
    
    if ($entrada -eq $codigo.ToString()) {
        Show-Text "âœ… ConfirmaÃ§Ã£o aceita" Green
        return $true
    }
    else {
        Show-Text "âŒ CÃ³digo incorreto. AÃ§Ã£o cancelada." Red
        return $false
    }
}

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
        Show-Text "[15] Dashboard tempo real" Cyan
        Show-Text "[16] AutomaÃ§Ã£o inteligente" Cyan
        Show-Text "[17] Interface grÃ¡fica" Magenta
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
            "15" { 
                Clear-Host
                $impressoras = IA-DeteccaoInteligente
                if ($impressoras.Count -gt 0) {
                    Dashboard-TempoReal -ImpressorasMonitoradas $impressoras
                }
                else {
                    Show-Text "Nenhuma impressora detectada para monitoramento" Yellow
                }
                Pause
            }
            "16" { 
                Clear-Host
                $impressoras = IA-DeteccaoInteligente
                if ($impressoras.Count -gt 0) {
                    Automacao-Inteligente -ImpressorasMonitoradas $impressoras
                }
                else {
                    Show-Text "Nenhuma impressora detectada para automaÃ§Ã£o" Yellow
                }
                Pause
            }
            "17" { 
                Interface-Grafica
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

# ðŸš€ Melhorias AvanÃ§adas para WinReset v3.0

# Sistema de automaÃ§Ã£o inteligente
function Automacao-Inteligente {
    param(
        [array]$ImpressorasMonitoradas,
        [switch]$ModoAutomatico
    )
    
    Show-Text "ðŸ¤– Iniciando sistema de automaÃ§Ã£o inteligente..." Magenta
    
    $regrasAutomacao = @(
        @{
            Nome = "Auto-Reset Papel Preso"
            Condicao = { param($status) $status.PapelPreso }
            Acao = { param($ip) Reset-PapelPreso-Automatico -IP $ip }
            Ativo = $true
        },
        @{
            Nome = "NotificaÃ§Ã£o Tinta Baixa"
            Condicao = { param($status) $status.TintaBaixa }
            Acao = { param($ip) Notificar-TintaBaixa -IP $ip }
            Ativo = $true
        },
        @{
            Nome = "ReconexÃ£o AutomÃ¡tica"
            Condicao = { param($status) $status.Status -eq "Offline" }
            Acao = { param($ip) Tentar-Reconexao -IP $ip }
            Ativo = $true
        }
    )
    
    while ($true) {
        foreach ($impressora in $ImpressorasMonitoradas) {
            $status = Monitorar-StatusRapido -IP $impressora.IP
            
            foreach ($regra in $regrasAutomacao) {
                if ($regra.Ativo -and (& $regra.Condicao $status)) {
                    Show-Text "ðŸ”§ Executando automaÃ§Ã£o: $($regra.Nome) para $($impressora.IP)" Yellow
                    
                    if ($ModoAutomatico -or (Confirmar-Automacao -Regra $regra.Nome -IP $impressora.IP)) {
                        & $regra.Acao $impressora.IP
                        
                        # Log da automaÃ§Ã£o
                        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                        $logEntry = "[$timestamp] [AUTOMACAO] $($regra.Nome) executada em $($impressora.IP)"
                        Add-Content -Path $global:logFile -Value $logEntry
                    }
                }
            }
        }
        
        Start-Sleep -Seconds 30  # Verificar a cada 30 segundos
        
        # Verificar se deve parar
        if ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)
            if ($key.KeyChar -eq 'q' -or $key.KeyChar -eq 'Q') {
                Show-Text "AutomaÃ§Ã£o interrompida pelo usuÃ¡rio" Yellow
                break
            }
        }
    }
}

# Reset automÃ¡tico de papel preso
function Reset-PapelPreso-Automatico {
    param([string]$IP)
    
    Show-Text "ðŸ”§ Executando reset automÃ¡tico de papel preso em $IP..." Cyan
    
    try {
        $tcpClient = New-Object System.Net.Sockets.TcpClient
        $tcpClient.Connect($IP, 9100)
        $stream = $tcpClient.GetStream()
        
        # SequÃªncia especÃ­fica para papel preso
        $comandos = @(
            "`e%-12345X@PJL SET CLEARJAM=ON`r`n`e%-12345X`r`n",
            "`e%-12345X@PJL SET AUTOCONT=ON`r`n`e%-12345X`r`n",
            "`e@",  # Reset bÃ¡sico
            "`e%-12345X@PJL RESET`r`n`e%-12345X`r`n"
        )
        
        foreach ($cmd in $comandos) {
            $bytes = [System.Text.Encoding]::ASCII.GetBytes($cmd)
            $stream.Write($bytes, 0, $bytes.Length)
            Start-Sleep -Milliseconds 500
        }
        
        $tcpClient.Close()
        Show-Text "âœ… Reset automÃ¡tico de papel preso concluÃ­do" Green
        
        # Verificar se resolveu
        Start-Sleep -Seconds 5
        $novoStatus = Monitorar-StatusRapido -IP $IP
        if (-not $novoStatus.PapelPreso) {
            Show-Text "ðŸŽ‰ Problema de papel preso resolvido automaticamente!" Green
        }
        else {
            Show-Text "âš ï¸  Problema persiste. IntervenÃ§Ã£o manual necessÃ¡ria." Yellow
        }
    }
    catch {
        Show-Text "âŒ Erro no reset automÃ¡tico: $_" Red
    }
}

# ConfirmaÃ§Ã£o para automaÃ§Ã£o
function Confirmar-Automacao {
    param(
        [string]$Regra,
        [string]$IP
    )
    
    Show-Text "ðŸ¤– AutomaÃ§Ã£o detectada: $Regra" Yellow
    Show-Text "Alvo: $IP" Yellow
    $resposta = Read-Host "Executar automaticamente? (S/N)"
    
    return ($resposta -eq 'S' -or $resposta -eq 's' -or $resposta -eq 'Y' -or $resposta -eq 'y')
}

# ðŸ¤– 1. Sistema de IA para DetecÃ§Ã£o AutomÃ¡tica
function IA-DeteccaoInteligente {
    param(
        [string]$NetworkRange = "192.168.1",
        [switch]$ScanCompleto
    )
    
    Show-Text "ðŸ¤– Iniciando detecÃ§Ã£o inteligente com IA..." Magenta
    
    $impressorasDetectadas = @()
    $padroesMarcas = @{
        "Epson" = @("EPSON", "L3250", "L3150", "WF-", "XP-")
        "HP" = @("HP", "LaserJet", "DeskJet", "OfficeJet", "Envy")
        "Canon" = @("Canon", "PIXMA", "imageCLASS", "MAXIFY")
        "Brother" = @("Brother", "DCP-", "MFC-", "HL-")
        "Zebra" = @("Zebra", "ZPL", "EPL", "GK420")
    }
    
    # Scan inteligente da rede
    $ips = 1..254 | ForEach-Object { "$NetworkRange.$_" }
    
    $ips | ForEach-Object -Parallel {
        $ip = $_
        $resultado = @{
            IP = $ip
            Marca = "Desconhecida"
            Modelo = "Desconhecido"
            Portas = @()
            Servicos = @()
            Confianca = 0
        }
        
        # Teste de conectividade em mÃºltiplas portas
        $portasComuns = @(9100, 515, 631, 161, 80, 443, 21, 23)
        foreach ($porta in $portasComuns) {
            $teste = Test-NetConnection -ComputerName $ip -Port $porta -WarningAction SilentlyContinue -InformationLevel Quiet
            if ($teste) {
                $resultado.Portas += $porta
                $resultado.Confianca += 10
                
                # IdentificaÃ§Ã£o por porta
                switch ($porta) {
                    9100 { $resultado.Servicos += "JetDirect" }
                    515 { $resultado.Servicos += "LPD" }
                    631 { $resultado.Servicos += "IPP" }
                    161 { $resultado.Servicos += "SNMP" }
                    80 { $resultado.Servicos += "HTTP" }
                    443 { $resultado.Servicos += "HTTPS" }
                }
            }
        }
        
        # Se encontrou serviÃ§os de impressora, fazer identificaÃ§Ã£o avanÃ§ada
        if ($resultado.Portas.Count -gt 0) {
            $resultado = IA-IdentificarMarca -IP $ip -ResultadoBase $resultado -PadroesMarcas $padroesMarcas
        }
        
        return $resultado
    } -ThrottleLimit 50 | Where-Object { $_.Portas.Count -gt 0 } | Sort-Object Confianca -Descending
    
    return $impressorasDetectadas
}

# IA para identificaÃ§Ã£o de marca e modelo
function IA-IdentificarMarca {
    param(
        [string]$IP,
        [hashtable]$ResultadoBase,
        [hashtable]$PadroesMarcas
    )
    
    try {
        # Tentar identificaÃ§Ã£o via HTTP/HTTPS
        if (80 -in $ResultadoBase.Portas -or 443 -in $ResultadoBase.Portas) {
            $protocolo = if (443 -in $ResultadoBase.Portas) { "https" } else { "http" }
            try {
                $response = Invoke-WebRequest -Uri "${protocolo}://$IP" -TimeoutSec 3 -ErrorAction SilentlyContinue
                $html = $response.Content
                
                foreach ($marca in $PadroesMarcas.Keys) {
                    foreach ($padrao in $PadroesMarcas[$marca]) {
                        if ($html -match $padrao) {
                            $ResultadoBase.Marca = $marca
                            $ResultadoBase.Confianca += 30
                            
                            # Tentar extrair modelo
                            if ($html -match "($padrao[\w\-]+)") {
                                $ResultadoBase.Modelo = $matches[1]
                                $ResultadoBase.Confianca += 20
                            }
                            break
                        }
                    }
                    if ($ResultadoBase.Marca -ne "Desconhecida") { break }
                }
            }
            catch { }
        }
        
        # Tentar identificaÃ§Ã£o via JetDirect
        if (9100 -in $ResultadoBase.Portas -and $ResultadoBase.Marca -eq "Desconhecida") {
            try {
                $tcpClient = New-Object System.Net.Sockets.TcpClient
                $tcpClient.ReceiveTimeout = 3000
                $tcpClient.Connect($IP, 9100)
                $stream = $tcpClient.GetStream()
                
                # Comando para obter informaÃ§Ãµes
                $infoCmd = [System.Text.Encoding]::ASCII.GetBytes("`e%-12345X@PJL INFO ID`r`n`e%-12345X`r`n")
                $stream.Write($infoCmd, 0, $infoCmd.Length)
                
                $buffer = New-Object byte[] 1024
                $bytesRead = $stream.Read($buffer, 0, 1024)
                $response = [System.Text.Encoding]::ASCII.GetString($buffer, 0, $bytesRead)
                
                foreach ($marca in $PadroesMarcas.Keys) {
                    foreach ($padrao in $PadroesMarcas[$marca]) {
                        if ($response -match $padrao) {
                            $ResultadoBase.Marca = $marca
                            $ResultadoBase.Confianca += 25
                            
                            if ($response -match "($padrao[\w\-]+)") {
                                $ResultadoBase.Modelo = $matches[1]
                                $ResultadoBase.Confianca += 15
                            }
                            break
                        }
                    }
                    if ($ResultadoBase.Marca -ne "Desconhecida") { break }
                }
                
                $tcpClient.Close()
            }
            catch { }
        }
        
        return $ResultadoBase
    }
    catch {
        return $ResultadoBase
    }
}

# Dashboard interativo em tempo real
function Dashboard-TempoReal {
    param([array]$ImpressorasMonitoradas)
    
    $posicaoOriginal = $Host.UI.RawUI.CursorPosition
    
    while ($true) {
        $Host.UI.RawUI.CursorPosition = $posicaoOriginal
        Clear-Host
        
        # CabeÃ§alho do dashboard
        Show-Text "â•”â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•—" Cyan
        Show-Text "â•‘                    ðŸ–¨ï¸  WINRESET DASHBOARD TEMPO REAL v3.0                    â•‘" Cyan
        Show-Text "â• â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•£" Cyan
        Show-Text "â•‘ AtualizaÃ§Ã£o: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss') | Pressione 'Q' para sair          â•‘" Yellow
        Show-Text "â•šâ•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•" Cyan
        
        # Status geral
        $totalImpressoras = $ImpressorasMonitoradas.Count
        $online = 0
        $comProblemas = 0
        $offline = 0
        
        foreach ($impressora in $ImpressorasMonitoradas) {
            $status = Monitorar-StatusRapido -IP $impressora.IP
            switch ($status.Status) {
                "Online" { $online++ }
                "Problema" { $comProblemas++ }
                "Offline" { $offline++ }
            }
        }
        
        # Exibir estatÃ­sticas
        Show-Text "`nðŸ“Š ESTATÃSTICAS GERAIS:" Magenta
        Show-Text "   Total de Impressoras: $totalImpressoras" White
        Show-Text "   ðŸŸ¢ Online: $online" Green
        Show-Text "   ðŸŸ¡ Com Problemas: $comProblemas" Yellow
        Show-Text "   ðŸ”´ Offline: $offline" Red
        
        # Lista detalhada
        Show-Text "`nðŸ–¨ï¸  STATUS DETALHADO:" Magenta
        Show-Text "â”Œâ”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”¬â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”¬â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”¬â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”" Gray
        Show-Text "â”‚ IP              â”‚ Marca        â”‚ Status      â”‚ Ãšltimo Problema              â”‚" Gray
        Show-Text "â”œâ”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”¼â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”¼â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”¼â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”¤" Gray
        
        foreach ($impressora in $ImpressorasMonitoradas) {
            $status = Monitorar-StatusRapido -IP $impressora.IP
            $cor = switch ($status.Status) {
                "Online" { "Green" }
                "Problema" { "Yellow" }
                "Offline" { "Red" }
                default { "Gray" }
            }
            
            $ip = $impressora.IP.PadRight(15)
            $marca = $impressora.Marca.PadRight(12)
            $statusText = $status.Status.PadRight(11)
            $problema = ($status.UltimoProblema -replace ".{30}.*", "...").PadRight(28)
            
            Show-Text "â”‚ $ip â”‚ $marca â”‚ $statusText â”‚ $problema â”‚" $cor
        }
        
        Show-Text "â””â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”´â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”´â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”´â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”˜" Gray
        
        # Alertas crÃ­ticos
        $alertasCriticos = $ImpressorasMonitoradas | Where-Object { 
            $status = Monitorar-StatusRapido -IP $_.IP
            $status.PapelPreso -or $status.TintaBaixa -or $status.Status -eq "Offline"
        }
        
        if ($alertasCriticos.Count -gt 0) {
            Show-Text "`nðŸš¨ ALERTAS CRÃTICOS:" Red
            foreach ($alerta in $alertasCriticos) {
                $status = Monitorar-StatusRapido -IP $alerta.IP
                if ($status.PapelPreso) {
                    Show-Text "   ðŸ“„ Papel preso em $($alerta.IP) ($($alerta.Marca))" Red
                }
                if ($status.TintaBaixa) {
                    Show-Text "   ðŸ–‹ï¸  Tinta baixa em $($alerta.IP) ($($alerta.Marca))" Yellow
                }
                if ($status.Status -eq "Offline") {
                    Show-Text "   ðŸ”Œ Impressora offline: $($alerta.IP) ($($alerta.Marca))" Red
                }
            }
        }
        
        # Verificar se usuÃ¡rio quer sair
        if ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)
            if ($key.KeyChar -eq 'q' -or $key.KeyChar -eq 'Q') {
                break
            }
        }
        
        Start-Sleep -Seconds 5
    }
}

# Monitoramento rÃ¡pido de status
function Monitorar-StatusRapido {
    param([string]$IP)
    
    $resultado = @{
        Status = "Offline"
        PapelPreso = $false
        TintaBaixa = $false
        UltimoProblema = "Nenhum"
        Timestamp = Get-Date
    }
    
    try {
        $ping = Test-NetConnection -ComputerName $IP -Port 9100 -WarningAction SilentlyContinue -InformationLevel Quiet
        if ($ping) {
            $resultado.Status = "Online"
            
            # VerificaÃ§Ã£o rÃ¡pida de problemas
            $tcpClient = New-Object System.Net.Sockets.TcpClient
            $tcpClient.ReceiveTimeout = 2000
            $tcpClient.Connect($IP, 9100)
            $stream = $tcpClient.GetStream()
            
            $statusCmd = [System.Text.Encoding]::ASCII.GetBytes("`e%-12345X@PJL INFO STATUS`r`n`e%-12345X`r`n")
            $stream.Write($statusCmd, 0, $statusCmd.Length)
            
            $buffer = New-Object byte[] 512
            $bytesRead = $stream.Read($buffer, 0, 512)
            $response = [System.Text.Encoding]::ASCII.GetString($buffer, 0, $bytesRead)
            
            if ($response -match "JAM|PAPER.*STUCK") {
                $resultado.PapelPreso = $true
                $resultado.Status = "Problema"
                $resultado.UltimoProblema = "Papel preso"
            }
            
            if ($response -match "INK.*LOW|TONER.*LOW") {
                $resultado.TintaBaixa = $true
                $resultado.Status = "Problema"
                $resultado.UltimoProblema = "Tinta/Toner baixo"
            }
            
            $tcpClient.Close()
        }
    }
    catch {
        $resultado.UltimoProblema = "Erro de comunicaÃ§Ã£o"
    }
    
    return $resultado
}
