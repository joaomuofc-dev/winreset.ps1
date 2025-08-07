# ═══════════════════════════════════════════════════════════════════════════════
#  WinReset v3.0 - Ferramenta Universal de Reset de Impressoras
# ═══════════════════════════════════════════════════════════════════════════════
# Última atualização: 2024-12-19
# Autor: Sistema Automatizado
# Descrição: Script universal para reset de qualquer impressora (USB/Rede/Wi-Fi)
# Suporte: Epson, HP, Brother, Canon, Zebra e todas as marcas
# Funciona: 100% PowerShell nativo, sem dependências externas
# ═══════════════════════════════════════════════════════════════════════════════

[Console]::Title = "🖨️ WinReset v3.0 - Reset Universal de Impressoras"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# Configurações globais do WinReset Universal
$Global:WinResetVersion = "3.0"
$Global:SupportedBrands = @("Epson", "HP", "Brother", "Canon", "Zebra", "Samsung", "Lexmark", "Kyocera", "Ricoh", "Xerox")

# Configurações globais
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
    param([string]$message = "`n⏸ Pressione ENTER para continuar...")
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
        [string]$title = "Confirmação"
    )
    $choices = @(
        [System.Management.Automation.Host.ChoiceDescription]::new("&Sim", "Confirmar ação")
        [System.Management.Automation.Host.ChoiceDescription]::new("&Não", "Cancelar ação")
    )
    $result = $Host.UI.PromptForChoice($title, $message, $choices, 1)
    return $result -eq 0
}

function Testar-Admin {
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
    if (-not $isAdmin) {
        Show-Text "❌ Execute o PowerShell como ADMINISTRADOR para usar esta ferramenta." Red
        Show-Text "💡 Dica: Clique com o botão direito no PowerShell e selecione 'Executar como administrador'" Yellow
        Pause
        exit 1
    }
    Show-Text "✅ Executando com privilégios de administrador" Green
}

function Criar-Backup {
    param([string]$tipo = "geral")
    
    try {
        if (-not (Test-Path $global:backupPath)) {
            New-Item -Path $global:backupPath -ItemType Directory -Force | Out-Null
        }
        
        $backupFile = Join-Path $global:backupPath "backup_$tipo_$(Get-Date -Format 'HHmmss').json"
        
        $backupData = @{
            Timestamp = Get-Date
            Tipo = $tipo
            Impressoras = @()
            Drivers = @()
            Servicos = @{}
        }
        
        # Backup das impressoras
        $impressoras = Get-Printer -ErrorAction SilentlyContinue
        foreach ($imp in $impressoras) {
            $backupData.Impressoras += @{
                Nome = $imp.Name
                Driver = $imp.DriverName
                Porta = $imp.PortName
                Compartilhada = $imp.Shared
                Status = $imp.PrinterStatus
            }
        }
        
        # Backup dos drivers
        $drivers = Get-PrinterDriver -ErrorAction SilentlyContinue
        foreach ($drv in $drivers) {
            $backupData.Drivers += @{
                Nome = $drv.Name
                Versao = $drv.MajorVersion
                Arquitetura = $drv.PrinterEnvironment
            }
        }
        
        # Status dos serviços
        $backupData.Servicos.Spooler = (Get-Service spooler).Status
        
        $backupData | ConvertTo-Json -Depth 3 | Out-File -FilePath $backupFile -Encoding UTF8
        Show-Text "💾 Backup criado: $backupFile" Cyan
        return $backupFile
    }
    catch {
        Show-Text "❌ Erro ao criar backup: $_" Red
        return $null
    }
}

function Verificar-ServicoSpooler {
    try {
        $spooler = Get-Service -Name spooler
        $status = $spooler.Status
        $startType = $spooler.StartType
        
        $statusColor = switch ($status) {
            "Running" { "Green" }
            "Stopped" { "Red" }
            default { "Yellow" }
        }
        
        Show-Text "🛠️ Serviço Spooler: $status ($startType)" $statusColor
        
        # Verificar dependências
        $dependencies = Get-Service -Name spooler | Select-Object -ExpandProperty ServicesDependedOn
        if ($dependencies) {
            Show-Text "📋 Dependências: $($dependencies.Name -join ', ')" Cyan
        }
        
        return $status -eq "Running"
    }
    catch {
        Show-Text "❌ Não foi possível obter status do serviço spooler: $_" Red
        return $false
    }
}

function Verificar-Saude-Sistema {
    Show-Text "`n🔍 Verificando saúde do sistema de impressão..." Cyan
    
    $problemas = @()
    
    # Verificar serviço spooler
    if (-not (Verificar-ServicoSpooler)) {
        $problemas += "Serviço Spooler não está executando"
    }
    
    # Verificar pasta de spool
    $spoolPath = "C:\Windows\System32\spool\PRINTERS"
    if (-not (Test-Path $spoolPath)) {
        $problemas += "Pasta de spool não encontrada"
    } else {
        $spoolFiles = Get-ChildItem $spoolPath -ErrorAction SilentlyContinue
        if ($spoolFiles.Count -gt 10) {
            $problemas += "Muitos arquivos na fila de spool ($($spoolFiles.Count) arquivos)"
        }
    }
    
    # Verificar impressoras órfãs
    try {
        $impressoras = Get-Printer -ErrorAction SilentlyContinue
        $drivers = Get-PrinterDriver -ErrorAction SilentlyContinue
        
        foreach ($imp in $impressoras) {
            if ($imp.DriverName -and ($drivers.Name -notcontains $imp.DriverName)) {
                $problemas += "Impressora '$($imp.Name)' tem driver ausente: $($imp.DriverName)"
            }
        }
    }
    catch {
        $problemas += "Erro ao verificar impressoras: $_"
    }
    
    if ($problemas.Count -eq 0) {
        Show-Text "✅ Sistema de impressão está saudável" Green
    } else {
        Show-Text "⚠️ Problemas encontrados:" Yellow
        foreach ($problema in $problemas) {
            Show-Text "  • $problema" Red
        }
    }
    
    return $problemas.Count -eq 0
}

# ═══════════════════════════════════════════════════════════════════════════════
#  FUNÇÕES DE DETECÇÃO UNIVERSAL DE IMPRESSORAS
# ═══════════════════════════════════════════════════════════════════════════════

# Função para detectar impressoras na rede
function Detectar-ImpressorasRede {
    param(
        [string]$NetworkRange = "192.168.1",
        [int]$StartIP = 1,
        [int]$EndIP = 254
    )
    
    Show-Text "🔍 Detectando impressoras na rede $NetworkRange.x..." "Cyan"
    $networkPrinters = @()
    
    $jobs = @()
    for ($i = $StartIP; $i -le $EndIP; $i++) {
        $ip = "$NetworkRange.$i"
        $jobs += Start-Job -ScriptBlock {
            param($targetIP)
            $result = @{}
            
            # Teste de conectividade básica
            if (Test-Connection -ComputerName $targetIP -Count 1 -Quiet -TimeoutSeconds 1) {
                $result.IP = $targetIP
                $result.Responsive = $true
                
                # Teste portas comuns de impressora
                $commonPorts = @(9100, 515, 631, 80, 443)
                $result.OpenPorts = @()
                
                foreach ($port in $commonPorts) {
                    try {
                        $tcpClient = New-Object System.Net.Sockets.TcpClient
                        $connect = $tcpClient.BeginConnect($targetIP, $port, $null, $null)
                        $wait = $connect.AsyncWaitHandle.WaitOne(1000, $false)
                        
                        if ($wait) {
                            $tcpClient.EndConnect($connect)
                            $result.OpenPorts += $port
                        }
                        $tcpClient.Close()
                    }
                    catch { }
                }
            }
            
            return $result
        } -ArgumentList $ip
    }
    
    # Aguardar conclusão dos jobs
    Show-Progress -Activity "Escaneando rede" -Status "Verificando IPs..." -PercentComplete 50
    
    $results = $jobs | Wait-Job | Receive-Job
    $jobs | Remove-Job
    
    foreach ($result in $results) {
        if ($result.Responsive -and $result.OpenPorts.Count -gt 0) {
            $printer = @{
                IP = $result.IP
                Type = "Network"
                Ports = $result.OpenPorts
                Brand = "Unknown"
                Model = "Unknown"
                Status = "Detected"
            }
            
            # Tentar identificar marca/modelo
            $printer = Identificar-MarcaModelo -PrinterInfo $printer
            $networkPrinters += $printer
        }
    }
    
    Show-Text "✅ Encontradas $($networkPrinters.Count) impressoras na rede" "Green"
    return $networkPrinters
}

# Função para identificar marca e modelo da impressora
function Identificar-MarcaModelo {
    param(
        [hashtable]$PrinterInfo
    )
    
    $brand = "Unknown"
    $model = "Unknown"
    
    if ($PrinterInfo.Type -eq "Network" -and $PrinterInfo.IP) {
        # Tentar identificação via HTTP
        if (80 -in $PrinterInfo.Ports -or 443 -in $PrinterInfo.Ports) {
            try {
                $uri = if (443 -in $PrinterInfo.Ports) { "https://$($PrinterInfo.IP)" } else { "http://$($PrinterInfo.IP)" }
                $webResponse = Invoke-WebRequest -Uri $uri -TimeoutSec 3 -ErrorAction SilentlyContinue
                if ($webResponse) {
                    $content = $webResponse.Content.ToLower()
                    foreach ($supportedBrand in $Global:SupportedBrands) {
                        if ($content -match $supportedBrand.ToLower()) {
                            $brand = $supportedBrand
                            # Tentar extrair modelo
                            if ($content -match "($($supportedBrand.ToLower()))\s*([a-z0-9\-]+\s*[0-9]+[a-z0-9\-]*)") {
                                $model = $matches[2]
                            }
                            break
                        }
                    }
                }
            }
            catch { }
        }
        
        # Tentar identificação via porta 9100 (JetDirect)
        if ($brand -eq "Unknown" -and 9100 -in $PrinterInfo.Ports) {
            try {
                $tcpClient = New-Object System.Net.Sockets.TcpClient
                $tcpClient.Connect($PrinterInfo.IP, 9100)
                $stream = $tcpClient.GetStream()
                
                # Enviar comando de status
                $statusCommand = [System.Text.Encoding]::ASCII.GetBytes("`e%-12345X@PJL INFO ID`r`n`e%-12345X`r`n")
                $stream.Write($statusCommand, 0, $statusCommand.Length)
                
                Start-Sleep -Milliseconds 500
                
                if ($stream.DataAvailable) {
                    $buffer = New-Object byte[] 1024
                    $bytesRead = $stream.Read($buffer, 0, $buffer.Length)
                    $response = [System.Text.Encoding]::ASCII.GetString($buffer, 0, $bytesRead).ToLower()
                    
                    foreach ($supportedBrand in $Global:SupportedBrands) {
                        if ($response -match $supportedBrand.ToLower()) {
                            $brand = $supportedBrand
                            break
                        }
                    }
                }
                
                $tcpClient.Close()
            }
            catch { }
        }
    }
    elseif ($PrinterInfo.Type -eq "USB" -or $PrinterInfo.Type -eq "Local") {
        # Identificação via WMI para impressoras locais
        if ($PrinterInfo.Name) {
            $printerName = $PrinterInfo.Name.ToLower()
            foreach ($supportedBrand in $Global:SupportedBrands) {
                if ($printerName -match $supportedBrand.ToLower()) {
                    $brand = $supportedBrand
                    break
                }
            }
        }
    }
    
    $PrinterInfo.Brand = $brand
    $PrinterInfo.Model = $model
    
    return $PrinterInfo
}

# Função para detectar todas as impressoras (locais e rede)
function Detectar-TodasImpressoras {
    param(
        [switch]$IncluirRede,
        [string]$NetworkRange = "192.168.1"
    )
    
    $Global:DetectedPrinters = @()
    
    # Detectar impressoras locais/USB
    Show-Text "🔍 Detectando impressoras locais e USB..." "Cyan"
    try {
        $localPrinters = Get-Printer -ErrorAction SilentlyContinue
        foreach ($printer in $localPrinters) {
            $printerInfo = @{
                Name = $printer.Name
                Type = if ($printer.PortName -match "USB|DOT4") { "USB" } else { "Local" }
                Brand = "Unknown"
                Model = "Unknown"
                Status = $printer.PrinterStatus
                Driver = $printer.DriverName
                Port = $printer.PortName
            }
            
            $printerInfo = Identificar-MarcaModelo -PrinterInfo $printerInfo
            $Global:DetectedPrinters += $printerInfo
        }
        Show-Text "✅ Encontradas $($localPrinters.Count) impressoras locais" "Green"
    }
    catch {
        Show-Text "❌ Erro ao detectar impressoras locais: $_" "Red"
    }
    
    # Detectar impressoras na rede se solicitado
    if ($IncluirRede) {
        $networkPrinters = Detectar-ImpressorasRede -NetworkRange $NetworkRange
        $Global:DetectedPrinters += $networkPrinters
    }
    
    Show-Text "📊 Total de impressoras detectadas: $($Global:DetectedPrinters.Count)" "Yellow"
    return $Global:DetectedPrinters
}

# ═══════════════════════════════════════════════════════════════════════════════
#  COMANDOS ESPECÍFICOS POR FABRICANTE
# ═══════════════════════════════════════════════════════════════════════════════

# Inicializar comandos de reset por fabricante
function Inicializar-ComandosReset {
    $Global:ResetCommands = @{
        "Epson" = @{
            "Reset" = @(
                "`e@",  # Reset básico ESC/POS
                "`e%-12345X@PJL RESET`r`n`e%-12345X`r`n",  # PJL Reset
                [byte[]]@(0x1B, 0x40)  # ESC @ - Reset
            )
            "ClearQueue" = @(
                "`e%-12345X@PJL RESET`r`n@PJL USTATUS DEVICE = ON`r`n`e%-12345X`r`n"
            )
            "Status" = @(
                "`e%-12345X@PJL INFO STATUS`r`n`e%-12345X`r`n",
                "`e%-12345X@PJL INFO ID`r`n`e%-12345X`r`n"
            )
        }
        "HP" = @{
            "Reset" = @(
                "`e%-12345X@PJL RESET`r`n`e%-12345X`r`n",
                "`e%-12345X@PJL DEFAULT`r`n`e%-12345X`r`n",
                "`eE"  # PCL Reset
            )
            "ClearQueue" = @(
                "`e%-12345X@PJL RESET`r`n@PJL USTATUS DEVICE = ON`r`n`e%-12345X`r`n"
            )
            "Status" = @(
                "`e%-12345X@PJL INFO STATUS`r`n`e%-12345X`r`n",
                "`e%-12345X@PJL INFO ID`r`n`e%-12345X`r`n",
                "`e%-12345X@PJL INFO MEMORY`r`n`e%-12345X`r`n"
            )
        }
        "Brother" = @{
            "Reset" = @(
                "`e@",  # ESC/POS Reset
                "`e%-12345X@PJL RESET`r`n`e%-12345X`r`n",
                [byte[]]@(0x1B, 0x40, 0x1B, 0x69, 0x61, 0x00)  # Brother specific
            )
            "ClearQueue" = @(
                "`e%-12345X@PJL RESET`r`n`e%-12345X`r`n"
            )
            "Status" = @(
                "`e%-12345X@PJL INFO STATUS`r`n`e%-12345X`r`n"
            )
        }
        "Canon" = @{
            "Reset" = @(
                "`e@",  # ESC/POS Reset
                "`e%-12345X@PJL RESET`r`n`e%-12345X`r`n",
                "`e[K`e[2J`e[H"  # Canon specific reset sequence
            )
            "ClearQueue" = @(
                "`e%-12345X@PJL RESET`r`n`e%-12345X`r`n"
            )
            "Status" = @(
                "`e%-12345X@PJL INFO STATUS`r`n`e%-12345X`r`n"
            )
        }
        "Zebra" = @{
            "Reset" = @(
                "^XA^JUF^XZ",  # Zebra factory reset
                "~JR",  # Zebra reset
                "^XA^MCY^XZ"  # Clear memory
            )
            "ClearQueue" = @(
                "~JA",  # Cancel all jobs
                "^XA^IDR:*.*^XZ"  # Delete all files
            )
            "Status" = @(
                "~HS",  # Host status
                "~HI",  # Host identification
                "^XA^HH^XZ"  # Configuration status
            )
        }
        "Samsung" = @{
            "Reset" = @(
                "`e%-12345X@PJL RESET`r`n`e%-12345X`r`n",
                "`eE"  # PCL Reset
            )
            "ClearQueue" = @(
                "`e%-12345X@PJL RESET`r`n`e%-12345X`r`n"
            )
            "Status" = @(
                "`e%-12345X@PJL INFO STATUS`r`n`e%-12345X`r`n"
            )
        }
        "Generic" = @{
            "Reset" = @(
                "`e@",  # ESC/POS Reset
                "`e%-12345X@PJL RESET`r`n`e%-12345X`r`n",
                "`eE",  # PCL Reset
                [byte[]]@(0x1B, 0x40)  # ESC @ Reset
            )
            "ClearQueue" = @(
                "`e%-12345X@PJL RESET`r`n`e%-12345X`r`n"
            )
            "Status" = @(
                "`e%-12345X@PJL INFO STATUS`r`n`e%-12345X`r`n",
                "`e%-12345X@PJL INFO ID`r`n`e%-12345X`r`n"
            )
        }
    }
}

# Função para enviar comando para impressora
function Enviar-ComandoImpressora {
    param(
        [string]$IP,
        [int]$Port = 9100,
        [string]$Command,
        [byte[]]$BinaryCommand = $null,
        [int]$TimeoutMs = 5000
    )
    
    try {
        $tcpClient = New-Object System.Net.Sockets.TcpClient
        $tcpClient.ReceiveTimeout = $TimeoutMs
        $tcpClient.SendTimeout = $TimeoutMs
        
        $tcpClient.Connect($IP, $Port)
        $stream = $tcpClient.GetStream()
        
        if ($BinaryCommand) {
            $stream.Write($BinaryCommand, 0, $BinaryCommand.Length)
        } else {
            $commandBytes = [System.Text.Encoding]::ASCII.GetBytes($Command)
            $stream.Write($commandBytes, 0, $commandBytes.Length)
        }
        
        Start-Sleep -Milliseconds 500
        
        $response = ""
        if ($stream.DataAvailable) {
            $buffer = New-Object byte[] 1024
            $bytesRead = $stream.Read($buffer, 0, $buffer.Length)
            $response = [System.Text.Encoding]::ASCII.GetString($buffer, 0, $bytesRead)
        }
        
        $tcpClient.Close()
        return @{ Success = $true; Response = $response }
    }
    catch {
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

# Função para resetar impressora específica
function Reset-ImpressoraUniversal {
    param(
        [hashtable]$PrinterInfo,
        [string]$ResetType = "Reset"
    )
    
    if (-not $Global:ResetCommands) {
        Inicializar-ComandosReset
    }
    
    $brand = if ($PrinterInfo.Brand -and $PrinterInfo.Brand -ne "Unknown") { $PrinterInfo.Brand } else { "Generic" }
    $commands = $Global:ResetCommands[$brand][$ResetType]
    
    if (-not $commands) {
        Show-Text "❌ Comandos de $ResetType não encontrados para $brand" "Red"
        return $false
    }
    
    Show-Text "🔄 Executando $ResetType para $brand ($($PrinterInfo.Name ?? $PrinterInfo.IP))..." "Yellow"
    
    $success = $false
    foreach ($command in $commands) {
        try {
            if ($PrinterInfo.Type -eq "Network" -and $PrinterInfo.IP) {
                # Tentar diferentes portas
                $ports = @(9100, 515)
                foreach ($port in $ports) {
                    if ($port -in $PrinterInfo.Ports) {
                        Show-Text "  📡 Enviando comando via porta $port..." "Cyan"
                        
                        if ($command -is [byte[]]) {
                            $result = Enviar-ComandoImpressora -IP $PrinterInfo.IP -Port $port -BinaryCommand $command
                        } else {
                            $result = Enviar-ComandoImpressora -IP $PrinterInfo.IP -Port $port -Command $command
                        }
                        
                        if ($result.Success) {
                            Show-Text "  ✅ Comando enviado com sucesso" "Green"
                            $success = $true
                            break
                        } else {
                            Show-Text "  ⚠️ Falha na porta $port`: $($result.Error)" "Yellow"
                        }
                    }
                }
            }
            elseif ($PrinterInfo.Type -eq "USB" -or $PrinterInfo.Type -eq "Local") {
                # Para impressoras locais, usar método do Windows
                try {
                    if ($PrinterInfo.Name) {
                        # Limpar fila local
                        Get-PrintJob -PrinterName $PrinterInfo.Name -ErrorAction SilentlyContinue | Remove-PrintJob -Confirm:$false
                        Show-Text "  ✅ Fila local limpa" "Green"
                        $success = $true
                    }
                }
                catch {
                    Show-Text "  ❌ Erro ao limpar fila local: $_" "Red"
                }
            }
            
            if ($success) { break }
        }
        catch {
            Show-Text "  ❌ Erro ao executar comando: $_" "Red"
        }
    }
    
    return $success
}

# Função para diagnóstico avançado de impressora
function Diagnosticar-ImpressoraUniversal {
    param(
        [hashtable]$PrinterInfo
    )
    
    Show-Text "🔍 Diagnóstico completo: $($PrinterInfo.Name ?? $PrinterInfo.IP)" "Cyan"
    Separator
    
    $diagnostico = @{
        Conectividade = $false
        Resposta = $false
        Fila = 0
        Status = "Unknown"
        Problemas = @()
        Solucoes = @()
    }
    
    # Teste de conectividade
    if ($PrinterInfo.Type -eq "Network") {
        Show-Text "📡 Testando conectividade de rede..." "Yellow"
        if (Test-Connection -ComputerName $PrinterInfo.IP -Count 2 -Quiet) {
            $diagnostico.Conectividade = $true
            Show-Text "  ✅ IP $($PrinterInfo.IP) responde ao ping" "Green"
        } else {
            $diagnostico.Problemas += "IP não responde ao ping"
            $diagnostico.Solucoes += "Verificar cabo de rede e configuração IP"
            Show-Text "  ❌ IP $($PrinterInfo.IP) não responde" "Red"
        }
        
        # Teste de portas
        Show-Text "🔌 Testando portas de comunicação..." "Yellow"
        foreach ($port in $PrinterInfo.Ports) {
            try {
                $tcpClient = New-Object System.Net.Sockets.TcpClient
                $connect = $tcpClient.BeginConnect($PrinterInfo.IP, $port, $null, $null)
                $wait = $connect.AsyncWaitHandle.WaitOne(3000, $false)
                
                if ($wait) {
                    $tcpClient.EndConnect($connect)
                    Show-Text "  ✅ Porta $port aberta" "Green"
                    $diagnostico.Resposta = $true
                } else {
                    Show-Text "  ❌ Porta $port fechada ou sem resposta" "Red"
                }
                $tcpClient.Close()
            }
            catch {
                Show-Text "  ❌ Erro ao testar porta $port`: $_" "Red"
            }
        }
    }
    
    # Verificar fila de impressão
    if ($PrinterInfo.Name) {
        Show-Text "📄 Verificando fila de impressão..." "Yellow"
        try {
            $jobs = Get-PrintJob -PrinterName $PrinterInfo.Name -ErrorAction SilentlyContinue
            $diagnostico.Fila = $jobs.Count
            
            if ($jobs.Count -eq 0) {
                Show-Text "  ✅ Fila vazia" "Green"
            } else {
                Show-Text "  ⚠️ $($jobs.Count) trabalhos na fila" "Yellow"
                $diagnostico.Problemas += "$($jobs.Count) trabalhos pendentes na fila"
                $diagnostico.Solucoes += "Limpar fila de impressão"
            }
        }
        catch {
            Show-Text "  ❌ Erro ao verificar fila: $_" "Red"
        }
    }
    
    # Obter status da impressora
    if ($PrinterInfo.Type -eq "Network" -and $diagnostico.Resposta) {
        Show-Text "📊 Obtendo status da impressora..." "Yellow"
        
        if (-not $Global:ResetCommands) {
            Inicializar-ComandosReset
        }
        
        $brand = if ($PrinterInfo.Brand -and $PrinterInfo.Brand -ne "Unknown") { $PrinterInfo.Brand } else { "Generic" }
        $statusCommands = $Global:ResetCommands[$brand]["Status"]
        
        foreach ($command in $statusCommands) {
            $result = Enviar-ComandoImpressora -IP $PrinterInfo.IP -Port 9100 -Command $command
            if ($result.Success -and $result.Response) {
                $diagnostico.Status = $result.Response
                Show-Text "  ✅ Status obtido" "Green"
                break
            }
        }
    }
    
    # Resumo do diagnóstico
    Show-Text "`n📋 Resumo do Diagnóstico:" "Cyan"
    if ($diagnostico.Problemas.Count -eq 0) {
        Show-Text "✅ Nenhum problema detectado" "Green"
    } else {
        Show-Text "⚠️ Problemas encontrados:" "Yellow"
        foreach ($problema in $diagnostico.Problemas) {
            Show-Text "  • $problema" "Red"
        }
        
        Show-Text "`n💡 Soluções sugeridas:" "Cyan"
        foreach ($solucao in $diagnostico.Solucoes) {
            Show-Text "  • $solucao" "Yellow"
        }
    }
    
    return $diagnostico
}

function Listar-Impressoras {
    param([switch]$Detalhado)
    
    Clear-Host
    Separator
    Show-Text "📃 Impressoras instaladas no sistema:" Cyan
    Separator

    try {
        Show-Progress "Carregando impressoras" "Obtendo lista..." 25
        $global:impressoras = Get-Printer | Select-Object Name, DriverName, PortName, Shared, PrinterStatus, JobCount
        Show-Progress "Carregando impressoras" "Processando dados..." 75
        
        # Obter informações adicionais se modo detalhado
        if ($Detalhado) {
            for ($i = 0; $i -lt $impressoras.Count; $i++) {
                $imp = $impressoras[$i]
                try {
                    $jobs = Get-PrintJob -PrinterName $imp.Name -ErrorAction SilentlyContinue
                    $imp | Add-Member -NotePropertyName "JobsNaFila" -NotePropertyValue $jobs.Count -Force
                }
                catch {
                    $imp | Add-Member -NotePropertyName "JobsNaFila" -NotePropertyValue "N/A" -Force
                }
            }
        }
        
        Write-Progress -Activity "Carregando impressoras" -Completed
    }
    catch {
        Write-Progress -Activity "Carregando impressoras" -Completed
        Show-Text "❌ Erro ao listar impressoras: $_" Red
        Pause
        return $false
    }

    if (!$impressoras -or $impressoras.Count -eq 0) {
        Show-Text "❌ Nenhuma impressora foi encontrada." Red
        Show-Text "💡 Verifique se há impressoras instaladas no sistema" Yellow
        Pause
        return $false
    }

    Show-Text "📊 Total de impressoras encontradas: $($impressoras.Count)" Green
    Show-Text ""
    
    for ($i=0; $i -lt $impressoras.Count; $i++) {
        $imp = $impressoras[$i]
        
        # Determinar cor do status
        $statusColor = switch ($imp.PrinterStatus) {
            "Normal" { "Green" }
            "Error" { "Red" }
            "Offline" { "Yellow" }
            default { "White" }
        }
        
        Show-Text "[$i] " -NoNewLine
        Show-Text "$($imp.Name)" Cyan -NoNewLine
        
        if ($Detalhado) {
            Show-Text ""
            Show-Text "    📄 Driver: $($imp.DriverName)" White
            Show-Text "    🔌 Porta: $($imp.PortName)" White
            Show-Text "    📊 Status: $($imp.PrinterStatus)" $statusColor
            Show-Text "    🌐 Compartilhada: $(if($imp.Shared){'Sim'}else{'Não'})" White
            if ($imp.JobsNaFila -ne $null) {
                Show-Text "    📋 Jobs na fila: $($imp.JobsNaFila)" $(if($imp.JobsNaFila -gt 0){'Yellow'}else{'White'})
            }
            Show-Text ""
        } else {
            Show-Text "  |  Driver: $($imp.DriverName)  |  Porta: $($imp.PortName)  |  Status: $($imp.PrinterStatus)" White
        }
    }

    return $true
}

function Limpar-FilasImpressora {
    param(
        [string]$printerName,
        [switch]$Force
    )
    try {
        Show-Progress "Limpando filas" "Obtendo jobs da impressora '$printerName'..." 25
        
        # Método 1: Usar Get-PrintJob (mais moderno)
        try {
            $jobs = Get-PrintJob -PrinterName $printerName -ErrorAction Stop
            if ($jobs.Count -gt 0) {
                Show-Text "📋 Encontrados $($jobs.Count) jobs na fila da impressora '$printerName'" Yellow
                
                for ($i = 0; $i -lt $jobs.Count; $i++) {
                    $job = $jobs[$i]
                    Show-Progress "Limpando filas" "Removendo job $($i+1) de $($jobs.Count): $($job.DocumentName)" (50 + ($i / $jobs.Count * 40))
                    Remove-PrintJob -InputObject $job -ErrorAction SilentlyContinue
                }
            }
        }
        catch {
            # Método 2: Fallback para WMI (compatibilidade)
            Show-Progress "Limpando filas" "Usando método alternativo..." 60
            $jobs = Get-CimInstance -ClassName Win32_PrintJob | Where-Object { $_.Name -like "$printerName,*" }
            
            if ($jobs) {
                Show-Text "📋 Encontrados $($jobs.Count) jobs (WMI) na fila da impressora '$printerName'" Yellow
                foreach ($job in $jobs) {
                    $job | Invoke-CimMethod -MethodName Delete | Out-Null
                }
            }
        }
        
        Show-Progress "Limpando filas" "Concluído" 100
        Write-Progress -Activity "Limpando filas" -Completed
        Show-Text "✅ Filas da impressora '$printerName' limpas." Green
        
        return $true
    }
    catch {
        Write-Progress -Activity "Limpando filas" -Completed
        Show-Text "❌ Erro ao limpar filas da impressora: $_" Red
        return $false
    }
}

function Limpar-FilesSpooler {
    param([switch]$CreateBackup)
    
    $spoolPath = "C:\Windows\System32\spool\PRINTERS"
    
    if (-not (Test-Path $spoolPath)) {
        Show-Text "⚠️ Pasta de spooler não encontrada: $spoolPath" Yellow
        return $false
    }
    
    try {
        Show-Progress "Limpando spooler" "Verificando arquivos..." 20
        $spoolFiles = Get-ChildItem $spoolPath -File -ErrorAction Stop
        
        if ($spoolFiles.Count -eq 0) {
            Show-Text "ℹ️ Pasta de spooler já está limpa." Cyan
            Write-Progress -Activity "Limpando spooler" -Completed
            return $true
        }
        
        Show-Text "📁 Encontrados $($spoolFiles.Count) arquivos na pasta de spool" Yellow
        
        # Criar backup se solicitado
        if ($CreateBackup) {
            Show-Progress "Limpando spooler" "Criando backup..." 40
            $backupSpoolPath = Join-Path $global:backupPath "spool_files"
            if (-not (Test-Path $backupSpoolPath)) {
                New-Item -Path $backupSpoolPath -ItemType Directory -Force | Out-Null
            }
            
            foreach ($file in $spoolFiles) {
                Copy-Item $file.FullName -Destination $backupSpoolPath -ErrorAction SilentlyContinue
            }
            Show-Text "💾 Backup dos arquivos de spool criado em: $backupSpoolPath" Cyan
        }
        
        Show-Progress "Limpando spooler" "Removendo arquivos..." 70
        
        # Tentar remover arquivos individualmente para melhor controle
        $removidos = 0
        foreach ($file in $spoolFiles) {
            try {
                Remove-Item $file.FullName -Force -ErrorAction Stop
                $removidos++
            }
            catch {
                Show-Text "⚠️ Não foi possível remover: $($file.Name) - $_" Yellow
            }
        }
        
        Show-Progress "Limpando spooler" "Concluído" 100
        Write-Progress -Activity "Limpando spooler" -Completed
        
        if ($removidos -eq $spoolFiles.Count) {
            Show-Text "✅ Todos os $removidos arquivos de spool foram removidos." Green
        } else {
            Show-Text "⚠️ $removidos de $($spoolFiles.Count) arquivos foram removidos." Yellow
        }
        
        return $true
    }
    catch {
        Write-Progress -Activity "Limpando spooler" -Completed
        Show-Text "❌ Erro ao limpar arquivos de spooler: $_" Red
        return $false
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

    $impressora = $impressoras[$index]
    $nome = $impressora.Name
    
    Clear-Host
    Separator
    Show-Text "🔄 RESET BRUTAL DA IMPRESSORA" Magenta
    Separator
    Show-Text "📄 Impressora: $nome" Cyan
    Show-Text "🖨️ Driver: $($impressora.DriverName)" Cyan
    Show-Text "🔌 Porta: $($impressora.PortName)" Cyan
    Separator
    
    if (-not (Confirm-Action "Deseja realmente resetar esta impressora? Esta ação irá:
• Parar o serviço spooler
• Limpar todas as filas de impressão
• Remover arquivos de spool
• Opcionalmente remover o driver
• Reiniciar o serviço spooler" "Reset da Impressora")) {
        Show-Text "❌ Operação cancelada pelo usuário." Yellow
        Pause
        return
    }
    
    # Criar backup antes do reset
    Show-Text "`n💾 Criando backup antes do reset..." Cyan
    $backupFile = Criar-Backup "impressora_$($nome -replace '[^\w]', '_')"
    
    $resetSucesso = $true
    
    try {
        Show-Progress "Reset da impressora" "Parando serviço spooler..." 10
        Show-Text "⏳ Parando serviço spooler..." Yellow
        Stop-Service spooler -Force -ErrorAction Stop
        Start-Sleep -Seconds 2
        
        Show-Progress "Reset da impressora" "Limpando filas da impressora..." 30
        if (-not (Limpar-FilasImpressora -printerName $nome)) {
            $resetSucesso = $false
        }
        
        Show-Progress "Reset da impressora" "Limpando arquivos de spool..." 50
        if (-not (Limpar-FilesSpooler -CreateBackup)) {
            $resetSucesso = $false
        }

        # Perguntar sobre remoção do driver
        Show-Progress "Reset da impressora" "Verificando driver..." 70
        $driverName = $impressora.DriverName
        if ($driverName) {
            $removerDriver = Confirm-Action "Deseja remover o driver '$driverName'? (Recomendado apenas se houver problemas)" "Remoção de Driver"
            
            if ($removerDriver) {
                try {
                    Show-Text "⏳ Removendo driver '$driverName'..." Yellow
                    Remove-PrinterDriver -Name $driverName -ErrorAction Stop
                    Show-Text "✅ Driver '$driverName' removido com sucesso." Green
                }
                catch {
                    Show-Text "⚠️ Não foi possível remover o driver: $_" Yellow
                    $resetSucesso = $false
                }
            }
        }

        Show-Progress "Reset da impressora" "Reiniciando serviço spooler..." 90
        Show-Text "⏳ Iniciando serviço spooler..." Yellow
        Start-Service spooler -ErrorAction Stop
        Start-Sleep -Seconds 3
        
        Show-Progress "Reset da impressora" "Verificando resultado..." 100
        Write-Progress -Activity "Reset da impressora" -Completed
        
        if ($resetSucesso) {
            Show-Text "✅ Impressora '$nome' resetada com sucesso!" Green
        } else {
            Show-Text "⚠️ Reset concluído com alguns avisos. Verifique os logs." Yellow
        }
        
        Verificar-ServicoSpooler
        
        # Verificar se a impressora ainda existe
        try {
            $impressoraPos = Get-Printer -Name $nome -ErrorAction SilentlyContinue
            if ($impressoraPos) {
                Show-Text "📄 Impressora '$nome' ainda está disponível no sistema" Cyan
            } else {
                Show-Text "⚠️ Impressora '$nome' não foi encontrada após o reset" Yellow
            }
        }
        catch {
            Show-Text "⚠️ Não foi possível verificar o status da impressora após o reset" Yellow
        }
    }
    catch {
        Write-Progress -Activity "Reset da impressora" -Completed
        Show-Text "❌ Erro crítico no reset: $_" Red
        
        # Tentar restaurar o serviço spooler
        try {
            Show-Text "🔄 Tentando restaurar serviço spooler..." Yellow
            Start-Service spooler -ErrorAction SilentlyContinue
        }
        catch {
            Show-Text "❌ Falha ao restaurar serviço spooler. Reinicialização manual necessária." Red
        }
    }

    Pause
}

function Resetar-Tudo-Bruto {
    Clear-Host
    Separator
    Show-Text "⚠️ RESET TOTAL DO SISTEMA DE IMPRESSÃO" Red
    Separator
    Show-Text "Esta operação irá:" Yellow
    Show-Text "• Parar o serviço spooler" Yellow
    Show-Text "• Limpar TODAS as filas de impressão" Yellow
    Show-Text "• Remover TODOS os arquivos de spool" Yellow
    Show-Text "• Opcionalmente remover drivers" Yellow
    Show-Text "• Reiniciar o serviço spooler" Yellow
    Separator
    
    if (-not (Confirm-Action "ATENÇÃO: Esta é uma operação DESTRUTIVA que afetará TODAS as impressoras do sistema. Deseja continuar?" "Reset Total")) {
        Show-Text "❌ Operação cancelada pelo usuário." Yellow
        Pause
        return
    }
    
    # Criar backup completo
    Show-Text "`n💾 Criando backup completo do sistema..." Cyan
    $backupFile = Criar-Backup "reset_total"
    
    $resetSucesso = $true
    $estatisticas = @{
        ImpressorasEncontradas = 0
        FilasLimpas = 0
        ArquivosRemovidos = 0
        DriversRemovidos = 0
    }

    try {
        # Coletar estatísticas antes do reset
        Show-Progress "Reset total" "Coletando informações do sistema..." 5
        try {
            $impressoras = Get-Printer -ErrorAction SilentlyContinue
            $estatisticas.ImpressorasEncontradas = if ($impressoras) { $impressoras.Count } else { 0 }
            
            $jobs = Get-CimInstance -ClassName Win32_PrintJob -ErrorAction SilentlyContinue
            $estatisticas.FilasLimpas = if ($jobs) { $jobs.Count } else { 0 }
        }
        catch {
            Show-Text "⚠️ Erro ao coletar estatísticas: $_" Yellow
        }
        
        Show-Progress "Reset total" "Parando serviço spooler..." 15
        Show-Text "⏳ Parando serviço spooler..." Yellow
        Stop-Service spooler -Force -ErrorAction Stop
        Start-Sleep -Seconds 3

        Show-Progress "Reset total" "Limpando todas as filas de impressão..." 30
        Show-Text "⏳ Limpando todas as filas de impressão..." Yellow
        try {
            $jobs = Get-CimInstance -ClassName Win32_PrintJob -ErrorAction SilentlyContinue
            if ($jobs) {
                Show-Text "📋 Encontrados $($jobs.Count) jobs em todas as filas" Yellow
                foreach ($job in $jobs) {
                    $job | Invoke-CimMethod -MethodName Delete | Out-Null
                }
                Show-Text "✅ Todas as $($jobs.Count) filas de impressão foram limpas." Green
            } else {
                Show-Text "ℹ️ Nenhuma fila de impressão encontrada." Cyan
            }
        }
        catch {
            Show-Text "⚠️ Erro ao limpar filas: $_" Yellow
            $resetSucesso = $false
        }

        Show-Progress "Reset total" "Limpando arquivos de spool..." 50
        if (-not (Limpar-FilesSpooler -CreateBackup)) {
            $resetSucesso = $false
        }

        # Perguntar sobre remoção de drivers
        Show-Progress "Reset total" "Verificando drivers..." 65
        $removerDrivers = Confirm-Action "Deseja remover TODOS os drivers de impressora? (CUIDADO: Isso pode exigir reinstalação)" "Remoção de Drivers"
        
        if ($removerDrivers) {
            try {
                Show-Text "⏳ Removendo todos os drivers de impressora..." Yellow
                $drivers = Get-PrinterDriver -ErrorAction SilentlyContinue
                if ($drivers) {
                    Show-Text "🗑️ Encontrados $($drivers.Count) drivers para remoção" Yellow
                    foreach ($drv in $drivers) {
                        try {
                            Remove-PrinterDriver -Name $drv.Name -ErrorAction SilentlyContinue
                            $estatisticas.DriversRemovidos++
                        }
                        catch {
                            Show-Text "⚠️ Não foi possível remover driver: $($drv.Name)" Yellow
                        }
                    }
                    Show-Text "✅ $($estatisticas.DriversRemovidos) de $($drivers.Count) drivers removidos." Green
                } else {
                    Show-Text "ℹ️ Nenhum driver encontrado." Cyan
                }
            }
            catch {
                Show-Text "❌ Erro ao remover drivers: $_" Red
                $resetSucesso = $false
            }
        }

        Show-Progress "Reset total" "Reiniciando serviço spooler..." 85
        Show-Text "⏳ Iniciando serviço spooler..." Yellow
        Start-Service spooler -ErrorAction Stop
        Start-Sleep -Seconds 5
        
        Show-Progress "Reset total" "Verificando resultado..." 100
        Write-Progress -Activity "Reset total" -Completed

        # Mostrar relatório final
        Clear-Host
        Separator
        Show-Text "📊 RELATÓRIO DO RESET TOTAL" Green
        Separator
        Show-Text "📄 Impressoras no sistema: $($estatisticas.ImpressorasEncontradas)" Cyan
        Show-Text "🗑️ Filas limpas: $($estatisticas.FilasLimpas)" Cyan
        Show-Text "🗑️ Drivers removidos: $($estatisticas.DriversRemovidos)" Cyan
        Show-Text "💾 Backup salvo em: $backupFile" Cyan
        Separator
        
        if ($resetSucesso) {
            Show-Text "✅ Reset total concluído com sucesso!" Green
        } else {
            Show-Text "⚠️ Reset concluído com alguns avisos. Verifique os logs." Yellow
        }
        
        Verificar-ServicoSpooler
    }
    catch {
        Write-Progress -Activity "Reset total" -Completed
        Show-Text "❌ Erro crítico no reset total: $_" Red
        
        # Tentar restaurar o serviço spooler
        try {
            Show-Text "🔄 Tentando restaurar serviço spooler..." Yellow
            Start-Service spooler -ErrorAction SilentlyContinue
        }
        catch {
            Show-Text "❌ Falha crítica! Reinicialização do sistema pode ser necessária." Red
        }
    }

    Pause
}

function Menu-WinReset {
    Testar-Admin
    
    # Inicializar comandos de reset
    Inicializar-ComandosReset
    
    # Mostrar informações iniciais
    Clear-Host
    Show-Text "🖨️ WinReset v3.0 Universal - Inicializando..." Cyan
    Show-Text "📁 Log será salvo em: $global:logFile" Cyan
    Show-Text "💾 Backups serão salvos em: $global:backupPath" Cyan
    Show-Text "🌐 Suporte: $($Global:SupportedBrands -join ', ')" Yellow
    Start-Sleep -Seconds 2

    do {
        Clear-Host
        Separator
        Show-Text "🖨️ WINRESET v3.0 - Ferramenta Universal de Reset de Impressoras" Magenta
        Show-Text "   Reset Universal: USB • Rede • Wi-Fi • Todas as Marcas" White
        Separator

        # Status do sistema
        $spoolerOk = Verificar-ServicoSpooler
        $sistemaOk = Verificar-Saude-Sistema
        
        if ($spoolerOk -and $sistemaOk) {
            Show-Text "🟢 Sistema de impressão: Saudável" Green
        } elseif ($spoolerOk) {
            Show-Text "🟡 Sistema de impressão: Funcionando com avisos" Yellow
        } else {
            Show-Text "🔴 Sistema de impressão: Problemas detectados" Red
        }
        
        Separator
        Show-Text "🔍 DETECÇÃO UNIVERSAL:" Cyan
        Show-Text "[1] 🖨️ Detectar impressoras locais/USB"
        Show-Text "[2] 🌐 Detectar impressoras na rede"
        Show-Text "[3] 📊 Detectar todas (locais + rede)"
        
        Separator
        Show-Text "📋 LISTAGEM TRADICIONAL:" Cyan
        Show-Text "[4] 📄 Listar impressoras instaladas (resumo)"
        Show-Text "[5] 📊 Listar impressoras instaladas (detalhado)"
        
        Separator
        Show-Text "🔧 RESET UNIVERSAL:" Yellow
        Show-Text "[6] 🎯 Reset de impressora específica (Universal)"
        Show-Text "[7] 🔁 Resetar impressora tradicional"
        Show-Text "[8] ♻️ Reset total do sistema de impressão"
        
        Separator
        Show-Text "🛠️ DIAGNÓSTICO AVANÇADO:" Magenta
        Show-Text "[9] 🔍 Diagnóstico universal de impressora"
        Show-Text "[10] 📊 Diagnóstico completo do sistema"
        Show-Text "[11] 🗂️ Gerenciar backups"
        Show-Text "[12] 📝 Visualizar logs"
        Show-Text "[13] ⚙️ Configurações"
        
        Separator
        Show-Text "[0] ❌ Sair" Red
        Separator

        $op = Read-Host "`n🎯 Escolha uma opção"
        switch ($op) {
            '1' { 
                Detectar-TodasImpressoras
                Pause
            }
            '2' { 
                $range = Read-Host "Digite a faixa de rede (ex: 192.168.1) [Enter para 192.168.1]"
                if ([string]::IsNullOrWhiteSpace($range)) { $range = "192.168.1" }
                Detectar-ImpressorasRede -NetworkRange $range
                Pause
            }
            '3' { 
                $range = Read-Host "Digite a faixa de rede (ex: 192.168.1) [Enter para 192.168.1]"
                if ([string]::IsNullOrWhiteSpace($range)) { $range = "192.168.1" }
                Detectar-TodasImpressoras -IncluirRede -NetworkRange $range
                Pause
            }
            '4' { 
                Listar-Impressoras
                Pause
            }
            '5' { 
                Listar-Impressoras -Detalhado
                Pause
            }
            '6' { Menu-ResetUniversal }
            '7' { Resetar-Impressora-Bruta }
            '8' { Resetar-Tudo-Bruto }
            '9' { Menu-DiagnosticoUniversal }
            '10' { Executar-Diagnostico }
            '11' { Gerenciar-Backups }
            '12' { Visualizar-Logs }
            '13' { Menu-Configuracoes }
            '0' { 
                Show-Text "`n👋 Obrigado por usar o WinReset v3.0 Universal!" Green
                Show-Text "📁 Logs salvos em: $global:logFile" Cyan
                if (Test-Path $global:backupPath) {
                    Show-Text "💾 Backups disponíveis em: $global:backupPath" Cyan
                }
                Pause "Pressione ENTER para sair..."
                break 
            }
            default { 
                Show-Text "❌ Opção inválida. Tente novamente." Red
                Pause
            }
        }
    } while ($true)
}

# Menu para Reset Universal
function Menu-ResetUniversal {
    if ($Global:DetectedPrinters.Count -eq 0) {
        Show-Text "⚠️ Nenhuma impressora detectada. Execute a detecção primeiro." Yellow
        Pause
        return
    }
    
    Clear-Host
    Separator
    Show-Text "🎯 RESET UNIVERSAL DE IMPRESSORAS" Yellow
    Separator
    
    Show-Text "📋 Impressoras detectadas:"
    for ($i = 0; $i -lt $Global:DetectedPrinters.Count; $i++) {
        $printer = $Global:DetectedPrinters[$i]
        $name = $printer.Name ?? $printer.IP
        $type = $printer.Type
        $brand = $printer.Brand
        Show-Text "[$i] $name ($type - $brand)" Cyan
    }
    
    Separator
    Show-Text "[A] 🔄 Reset de todas as impressoras detectadas" Red
    Show-Text "[0] ⬅️ Voltar ao menu principal" White
    Separator
    
    $choice = Read-Host "Escolha uma impressora para reset ou opção"
    
    if ($choice -eq "0") {
        return
    }
    elseif ($choice.ToUpper() -eq "A") {
        if (Confirm-Action "Deseja realmente resetar TODAS as impressoras detectadas?") {
            foreach ($printer in $Global:DetectedPrinters) {
                Reset-ImpressoraUniversal -PrinterInfo $printer
                Start-Sleep -Seconds 2
            }
        }
    }
    elseif ($choice -match '^\d+$' -and [int]$choice -lt $Global:DetectedPrinters.Count) {
        $selectedPrinter = $Global:DetectedPrinters[[int]$choice]
        
        Clear-Host
        Show-Text "🎯 Reset: $($selectedPrinter.Name ?? $selectedPrinter.IP)" Yellow
        Show-Text "Marca: $($selectedPrinter.Brand)" Cyan
        Show-Text "Tipo: $($selectedPrinter.Type)" Cyan
        
        if (Confirm-Action "Confirma o reset desta impressora?") {
            $success = Reset-ImpressoraUniversal -PrinterInfo $selectedPrinter
            if ($success) {
                Show-Text "✅ Reset executado com sucesso!" Green
            } else {
                Show-Text "❌ Falha no reset. Verifique os logs." Red
            }
        }
    }
    else {
        Show-Text "❌ Opção inválida." Red
    }
    
    Pause
}

# Menu para Diagnóstico Universal
function Menu-DiagnosticoUniversal {
    if ($Global:DetectedPrinters.Count -eq 0) {
        Show-Text "⚠️ Nenhuma impressora detectada. Execute a detecção primeiro." Yellow
        Pause
        return
    }
    
    Clear-Host
    Separator
    Show-Text "🔍 DIAGNÓSTICO UNIVERSAL DE IMPRESSORAS" Cyan
    Separator
    
    Show-Text "📋 Impressoras detectadas:"
    for ($i = 0; $i -lt $Global:DetectedPrinters.Count; $i++) {
        $printer = $Global:DetectedPrinters[$i]
        $name = $printer.Name ?? $printer.IP
        $type = $printer.Type
        $brand = $printer.Brand
        Show-Text "[$i] $name ($type - $brand)" Cyan
    }
    
    Separator
    Show-Text "[A] 🔍 Diagnóstico de todas as impressoras" Magenta
    Show-Text "[0] ⬅️ Voltar ao menu principal" White
    Separator
    
    $choice = Read-Host "Escolha uma impressora para diagnóstico ou opção"
    
    if ($choice -eq "0") {
        return
    }
    elseif ($choice.ToUpper() -eq "A") {
        foreach ($printer in $Global:DetectedPrinters) {
            Diagnosticar-ImpressoraUniversal -PrinterInfo $printer
            Show-Text "`n" + ("═" * 60) Yellow
            Start-Sleep -Seconds 2
        }
    }
    elseif ($choice -match '^\d+$' -and [int]$choice -lt $Global:DetectedPrinters.Count) {
        $selectedPrinter = $Global:DetectedPrinters[[int]$choice]
        Diagnosticar-ImpressoraUniversal -PrinterInfo $selectedPrinter
    }
    else {
        Show-Text "❌ Opção inválida." Red
    }
    
    Pause
}

function Executar-Diagnostico {
    Clear-Host
    Separator
    Show-Text "🔍 DIAGNÓSTICO COMPLETO DO SISTEMA" Cyan
    Separator
    
    Show-Progress "Diagnóstico" "Verificando serviços..." 20
    Verificar-ServicoSpooler | Out-Null
    
    Show-Progress "Diagnóstico" "Analisando saúde do sistema..." 40
    Verificar-Saude-Sistema | Out-Null
    
    Show-Progress "Diagnóstico" "Coletando informações detalhadas..." 60
    
    try {
        $impressoras = Get-Printer -ErrorAction SilentlyContinue
        $drivers = Get-PrinterDriver -ErrorAction SilentlyContinue
        $jobs = Get-CimInstance -ClassName Win32_PrintJob -ErrorAction SilentlyContinue
        
        Show-Progress "Diagnóstico" "Gerando relatório..." 80
        
        Clear-Host
        Separator
        Show-Text "📊 RELATÓRIO DE DIAGNÓSTICO" Green
        Separator
        Show-Text "📄 Impressoras instaladas: $(if($impressoras){$impressoras.Count}else{0})" White
        Show-Text "🖨️ Drivers instalados: $(if($drivers){$drivers.Count}else{0})" White
        Show-Text "📋 Jobs na fila: $(if($jobs){$jobs.Count}else{0})" White
        
        if ($impressoras) {
            Show-Text "`n📄 DETALHES DAS IMPRESSORAS:" Cyan
            foreach ($imp in $impressoras) {
                $statusColor = switch ($imp.PrinterStatus) {
                    "Normal" { "Green" }
                    "Error" { "Red" }
                    "Offline" { "Yellow" }
                    default { "White" }
                }
                Show-Text "  • $($imp.Name) - Status: $($imp.PrinterStatus)" $statusColor
            }
        }
        
        Show-Progress "Diagnóstico" "Concluído" 100
        Write-Progress -Activity "Diagnóstico" -Completed
    }
    catch {
        Write-Progress -Activity "Diagnóstico" -Completed
        Show-Text "❌ Erro durante o diagnóstico: $_" Red
    }
    
    Pause
}

function Gerenciar-Backups {
    Clear-Host
    Separator
    Show-Text "🗂️ GERENCIADOR DE BACKUPS" Cyan
    Separator
    
    if (-not (Test-Path $global:backupPath)) {
        Show-Text "📁 Nenhum backup encontrado." Yellow
        Pause
        return
    }
    
    $backups = Get-ChildItem $global:backupPath -Filter "*.json" | Sort-Object LastWriteTime -Descending
    
    if ($backups.Count -eq 0) {
        Show-Text "📁 Nenhum arquivo de backup encontrado." Yellow
        Pause
        return
    }
    
    Show-Text "📋 Backups disponíveis:" Green
    for ($i = 0; $i -lt $backups.Count; $i++) {
        $backup = $backups[$i]
        $size = [math]::Round($backup.Length / 1KB, 2)
        Show-Text "[$i] $($backup.Name) - $($backup.LastWriteTime) - $size KB" White
    }
    
    Show-Text "`n[V] Ver conteúdo de um backup"
    Show-Text "[L] Limpar backups antigos"
    Show-Text "[0] Voltar"
    
    $opcao = Read-Host "`nEscolha uma opção"
    
    switch ($opcao.ToUpper()) {
        'V' {
            $index = Read-Host "Digite o número do backup para visualizar"
            if ($index -match '^\d+$' -and [int]$index -lt $backups.Count) {
                $conteudo = Get-Content $backups[$index].FullName | ConvertFrom-Json
                Show-Text "`n📄 Conteúdo do backup:" Cyan
                Show-Text "Data: $($conteudo.Timestamp)" White
                Show-Text "Tipo: $($conteudo.Tipo)" White
                Show-Text "Impressoras: $($conteudo.Impressoras.Count)" White
                Show-Text "Drivers: $($conteudo.Drivers.Count)" White
            }
        }
        'L' {
            if (Confirm-Action "Deseja remover backups com mais de 7 dias?") {
                $limite = (Get-Date).AddDays(-7)
                $removidos = 0
                foreach ($backup in $backups) {
                    if ($backup.LastWriteTime -lt $limite) {
                        Remove-Item $backup.FullName -Force
                        $removidos++
                    }
                }
                Show-Text "🗑️ $removidos backups antigos removidos." Green
            }
        }
    }
    
    Pause
}

function Visualizar-Logs {
    Clear-Host
    Separator
    Show-Text "📝 VISUALIZADOR DE LOGS" Cyan
    Separator
    
    if (-not (Test-Path $global:logFile)) {
        Show-Text "📄 Arquivo de log não encontrado." Yellow
        Pause
        return
    }
    
    $linhas = Get-Content $global:logFile -Tail 50
    Show-Text "📋 Últimas 50 linhas do log:" Green
    Show-Text ""
    
    foreach ($linha in $linhas) {
        $cor = "White"
        if ($linha -match "\[ERROR\]") { $cor = "Red" }
        elseif ($linha -match "\[WARNING\]") { $cor = "Yellow" }
        elseif ($linha -match "\[SUCCESS\]") { $cor = "Green" }
        
        Show-Text $linha $cor -NoLog
    }
    
    Pause
}

function Menu-Configuracoes {
    Clear-Host
    Separator
    Show-Text "⚙️ CONFIGURAÇÕES" Cyan
    Separator
    
    Show-Text "[1] 🔊 Alternar modo verboso: $(if($global:verboseMode){'Ativado'}else{'Desativado'})" White
    Show-Text "[2] 📁 Abrir pasta de logs"
    Show-Text "[3] 💾 Abrir pasta de backups"
    Show-Text "[4] 🔄 Reiniciar serviço spooler"
    Show-Text "[0] Voltar"
    
    $opcao = Read-Host "`nEscolha uma opção"
    
    switch ($opcao) {
        '1' {
            $global:verboseMode = -not $global:verboseMode
            Show-Text "🔊 Modo verboso: $(if($global:verboseMode){'Ativado'}else{'Desativado'})" Green
            Start-Sleep -Seconds 1
        }
        '2' {
            if (Test-Path (Split-Path $global:logFile)) {
                Start-Process "explorer.exe" -ArgumentList (Split-Path $global:logFile)
            }
        }
        '3' {
            if (Test-Path $global:backupPath) {
                Start-Process "explorer.exe" -ArgumentList $global:backupPath
            } else {
                Show-Text "📁 Pasta de backup não existe ainda." Yellow
                Start-Sleep -Seconds 2
            }
        }
        '4' {
            if (Confirm-Action "Deseja reiniciar o serviço spooler?") {
                try {
                    Restart-Service spooler -Force
                    Show-Text "✅ Serviço spooler reiniciado." Green
                } catch {
                    Show-Text "❌ Erro ao reiniciar: $_" Red
                }
                Start-Sleep -Seconds 2
            }
        }
    }
}

# Executa o menu
Menu-WinReset
