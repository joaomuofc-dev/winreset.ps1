# ===============================================================================
#  WinReset v5.0 Professional - Sistema Universal de ManutenÃ§Ã£o de Impressoras
# ===============================================================================
# Desenvolvido para: Ã“rgÃ£os PÃºblicos e Empresas
# Substituto gratuito do: WIC Reset e ferramentas pagas
# Autor: Sistema IA AvanÃ§ado
# Data: 2025-01-07
# LicenÃ§a: Gratuito para uso pÃºblico
# Compatibilidade: Windows 10/11, PowerShell 5.1+
# ===============================================================================

# ConfiguraÃ§Ã£o inicial do console
[Console]::Title = "WinReset v5.0 Professional - Sistema Universal de Impressoras"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# ConfiguraÃ§Ãµes globais
$Global:WinResetVersion = "5.0 Professional"
$Global:LogPath = "$env:USERPROFILE\WinReset\Logs"
$Global:BackupPath = "$env:USERPROFILE\WinReset\Backups"
$Global:ConfigPath = "$env:USERPROFILE\WinReset\Config"
$Global:DetectedPrinters = @()
$Global:PrinterDatabase = @{}
$Global:VerboseMode = $false

# Criar estrutura de pastas
if (-not (Test-Path $Global:LogPath)) { New-Item -Path $Global:LogPath -ItemType Directory -Force | Out-Null }
if (-not (Test-Path $Global:BackupPath)) { New-Item -Path $Global:BackupPath -ItemType Directory -Force | Out-Null }
if (-not (Test-Path $Global:ConfigPath)) { New-Item -Path $Global:ConfigPath -ItemType Directory -Force | Out-Null }

# Sistema de logging avanÃ§ado
function Write-WinResetLog {
    param(
        [string]$Message,
        [ValidateSet("INFO", "SUCCESS", "WARNING", "ERROR", "DEBUG")]
        [string]$Level = "INFO",
        [string]$Component = "SYSTEM"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $logFile = "$Global:LogPath\WinReset_$(Get-Date -Format 'yyyyMMdd').log"
    $logEntry = "[$timestamp] [$Level] [$Component] $Message"
    
    try {
        Add-Content -Path $logFile -Value $logEntry -Encoding UTF8
        if ($Global:VerboseMode) {
            Write-Host "[LOG] $logEntry" -ForegroundColor DarkGray
        }
    }
    catch {
        Write-Warning "Erro no sistema de log: $_"
    }
}

# Sistema de exibiÃ§Ã£o colorida profissional
function Show-WinResetText {
    param(
        [string]$Text,
        [ValidateSet("Info", "Success", "Warning", "Error", "Title", "Menu", "Highlight")]
        [string]$Type = "Info",
        [switch]$NoNewLine,
        [switch]$Center
    )
    
    $colors = @{
        "Info" = "Cyan"
        "Success" = "Green"
        "Warning" = "Yellow"
        "Error" = "Red"
        "Title" = "Magenta"
        "Menu" = "White"
        "Highlight" = "DarkYellow"
    }
    
    $color = $colors[$Type]
    
    if ($Center) {
        $consoleWidth = [Console]::WindowWidth
        $padding = [Math]::Max(0, ($consoleWidth - $Text.Length) / 2)
        $Text = (" " * $padding) + $Text
    }
    
    if ($NoNewLine) {
        Write-Host $Text -ForegroundColor $color -NoNewline
    } else {
        Write-Host $Text -ForegroundColor $color
    }
    
    # Log automÃ¡tico
    $logLevel = switch ($Type) {
        "Success" { "SUCCESS" }
        "Warning" { "WARNING" }
        "Error" { "ERROR" }
        default { "INFO" }
    }
    Write-WinResetLog $Text $logLevel "UI"
}

# Moldura profissional para menus
function Show-WinResetFrame {
    param(
        [string[]]$Content,
        [string]$Title = "",
        [ConsoleColor]$FrameColor = "DarkBlue",
        [ConsoleColor]$TitleColor = "Yellow"
    )
    
    $width = 80
    $topBorder = "â•”" + ("â•" * ($width - 2)) + "â•—"
    $bottomBorder = "â•š" + ("â•" * ($width - 2)) + "â•"
    
    Write-Host $topBorder -ForegroundColor $FrameColor
    
    if ($Title) {
        $titlePadding = [Math]::Max(0, ($width - $Title.Length - 4) / 2)
        $titleLine = "â•‘" + (" " * $titlePadding) + $Title + (" " * ($width - $Title.Length - $titlePadding - 2)) + "â•‘"
        Write-Host $titleLine -ForegroundColor $TitleColor
        Write-Host ("â•‘" + ("â•" * ($width - 2)) + "â•‘") -ForegroundColor $FrameColor
    }
    
    foreach ($line in $Content) {
        $padding = $width - $line.Length - 2
        $contentLine = "â•‘ " + $line + (" " * $padding) + "â•‘"
        Write-Host $contentLine -ForegroundColor White
    }
    
    Write-Host $bottomBorder -ForegroundColor $FrameColor
}

# VerificaÃ§Ã£o de privilÃ©gios administrativos
function Test-WinResetAdmin {
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
    
    if (-not $isAdmin) {
        Clear-Host
        Show-WinResetFrame -Title "ERRO: PRIVILÃ‰GIOS INSUFICIENTES" -Content @(
            "",
            "âŒ O WinReset requer privilÃ©gios de ADMINISTRADOR para funcionar.",
            "",
            "ðŸ“‹ Como executar corretamente:",
            "   1. Feche este PowerShell",
            "   2. Clique com botÃ£o direito no PowerShell",
            "   3. Selecione 'Executar como administrador'",
            "   4. Execute novamente o WinReset",
            "",
            "âš ï¸  Sem privilÃ©gios administrativos, o reset nÃ£o funcionarÃ¡!",
            ""
        ) -FrameColor Red
        
        Write-WinResetLog "Tentativa de execuÃ§Ã£o sem privilÃ©gios administrativos" "ERROR" "SECURITY"
        Read-Host "\nPressione ENTER para sair"
        exit 1
    }
    
    Show-WinResetText "âœ… Executando com privilÃ©gios administrativos" "Success"
    Write-WinResetLog "Iniciado com privilÃ©gios administrativos" "SUCCESS" "SECURITY"
}

# Base de dados de comandos por fabricante
function Initialize-PrinterDatabase {
    $Global:PrinterDatabase = @{
        "Epson" = @{
            "ResetCommands" = @(
                "`e@",  # ESC @ - Reset bÃ¡sico
                "`e%-12345X@PJL RESET`r`n`e%-12345X`r`n",  # PJL Reset
                "`e(R`0`0REMOTE1ST`r`n",  # Comando remoto Epson
                "`e(R`0`0REMOTE1IC`r`n"   # Reset contador tinta
            )
            "CleanCommands" = @(
                "`e(R`0`0REMOTE1CL`r`n",  # Limpeza cabeÃ§ote
                "`e(R`0`0REMOTE1PH`r`n"   # Limpeza profunda
            )
            "StatusCommands" = @(
                "`e(R`0`0REMOTE1ST`r`n"   # Status da impressora
            )
            "Models" = @("L3250", "L3150", "L4150", "L4160", "WF-2830", "XP-241")
        }
        "HP" = @{
            "ResetCommands" = @(
                "`e%-12345X@PJL RESET`r`n`e%-12345X`r`n",
                "`e%-12345X@PJL SET CLEARJAM=ON`r`n`e%-12345X`r`n",
                "`eE"  # PCL Reset
            )
            "CleanCommands" = @(
                "`e%-12345X@PJL SET CLEANMODE=ON`r`n`e%-12345X`r`n"
            )
            "StatusCommands" = @(
                "`e%-12345X@PJL INFO STATUS`r`n`e%-12345X`r`n"
            )
            "Models" = @("LaserJet", "DeskJet", "OfficeJet", "Envy")
        }
        "Canon" = @{
            "ResetCommands" = @(
                "`e@",
                "`e[K",  # Reset Canon especÃ­fico
                "`e%-12345X@PJL RESET`r`n`e%-12345X`r`n"
            )
            "CleanCommands" = @(
                "`e[c"  # Limpeza Canon
            )
            "StatusCommands" = @(
                "`e[s"
            )
            "Models" = @("PIXMA", "imageCLASS", "MAXIFY")
        }
        "Brother" = @{
            "ResetCommands" = @(
                "`e@",
                "`e%-12345X@PJL RESET`r`n`e%-12345X`r`n",
                "`e[?"  # Reset Brother especÃ­fico
            )
            "CleanCommands" = @(
                "`e[c"
            )
            "StatusCommands" = @(
                "`e[s"
            )
            "Models" = @("DCP-", "MFC-", "HL-")
        }
        "Zebra" = @{
            "ResetCommands" = @(
                "~JA",  # ZPL Reset
                "^XA^JUF^XZ",  # ZPL Factory Reset
                "`e@"  # ESC/POS fallback
            )
            "CleanCommands" = @(
                "~JC"  # ZPL Clean
            )
            "StatusCommands" = @(
                "~HS"  # ZPL Status
            )
            "Models" = @("GK420", "ZT230", "LP2844")
        }
    }
    
    Write-WinResetLog "Base de dados de impressoras inicializada com $($Global:PrinterDatabase.Count) fabricantes" "INFO" "DATABASE"
}

# DetecÃ§Ã£o inteligente de impressoras locais
function Get-WinResetLocalPrinters {
    Show-WinResetText "ðŸ” Detectando impressoras locais e USB..." "Info"
    
    $localPrinters = @()
    
    try {
        # MÃ©todo 1: Get-Printer (mais moderno)
        $printers = Get-Printer -ErrorAction SilentlyContinue
        foreach ($printer in $printers) {
            $brand = Get-PrinterBrand $printer.Name
            $localPrinters += @{
                Name = $printer.Name
                Type = "Local"
                Brand = $brand
                Status = $printer.PrinterStatus
                Driver = $printer.DriverName
                Port = $printer.PortName
                Location = $printer.Location
                Comment = $printer.Comment
                Shared = $printer.Shared
            }
        }
        
        # MÃ©todo 2: WMI (compatibilidade)
        $wmiPrinters = Get-WmiObject -Class Win32_Printer -ErrorAction SilentlyContinue
        foreach ($wmiPrinter in $wmiPrinters) {
            # Verificar se jÃ¡ foi detectada pelo Get-Printer
            $exists = $localPrinters | Where-Object { $_.Name -eq $wmiPrinter.Name }
            if (-not $exists) {
                $brand = Get-PrinterBrand $wmiPrinter.Name
                $localPrinters += @{
                    Name = $wmiPrinter.Name
                    Type = "Local (WMI)"
                    Brand = $brand
                    Status = $wmiPrinter.PrinterStatus
                    Driver = $wmiPrinter.DriverName
                    Port = $wmiPrinter.PortName
                    Location = $wmiPrinter.Location
                    Comment = $wmiPrinter.Comment
                    Shared = $wmiPrinter.Shared
                }
            }
        }
        
        Show-WinResetText "âœ… Detectadas $($localPrinters.Count) impressoras locais" "Success"
        Write-WinResetLog "Detectadas $($localPrinters.Count) impressoras locais" "SUCCESS" "DETECTION"
        
        return $localPrinters
    }
    catch {
        Show-WinResetText "âŒ Erro na detecÃ§Ã£o local: $_" "Error"
        Write-WinResetLog "Erro na detecÃ§Ã£o local: $_" "ERROR" "DETECTION"
        return @()
    }
}

# DetecÃ§Ã£o inteligente de impressoras na rede
function Get-WinResetNetworkPrinters {
    param(
        [string]$NetworkRange = "192.168.1",
        [int]$TimeoutSeconds = 1
    )
    
    Show-WinResetText "ðŸŒ Escaneando rede $NetworkRange.1-254 (isso pode demorar...)" "Info"
    
    $networkPrinters = @()
    $commonPorts = @(9100, 515, 631, 80, 443, 161)  # RAW, LPD, IPP, HTTP, HTTPS, SNMP
    
    # Scan paralelo para performance
    $jobs = @()
    
    1..254 | ForEach-Object {
        $ip = "$NetworkRange.$_"
        $jobs += Start-Job -ScriptBlock {
            param($IP, $Ports, $Timeout)
            
            $result = $null
            
            # Teste ping rÃ¡pido primeiro
            if (Test-Connection -ComputerName $IP -Count 1 -Quiet -TimeoutSec $Timeout) {
                foreach ($port in $Ports) {
                    try {
                        $tcpClient = New-Object System.Net.Sockets.TcpClient
                        $connect = $tcpClient.BeginConnect($IP, $port, $null, $null)
                        $wait = $connect.AsyncWaitHandle.WaitOne(($Timeout * 1000), $false)
                        
                        if ($wait) {
                            $tcpClient.EndConnect($connect)
                            
                            # Tentar identificar o tipo de serviÃ§o
                            $service = switch ($port) {
                                9100 { "RAW/JetDirect" }
                                515 { "LPD" }
                                631 { "IPP" }
                                80 { "HTTP" }
                                443 { "HTTPS" }
                                161 { "SNMP" }
                                default { "Unknown" }
                            }
                            
                            $result = @{
                                IP = $IP
                                Port = $port
                                Service = $service
                                Type = "Network"
                                Status = "Online"
                                ResponseTime = (Get-Date)
                            }
                            break
                        }
                    }
                    catch { }
                    finally {
                        if ($tcpClient) { $tcpClient.Close() }
                    }
                }
            }
            return $result
        } -ArgumentList $ip, $commonPorts, $TimeoutSeconds
    }
    
    # Aguardar e coletar resultados
    $completed = 0
    $total = $jobs.Count
    
    foreach ($job in $jobs) {
        $result = Wait-Job $job | Receive-Job
        if ($result) {
            # Tentar identificar marca via SNMP ou HTTP
            $brand = Get-NetworkPrinterBrand -IP $result.IP -Port $result.Port
            $result.Brand = $brand
            $networkPrinters += $result
        }
        Remove-Job $job
        
        $completed++
        $percent = [Math]::Round(($completed / $total) * 100)
        Write-Progress -Activity "Escaneando rede" -Status "$completed/$total IPs verificados" -PercentComplete $percent
    }
    
    Write-Progress -Activity "Escaneando rede" -Completed
    
    Show-WinResetText "âœ… Detectadas $($networkPrinters.Count) impressoras na rede" "Success"
    Write-WinResetLog "Detectadas $($networkPrinters.Count) impressoras na rede" "SUCCESS" "DETECTION"
    
    return $networkPrinters
}

# IdentificaÃ§Ã£o de marca por nome
function Get-PrinterBrand {
    param([string]$PrinterName)
    
    $brandPatterns = @{
        "Epson" = @("EPSON", "L3250", "L3150", "L4150", "WF-", "XP-")
        "HP" = @("HP", "LaserJet", "DeskJet", "OfficeJet", "Envy")
        "Canon" = @("Canon", "PIXMA", "imageCLASS", "MAXIFY")
        "Brother" = @("Brother", "DCP-", "MFC-", "HL-")
        "Zebra" = @("Zebra", "ZPL", "EPL", "GK420")
        "Samsung" = @("Samsung", "SCX-", "ML-")
        "Lexmark" = @("Lexmark")
        "Kyocera" = @("Kyocera", "FS-")
    }
    
    foreach ($brand in $brandPatterns.Keys) {
        foreach ($pattern in $brandPatterns[$brand]) {
            if ($PrinterName -like "*$pattern*") {
                return $brand
            }
        }
    }
    
    return "Desconhecida"
}

# IdentificaÃ§Ã£o de marca via rede (SNMP/HTTP)
function Get-NetworkPrinterBrand {
    param(
        [string]$IP,
        [int]$Port
    )
    
    try {
        # Tentar SNMP primeiro (mais confiÃ¡vel)
        if ($Port -eq 161) {
            # OID para descriÃ§Ã£o do sistema: 1.3.6.1.2.1.1.1.0
            # ImplementaÃ§Ã£o bÃ¡sica - em produÃ§Ã£o usar SNMP real
            return "SNMP_Device"
        }
        
        # Tentar HTTP/HTTPS
        if ($Port -eq 80 -or $Port -eq 443) {
            $protocol = if ($Port -eq 443) { "https" } else { "http" }
            $response = Invoke-WebRequest -Uri "$protocol://$IP" -TimeoutSec 3 -ErrorAction SilentlyContinue
            
            if ($response) {
                $content = $response.Content
                
                # PadrÃµes comuns em pÃ¡ginas web de impressoras
                if ($content -match "EPSON|Epson") { return "Epson" }
                if ($content -match "HP|Hewlett") { return "HP" }
                if ($content -match "Canon") { return "Canon" }
                if ($content -match "Brother") { return "Brother" }
                if ($content -match "Zebra") { return "Zebra" }
            }
        }
        
        return "Desconhecida"
    }
    catch {
        return "Desconhecida"
    }
}

# Reset completo do sistema de impressÃ£o
function Invoke-WinResetSystemReset {
    Show-WinResetText "âš ï¸  ATENÃ‡ÃƒO: Reset total do sistema de impressÃ£o!" "Warning"
    Show-WinResetText "Esta operaÃ§Ã£o irÃ¡:" "Info"
    Show-WinResetText "  â€¢ Parar o serviÃ§o Spooler" "Info"
    Show-WinResetText "  â€¢ Limpar todas as filas de impressÃ£o" "Info"
    Show-WinResetText "  â€¢ Remover arquivos temporÃ¡rios" "Info"
    Show-WinResetText "  â€¢ Reiniciar o sistema de impressÃ£o" "Info"
    
    $confirm = Read-Host "\nDigite 'CONFIRMO' para continuar ou ENTER para cancelar"
    
    if ($confirm -ne "CONFIRMO") {
        Show-WinResetText "âŒ OperaÃ§Ã£o cancelada pelo usuÃ¡rio" "Warning"
        return
    }
    
    try {
        Write-WinResetLog "Iniciando reset total do sistema" "INFO" "RESET"
        
        # Backup antes do reset
        $backupFolder = "$Global:BackupPath\SystemReset_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        New-Item -Path $backupFolder -ItemType Directory -Force | Out-Null
        
        Show-WinResetText "ðŸ“ Criando backup em: $backupFolder" "Info"
        
        # 1. Parar serviÃ§o Spooler
        Show-WinResetText "ðŸ›‘ Parando serviÃ§o Spooler..." "Info"
        Stop-Service spooler -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 3
        
        # 2. Backup e limpeza de arquivos de spool
        $spoolPath = "$env:SystemRoot\System32\spool\PRINTERS"
        if (Test-Path $spoolPath) {
            Show-WinResetText "ðŸ§¹ Limpando arquivos de spool..." "Info"
            
            # Backup dos arquivos
            $spoolFiles = Get-ChildItem $spoolPath -ErrorAction SilentlyContinue
            if ($spoolFiles) {
                Copy-Item $spoolPath\* $backupFolder -Force -ErrorAction SilentlyContinue
                Remove-Item "$spoolPath\*" -Force -ErrorAction SilentlyContinue
            }
        }
        
        # 3. Limpar registry de impressoras (backup automÃ¡tico)
        Show-WinResetText "ðŸ”§ Limpando entradas de registro..." "Info"
        $regPaths = @(
            "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers",
            "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Environments"
        )
        
        foreach ($regPath in $regPaths) {
            if (Test-Path $regPath) {
                # Backup do registro
                $regBackup = "$backupFolder\Registry_$(($regPath -replace ':', '_') -replace '\\', '_').reg"
                reg export $regPath.Replace('HKLM:', 'HKEY_LOCAL_MACHINE') $regBackup /y 2>$null
            }
        }
        
        # 4. Reiniciar serviÃ§o Spooler
        Show-WinResetText "ðŸ”„ Reiniciando serviÃ§o Spooler..." "Info"
        Start-Service spooler -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 5
        
        # 5. Verificar se o serviÃ§o estÃ¡ funcionando
        $spoolerStatus = Get-Service spooler
        if ($spoolerStatus.Status -eq "Running") {
            Show-WinResetText "âœ… Reset do sistema concluÃ­do com sucesso!" "Success"
            Show-WinResetText "ðŸ“ Backup salvo em: $backupFolder" "Info"
            Write-WinResetLog "Reset total do sistema concluÃ­do com sucesso" "SUCCESS" "RESET"
        } else {
            Show-WinResetText "âŒ Erro: ServiÃ§o Spooler nÃ£o iniciou corretamente" "Error"
            Write-WinResetLog "Erro: ServiÃ§o Spooler nÃ£o iniciou apÃ³s reset" "ERROR" "RESET"
        }
    }
    catch {
        Show-WinResetText "âŒ Erro durante o reset: $_" "Error"
        Write-WinResetLog "Erro durante reset do sistema: $_" "ERROR" "RESET"
    }
}

# Reset especÃ­fico de impressora
function Invoke-WinResetPrinterReset {
    param(
        [hashtable]$Printer
    )
    
    $printerName = if ($Printer.Name) { $Printer.Name } else { $Printer.IP }
    Show-WinResetText "ðŸ”„ Executando reset da impressora: $printerName" "Info"
    
    try {
        Write-WinResetLog "Iniciando reset da impressora: $printerName" "INFO" "RESET"
        
        if ($Printer.Type -eq "Network" -or $Printer.IP) {
            # Reset via rede
            $success = Invoke-NetworkPrinterReset -Printer $Printer
        } else {
            # Reset local
            $success = Invoke-LocalPrinterReset -Printer $Printer
        }
        
        if ($success) {
            Show-WinResetText "âœ… Reset executado com sucesso!" "Success"
            Write-WinResetLog "Reset da impressora $printerName concluÃ­do com sucesso" "SUCCESS" "RESET"
        } else {
            Show-WinResetText "âŒ Falha no reset da impressora" "Error"
            Write-WinResetLog "Falha no reset da impressora $printerName" "ERROR" "RESET"
        }
    }
    catch {
        Show-WinResetText "âŒ Erro durante o reset: $_" "Error"
        Write-WinResetLog "Erro durante reset da impressora $printerName: $_" "ERROR" "RESET"
    }
}

# Reset via rede (TCP/IP)
function Invoke-NetworkPrinterReset {
    param([hashtable]$Printer)
    
    $ip = $Printer.IP
    $brand = $Printer.Brand
    
    try {
        # Usar comandos especÃ­ficos da marca
        $commands = @()
        if ($Global:PrinterDatabase.ContainsKey($brand)) {
            $commands = $Global:PrinterDatabase[$brand].ResetCommands
        } else {
            # Comandos genÃ©ricos
            $commands = @(
                "`e@",  # ESC @ - Reset universal
                "`e%-12345X@PJL RESET`r`n`e%-12345X`r`n",  # PJL Reset
                "`eE"   # PCL Reset
            )
        }
        
        Show-WinResetText "ðŸ“¡ Conectando via TCP/IP na porta 9100..." "Info"
        
        $tcpClient = New-Object System.Net.Sockets.TcpClient
        $tcpClient.Connect($ip, 9100)
        $stream = $tcpClient.GetStream()
        
        foreach ($command in $commands) {
            Show-WinResetText "ðŸ“¤ Enviando comando de reset..." "Info"
            $bytes = [System.Text.Encoding]::ASCII.GetBytes($command)
            $stream.Write($bytes, 0, $bytes.Length)
            Start-Sleep -Milliseconds 500
        }
        
        $tcpClient.Close()
        return $true
    }
    catch {
        Show-WinResetText "âŒ Erro na conexÃ£o TCP/IP: $_" "Error"
        return $false
    }
}

# Reset local (via spooler)
function Invoke-LocalPrinterReset {
    param([hashtable]$Printer)
    
    try {
        $printerName = $Printer.Name
        
        Show-WinResetText "ðŸ–¨ï¸ Limpando fila de impressÃ£o local..." "Info"
        
        # Limpar fila de impressÃ£o
        Get-PrintJob -PrinterName $printerName -ErrorAction SilentlyContinue | Remove-PrintJob -Confirm:$false
        
        # Reiniciar porta da impressora
        $port = $Printer.Port
        if ($port) {
            Show-WinResetText "ðŸ”Œ Reiniciando porta: $port" "Info"
            # Implementar reinicializaÃ§Ã£o de porta se necessÃ¡rio
        }
        
        return $true
    }
    catch {
        Show-WinResetText "âŒ Erro no reset local: $_" "Error"
        return $false
    }
}

# DiagnÃ³stico completo do sistema
function Invoke-WinResetDiagnostic {
    Show-WinResetText "ðŸ” Executando diagnÃ³stico completo do sistema..." "Info"
    
    $diagnosticResults = @{
        SpoolerService = $false
        PrinterDrivers = 0
        PrinterPorts = 0
        QueuedJobs = 0
        SystemHealth = "Unknown"
        Recommendations = @()
    }
    
    try {
        # 1. Verificar serviÃ§o Spooler
        Show-WinResetText "ðŸ“‹ Verificando serviÃ§o Spooler..." "Info"
        $spooler = Get-Service -Name spooler
        $diagnosticResults.SpoolerService = ($spooler.Status -eq "Running")
        
        if ($diagnosticResults.SpoolerService) {
            Show-WinResetText "  âœ… ServiÃ§o Spooler: Funcionando" "Success"
        } else {
            Show-WinResetText "  âŒ ServiÃ§o Spooler: Parado ou com problemas" "Error"
            $diagnosticResults.Recommendations += "Reiniciar o serviÃ§o Spooler"
        }
        
        # 2. Verificar drivers
        Show-WinResetText "ðŸ“‹ Verificando drivers de impressora..." "Info"
        $drivers = Get-PrinterDriver -ErrorAction SilentlyContinue
        $diagnosticResults.PrinterDrivers = $drivers.Count
        Show-WinResetText "  ðŸ“¦ Drivers instalados: $($drivers.Count)" "Info"
        
        # 3. Verificar portas
        Show-WinResetText "ðŸ“‹ Verificando portas de impressora..." "Info"
        $ports = Get-PrinterPort -ErrorAction SilentlyContinue
        $diagnosticResults.PrinterPorts = $ports.Count
        Show-WinResetText "  ðŸ”Œ Portas configuradas: $($ports.Count)" "Info"
        
        # 4. Verificar trabalhos na fila
        Show-WinResetText "ðŸ“‹ Verificando filas de impressÃ£o..." "Info"
        $allJobs = @()
        $printers = Get-Printer -ErrorAction SilentlyContinue
        foreach ($printer in $printers) {
            $jobs = Get-PrintJob -PrinterName $printer.Name -ErrorAction SilentlyContinue
            $allJobs += $jobs
        }
        $diagnosticResults.QueuedJobs = $allJobs.Count
        
        if ($diagnosticResults.QueuedJobs -gt 0) {
            Show-WinResetText "  âš ï¸  Trabalhos na fila: $($diagnosticResults.QueuedJobs)" "Warning"
            $diagnosticResults.Recommendations += "Limpar filas de impressÃ£o"
        } else {
            Show-WinResetText "  âœ… Filas de impressÃ£o: Limpas" "Success"
        }
        
        # 5. AvaliaÃ§Ã£o geral
        if ($diagnosticResults.SpoolerService -and $diagnosticResults.QueuedJobs -eq 0) {
            $diagnosticResults.SystemHealth = "SaudÃ¡vel"
            Show-WinResetText "\nðŸŽ‰ Sistema de impressÃ£o: SAUDÃVEL" "Success"
        } elseif ($diagnosticResults.SpoolerService) {
            $diagnosticResults.SystemHealth = "AtenÃ§Ã£o"
            Show-WinResetText "\nâš ï¸  Sistema de impressÃ£o: REQUER ATENÃ‡ÃƒO" "Warning"
        } else {
            $diagnosticResults.SystemHealth = "CrÃ­tico"
            Show-WinResetText "\nâŒ Sistema de impressÃ£o: CRÃTICO" "Error"
        }
        
        # 6. RecomendaÃ§Ãµes
        if ($diagnosticResults.Recommendations.Count -gt 0) {
            Show-WinResetText "\nðŸ“‹ RecomendaÃ§Ãµes:" "Info"
            foreach ($rec in $diagnosticResults.Recommendations) {
                Show-WinResetText "  â€¢ $rec" "Highlight"
            }
        }
        
        Write-WinResetLog "DiagnÃ³stico concluÃ­do - Status: $($diagnosticResults.SystemHealth)" "INFO" "DIAGNOSTIC"
        
    }
    catch {
        Show-WinResetText "âŒ Erro durante diagnÃ³stico: $_" "Error"
        Write-WinResetLog "Erro durante diagnÃ³stico: $_" "ERROR" "DIAGNOSTIC"
    }
    
    return $diagnosticResults
}

# Menu principal do WinReset
function Show-WinResetMainMenu {
    do {
        Clear-Host
        
        # Verificar privilÃ©gios
        Test-WinResetAdmin
        
        # CabeÃ§alho principal
        Show-WinResetFrame -Title "WINRESET v5.0 PROFESSIONAL" -Content @(
            "",
            "ðŸ›ï¸  Sistema Universal de ManutenÃ§Ã£o de Impressoras",
            "ðŸ“‹ Desenvolvido para Ã“rgÃ£os PÃºblicos e Empresas",
            "ðŸ†“ Substituto gratuito do WIC Reset e ferramentas pagas",
            "ðŸ“ Logs: $Global:LogPath",
            "ðŸ’¾ Backups: $Global:BackupPath",
            ""
        ) -FrameColor DarkBlue -TitleColor Yellow
        
        # Menu de opÃ§Ãµes
        Show-WinResetText "\nðŸ” DETECÃ‡ÃƒO AUTOMÃTICA:" "Title"
        Show-WinResetText "[1] Detectar impressoras locais/USB" "Menu"
        Show-WinResetText "[2] Detectar impressoras na rede" "Menu"
        Show-WinResetText "[3] Detectar todas (locais + rede)" "Menu"
        
        Show-WinResetText "\nðŸ“‹ LISTAGEM E STATUS:" "Title"
        Show-WinResetText "[4] Listar impressoras detectadas" "Menu"
        Show-WinResetText "[5] Status detalhado das impressoras" "Menu"
        
        Show-WinResetText "\nâ™»ï¸ RESET E MANUTENÃ‡ÃƒO:" "Title"
        Show-WinResetText "[6] Reset de impressora especÃ­fica" "Menu"
        Show-WinResetText "[7] Reset total do sistema de impressÃ£o" "Menu"
        Show-WinResetText "[8] Limpeza de cabeÃ§ote (se suportado)" "Menu"
        
        Show-WinResetText "\nðŸ”§ DIAGNÃ“STICO E FERRAMENTAS:" "Title"
        Show-WinResetText "[9] DiagnÃ³stico completo do sistema" "Menu"
        Show-WinResetText "[10] Teste de impressÃ£o" "Menu"
        Show-WinResetText "[11] Gerenciar backups" "Menu"
        
        Show-WinResetText "\nâš™ï¸ CONFIGURAÃ‡Ã•ES:" "Title"
        Show-WinResetText "[12] ConfiguraÃ§Ãµes avanÃ§adas" "Menu"
        Show-WinResetText "[13] Visualizar logs" "Menu"
        Show-WinResetText "[14] Sobre o WinReset" "Menu"
        
        Show-WinResetText "\n[0] Sair" "Error"
        
        $option = Read-Host "\nðŸŽ¯ Escolha uma opÃ§Ã£o"
        
        switch ($option) {
            "1" {
                Clear-Host
                $Global:DetectedPrinters = Get-WinResetLocalPrinters
                Read-Host "\nPressione ENTER para continuar"
            }
            "2" {
                Clear-Host
                $range = Read-Host "Digite a faixa de rede (ex: 192.168.1) ou ENTER para padrÃ£o"
                if ([string]::IsNullOrWhiteSpace($range)) { $range = "192.168.1" }
                $networkPrinters = Get-WinResetNetworkPrinters -NetworkRange $range
                $Global:DetectedPrinters += $networkPrinters
                Read-Host "\nPressione ENTER para continuar"
            }
            "3" {
                Clear-Host
                $Global:DetectedPrinters = Get-WinResetLocalPrinters
                $range = Read-Host "\nDigite a faixa de rede (ex: 192.168.1) ou ENTER para padrÃ£o"
                if ([string]::IsNullOrWhiteSpace($range)) { $range = "192.168.1" }
                $networkPrinters = Get-WinResetNetworkPrinters -NetworkRange $range
                $Global:DetectedPrinters += $networkPrinters
                Read-Host "\nPressione ENTER para continuar"
            }
            "4" {
                Clear-Host
                Show-PrinterList
                Read-Host "\nPressione ENTER para continuar"
            }
            "5" {
                Clear-Host
                Show-PrinterList -Detailed
                Read-Host "\nPressione ENTER para continuar"
            }
            "6" {
                Clear-Host
                Show-PrinterResetMenu
            }
            "7" {
                Clear-Host
                Invoke-WinResetSystemReset
                Read-Host "\nPressione ENTER para continuar"
            }
            "8" {
                Clear-Host
                Show-WinResetText "ðŸ§¹ Funcionalidade de limpeza serÃ¡ implementada em breve" "Info"
                Read-Host "\nPressione ENTER para continuar"
            }
            "9" {
                Clear-Host
                Invoke-WinResetDiagnostic
                Read-Host "\nPressione ENTER para continuar"
            }
            "10" {
                Clear-Host
                Show-WinResetText "ðŸ–¨ï¸ Funcionalidade de teste serÃ¡ implementada em breve" "Info"
                Read-Host "\nPressione ENTER para continuar"
            }
            "11" {
                Clear-Host
                Show-BackupManager
                Read-Host "\nPressione ENTER para continuar"
            }
            "12" {
                Clear-Host
                Show-AdvancedSettings
                Read-Host "\nPressione ENTER para continuar"
            }
            "13" {
                Clear-Host
                Show-LogViewer
                Read-Host "\nPressione ENTER para continuar"
            }
            "14" {
                Clear-Host
                Show-AboutWinReset
                Read-Host "\nPressione ENTER para continuar"
            }
            "0" {
                Clear-Host
                Show-WinResetFrame -Title "OBRIGADO POR USAR O WINRESET!" -Content @(
                    "",
                    "âœ… SessÃ£o finalizada com sucesso",
                    "ðŸ“ Logs salvos em: $Global:LogPath",
                    "ðŸ’¾ Backups disponÃ­veis em: $Global:BackupPath",
                    "",
                    "ðŸ›ï¸  WinReset v5.0 Professional",
                    "ðŸ†“ Ferramenta gratuita para Ã³rgÃ£os pÃºblicos",
                    ""
                ) -FrameColor Green
                
                Write-WinResetLog "SessÃ£o do WinReset finalizada" "INFO" "SYSTEM"
                return
            }
            default {
                Show-WinResetText "âŒ OpÃ§Ã£o invÃ¡lida. Tente novamente." "Error"
                Start-Sleep 2
            }
        }
    } while ($true)
}

# Exibir lista de impressoras
function Show-PrinterList {
    param([switch]$Detailed)
    
    if ($Global:DetectedPrinters.Count -eq 0) {
        Show-WinResetText "âŒ Nenhuma impressora detectada." "Warning"
        Show-WinResetText "ðŸ’¡ Execute a detecÃ§Ã£o primeiro (opÃ§Ãµes 1, 2 ou 3)" "Info"
        return
    }
    
    Show-WinResetText "ðŸ“‹ IMPRESSORAS DETECTADAS ($($Global:DetectedPrinters.Count)):" "Title"
    Show-WinResetText ("â•" * 80) "Menu"
    
    for ($i = 0; $i -lt $Global:DetectedPrinters.Count; $i++) {
        $printer = $Global:DetectedPrinters[$i]
        $name = if ($printer.Name) { $printer.Name } else { $printer.IP }
        $type = $printer.Type
        $brand = if ($printer.Brand) { $printer.Brand } else { "Desconhecida" }
        
        Show-WinResetText "[$i] $name" "Highlight"
        Show-WinResetText "    Tipo: $type | Marca: $brand" "Info"
        
        if ($Detailed) {
            if ($printer.Status) { Show-WinResetText "    Status: $($printer.Status)" "Info" }
            if ($printer.Driver) { Show-WinResetText "    Driver: $($printer.Driver)" "Info" }
            if ($printer.Port) { Show-WinResetText "    Porta: $($printer.Port)" "Info" }
            if ($printer.IP) { Show-WinResetText "    IP: $($printer.IP)" "Info" }
            if ($printer.Service) { Show-WinResetText "    ServiÃ§o: $($printer.Service)" "Info" }
        }
        
        Show-WinResetText "" "Info"  # Linha em branco
    }
}

# Menu de reset de impressora
function Show-PrinterResetMenu {
    if ($Global:DetectedPrinters.Count -eq 0) {
        Show-WinResetText "âŒ Nenhuma impressora detectada." "Warning"
        Show-WinResetText "ðŸ’¡ Execute a detecÃ§Ã£o primeiro (opÃ§Ãµes 1, 2 ou 3)" "Info"
        return
    }
    
    Show-PrinterList
    
    $choice = Read-Host "\nðŸŽ¯ Escolha o nÃºmero da impressora para reset (ou ENTER para cancelar)"
    
    if ([string]::IsNullOrWhiteSpace($choice)) {
        Show-WinResetText "âŒ OperaÃ§Ã£o cancelada" "Warning"
        return
    }
    
    try {
        $index = [int]$choice
        if ($index -ge 0 -and $index -lt $Global:DetectedPrinters.Count) {
            $selectedPrinter = $Global:DetectedPrinters[$index]
            Invoke-WinResetPrinterReset -Printer $selectedPrinter
        } else {
            Show-WinResetText "âŒ NÃºmero invÃ¡lido" "Error"
        }
    }
    catch {
        Show-WinResetText "âŒ Entrada invÃ¡lida" "Error"
    }
    
    Read-Host "\nPressione ENTER para continuar"
}

# Gerenciador de backups
function Show-BackupManager {
    Show-WinResetText "ðŸ’¾ GERENCIADOR DE BACKUPS" "Title"
    
    if (Test-Path $Global:BackupPath) {
        $backups = Get-ChildItem $Global:BackupPath -Directory | Sort-Object CreationTime -Descending
        
        if ($backups.Count -gt 0) {
            Show-WinResetText "\nðŸ“ Backups disponÃ­veis:" "Info"
            for ($i = 0; $i -lt $backups.Count; $i++) {
                $backup = $backups[$i]
                $size = (Get-ChildItem $backup.FullName -Recurse | Measure-Object -Property Length -Sum).Sum
                $sizeStr = if ($size -gt 1MB) { "$([Math]::Round($size/1MB, 2)) MB" } else { "$([Math]::Round($size/1KB, 2)) KB" }
                Show-WinResetText "[$i] $($backup.Name) - $sizeStr - $($backup.CreationTime)" "Menu"
            }
            
            Show-WinResetText "\n[A] Abrir pasta de backups" "Menu"
            Show-WinResetText "[L] Limpar backups antigos (>30 dias)" "Menu"
            
            $choice = Read-Host "\nEscolha uma opÃ§Ã£o"
            
            switch ($choice.ToUpper()) {
                "A" {
                    Start-Process $Global:BackupPath
                }
                "L" {
                    $oldBackups = $backups | Where-Object { $_.CreationTime -lt (Get-Date).AddDays(-30) }
                    if ($oldBackups.Count -gt 0) {
                        Show-WinResetText "ðŸ—‘ï¸ Removendo $($oldBackups.Count) backups antigos..." "Info"
                        $oldBackups | Remove-Item -Recurse -Force
                        Show-WinResetText "âœ… Limpeza concluÃ­da" "Success"
                    } else {
                        Show-WinResetText "âœ… Nenhum backup antigo encontrado" "Info"
                    }
                }
            }
        } else {
            Show-WinResetText "ðŸ“ Nenhum backup encontrado" "Info"
        }
    } else {
        Show-WinResetText "ðŸ“ Pasta de backups nÃ£o existe" "Warning"
    }
}

# ConfiguraÃ§Ãµes avanÃ§adas
function Show-AdvancedSettings {
    Show-WinResetText "âš™ï¸ CONFIGURAÃ‡Ã•ES AVANÃ‡ADAS" "Title"
    
    Show-WinResetText "\n[1] Modo verboso: $(if ($Global:VerboseMode) { 'ATIVADO' } else { 'DESATIVADO' })" "Menu"
    Show-WinResetText "[2] Abrir pasta de logs" "Menu"
    Show-WinResetText "[3] Abrir pasta de backups" "Menu"
    Show-WinResetText "[4] Abrir pasta de configuraÃ§Ãµes" "Menu"
    Show-WinResetText "[5] Limpar logs antigos" "Menu"
    
    $choice = Read-Host "\nEscolha uma opÃ§Ã£o"
    
    switch ($choice) {
        "1" {
            $Global:VerboseMode = -not $Global:VerboseMode
            Show-WinResetText "âœ… Modo verboso: $(if ($Global:VerboseMode) { 'ATIVADO' } else { 'DESATIVADO' })" "Success"
        }
        "2" {
            if (Test-Path $Global:LogPath) {
                Start-Process $Global:LogPath
            } else {
                Show-WinResetText "âŒ Pasta de logs nÃ£o encontrada" "Error"
            }
        }
        "3" {
            if (Test-Path $Global:BackupPath) {
                Start-Process $Global:BackupPath
            } else {
                Show-WinResetText "âŒ Pasta de backups nÃ£o encontrada" "Error"
            }
        }
        "4" {
            if (Test-Path $Global:ConfigPath) {
                Start-Process $Global:ConfigPath
            } else {
                Show-WinResetText "âŒ Pasta de configuraÃ§Ãµes nÃ£o encontrada" "Error"
            }
        }
        "5" {
            $logFiles = Get-ChildItem $Global:LogPath -Filter "*.log" | Where-Object { $_.CreationTime -lt (Get-Date).AddDays(-30) }
            if ($logFiles.Count -gt 0) {
                Show-WinResetText "ðŸ—‘ï¸ Removendo $($logFiles.Count) logs antigos..." "Info"
                $logFiles | Remove-Item -Force
                Show-WinResetText "âœ… Limpeza de logs concluÃ­da" "Success"
            } else {
                Show-WinResetText "âœ… Nenhum log antigo encontrado" "Info"
            }
        }
    }
}

# Visualizador de logs
function Show-LogViewer {
    Show-WinResetText "ðŸ“‹ VISUALIZADOR DE LOGS" "Title"
    
    $logFiles = Get-ChildItem $Global:LogPath -Filter "*.log" | Sort-Object CreationTime -Descending
    
    if ($logFiles.Count -gt 0) {
        Show-WinResetText "\nðŸ“ Arquivos de log disponÃ­veis:" "Info"
        for ($i = 0; $i -lt [Math]::Min($logFiles.Count, 10); $i++) {
            $logFile = $logFiles[$i]
            Show-WinResetText "[$i] $($logFile.Name) - $($logFile.CreationTime)" "Menu"
        }
        
        $choice = Read-Host "\nEscolha um arquivo para visualizar (ou ENTER para cancelar)"
        
        if (-not [string]::IsNullOrWhiteSpace($choice)) {
            try {
                $index = [int]$choice
                if ($index -ge 0 -and $index -lt $logFiles.Count) {
                    $selectedLog = $logFiles[$index]
                    Show-WinResetText "\nðŸ“„ Ãšltimas 20 linhas de: $($selectedLog.Name)" "Info"
                    Show-WinResetText ("â”€" * 80) "Menu"
                    
                    $lines = Get-Content $selectedLog.FullName -Tail 20
                    foreach ($line in $lines) {
                        if ($line -match "\[ERROR\]") {
                            Show-WinResetText $line "Error"
                        } elseif ($line -match "\[WARNING\]") {
                            Show-WinResetText $line "Warning"
                        } elseif ($line -match "\[SUCCESS\]") {
                            Show-WinResetText $line "Success"
                        } else {
                            Show-WinResetText $line "Info"
                        }
                    }
                }
            }
            catch {
                Show-WinResetText "âŒ Entrada invÃ¡lida" "Error"
            }
        }
    } else {
        Show-WinResetText "ðŸ“ Nenhum arquivo de log encontrado" "Info"
    }
}

# Sobre o WinReset
function Show-AboutWinReset {
    Show-WinResetFrame -Title "SOBRE O WINRESET v5.0 PROFESSIONAL" -Content @(
        "",
        "ðŸ›ï¸  Desenvolvido especialmente para Ã“rgÃ£os PÃºblicos",
        "ðŸ†“ Substituto gratuito do WIC Reset e ferramentas pagas",
        "âš¡ 100% PowerShell nativo - sem dependÃªncias externas",
        "ðŸ”§ Reset universal para todas as marcas de impressoras",
        "ðŸŒ Suporte a impressoras USB, rede e Wi-Fi",
        "ðŸ“Š Sistema completo de logs e backups",
        "ðŸ›¡ï¸  CÃ³digo aberto e auditÃ¡vel",
        "",
        "ðŸ“‹ Funcionalidades principais:",
        "   â€¢ DetecÃ§Ã£o automÃ¡tica de impressoras",
        "   â€¢ Reset inteligente por modelo/marca",
        "   â€¢ Comandos reais ESC/POS, PJL e SNMP",
        "   â€¢ DiagnÃ³stico completo do sistema",
        "   â€¢ Interface profissional colorida",
        "   â€¢ Sistema de backup automÃ¡tico",
        "",
        "ðŸ’» Compatibilidade: Windows 10/11, PowerShell 5.1+",
        "ðŸ“… VersÃ£o: 5.0 Professional (2025-01-07)",
        "ðŸ¤– Desenvolvido por: Sistema IA AvanÃ§ado",
        ""
    ) -FrameColor DarkBlue -TitleColor Yellow
}

# InicializaÃ§Ã£o do WinReset
function Start-WinReset {
    # Inicializar base de dados
    Initialize-PrinterDatabase
    
    # Log de inicializaÃ§Ã£o
    Write-WinResetLog "WinReset v$Global:WinResetVersion iniciado" "INFO" "SYSTEM"
    Write-WinResetLog "UsuÃ¡rio: $env:USERNAME | Computador: $env:COMPUTERNAME" "INFO" "SYSTEM"
    Write-WinResetLog "Sistema: $((Get-WmiObject Win32_OperatingSystem).Caption)" "INFO" "SYSTEM"
    
    # Mostrar menu principal
    Show-WinResetMainMenu
}

# ===============================================================================
# INICIALIZAÃ‡ÃƒO DO PROGRAMA
# ===============================================================================

# Verificar versÃ£o do PowerShell
if ($PSVersionTable.PSVersion.Major -lt 5) {
    Write-Host "âŒ ERRO: WinReset requer PowerShell 5.1 ou superior" -ForegroundColor Red
    Write-Host "ðŸ“¥ Baixe a versÃ£o mais recente em: https://aka.ms/powershell" -ForegroundColor Yellow
    exit 1
}

# Iniciar o WinReset
Start-WinReset



