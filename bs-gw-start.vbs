Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cmd.exe /c bin\hysteria-windows-amd64.exe client --config=bin\config.yaml", 0, False
WshShell.Run "cmd.exe /c bin\ProxiFyre.exe run", 0, False