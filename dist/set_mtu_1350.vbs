Set objShell = CreateObject("Shell.Application")
If Not WScript.Arguments.Named.Exists("elevated") Then
    objShell.ShellExecute "wscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " /elevated", "", "runas", 1
    WScript.Quit
End If

Set WshShell = CreateObject("WScript.Shell")

strCommand = "powershell -Command ""$adapter = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' } | Select-Object -First 1; if ($adapter) { netsh interface ipv4 set subinterface $adapter.Name mtu=1350 store=persistent }"""

WshShell.Run strCommand, 0, True

MsgBox "Настройка MTU 1350 завершена для активного адаптера.", 64, "Настройка сети"
