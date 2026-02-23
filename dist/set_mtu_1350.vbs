Set objShell = CreateObject("Shell.Application")
If Not WScript.Arguments.Named.Exists("elevated") Then
    objShell.ShellExecute "wscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " /elevated", "", "runas", 1
    WScript.Quit
End If

Set WshShell = CreateObject("WScript.Shell")

psCommand = "powershell -WindowStyle Hidden -Command ""Add-Type -AssemblyName System.Windows.Forms; " & _
            "$adapter = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' } | Select-Object -First 1; " & _
            "if (-not $adapter) { " & _
            "  [System.Windows.Forms.MessageBox]::Show('No active network adapter found.', 'Network Setup'); exit " & _
            "}; " & _
            "$raw = netsh interface ipv4 show subinterface ($adapter.Name); " & _
            "$currentMtu = ($raw | Select-String '^\s*\d+' | ForEach-Object { $_.ToString().Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries) }); " & _
            "if ($currentMtu -eq '1350') { " & _
            "  [System.Windows.Forms.MessageBox]::Show('MTU is already set to 1350. No changes needed.', 'Network Setup') " & _
            "} else { " & _
            "  netsh interface ipv4 set subinterface ($adapter.Name) mtu=1350 store=persistent; " & _
            "  ipconfig /flushdns; " & _
            "  [System.Windows.Forms.MessageBox]::Show('MTU set to 1350 and DNS cache flushed successfully.', 'Network Setup') " & _
            "}"""

WshShell.Run psCommand, 0, True
