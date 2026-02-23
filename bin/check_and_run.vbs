Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

scriptPath = "set_mtu_1350.vbs"
fullPath = objFSO.GetAbsolutePathName(scriptPath)

psCheck = "powershell -WindowStyle Hidden -Command ""$adapter = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' } | Select-Object -First 1; " & _
          "if (-not $adapter) { exit 0 }; " & _
          "$mtu = (netsh interface ipv4 show subinterface $adapter.Name | Select-String '^\s*\d+' | ForEach-Object { $_.ToString().Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries)[0] }); " & _
          "if ($mtu -ne '1350') { exit 1 } else { exit 0 }"""

result = objShell.Run(psCheck, 0, True)

If result = 1 Then
    If objFSO.FileExists(fullPath) Then
        objShell.Run "wscript.exe """ & fullPath & """", 1, False
    End If
End If
