Set WshShell = CreateObject("WScript.Shell")
result = WshShell.Run("taskkill /F /IM hysteria-windows-amd64.exe", 0, True)
result = WshShell.Run("taskkill /F /IM ProxiFyre.exe", 0, True)
If result = 0 Then
    MsgBox "Hysteria has been successfully stopped.", 64, "Process Monitor"
ElseIf result = 128 Then
    MsgBox "Hysteria process not found. It might be already closed.", 48, "Warning"
Else
    MsgBox "An error occurred while trying to stop the application.", 16, "Error"
End If
