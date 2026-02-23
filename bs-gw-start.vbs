Set WshShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

Sub KillProcess(strProcessName)
    Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name = '" & strProcessName & "'")
    For Each objProcess in colProcessList
        objProcess.Terminate()
    Next
End Sub

KillProcess "hysteria-windows-amd64.exe"
KillProcess "ProxiFyre.exe"

WScript.Sleep 2000

WshShell.Run "cmd.exe /c bin\hysteria-windows-amd64.exe client --config=bin\config.yaml", 0, False
WshShell.Run "cmd.exe /c bin\ProxiFyre.exe run", 0, False

WScript.Sleep 5000

checkScriptPath = objFSO.GetAbsolutePathName("bin\check_and_run.vbs")

If objFSO.FileExists(checkScriptPath) Then
    WshShell.CurrentDirectory = objFSO.GetParentFolderName(checkScriptPath)
    WshShell.Run "wscript.exe check_and_run.vbs", 0, False
End If
