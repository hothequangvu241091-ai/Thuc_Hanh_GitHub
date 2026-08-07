Option Explicit

Dim shell, managerPath, command
managerPath = WScript.Arguments(0)
Set shell = CreateObject("WScript.Shell")
command = "powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File " & Chr(34) & managerPath & Chr(34)
shell.Run command, 0, False
