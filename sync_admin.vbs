WScript.Sleep 2000

Dim objFSO
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
CopyFile "\\10.66.4.215\sync-files\hosts", "C:\Windows\System32\drivers\etc\hosts", True