Set WshShell = CreateObject("WScript.Shell")

' Set working directory
WshShell.CurrentDirectory = "C:\Users\rindl\OneDrive\Desktop\exceljax"

' Run proxy silently
WshShell.Run "cmd /c npm run proxy", 0, False

' Wait 3 seconds
WScript.Sleep 3000

' Run Excel add-in silently
WshShell.Run "cmd /c npm start", 0, False
