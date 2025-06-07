Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = "C:\Users\rindl\OneDrive\Desktop\exceljax"
WshShell.Run "cmd /c npm start", 0, False
