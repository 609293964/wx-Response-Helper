Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

projectDir = fso.GetParentFolderName(WScript.ScriptFullName)
pythonw = projectDir & "\.venv\Scripts\pythonw.exe"
script = projectDir & "\wechat_gui_momo.py"

shell.CurrentDirectory = projectDir
shell.Run """" & pythonw & """ """ & script & """", 1, False
