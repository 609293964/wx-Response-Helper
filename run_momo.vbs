Set shell = CreateObject("WScript.Shell")
Set app = CreateObject("Shell.Application")
Set fso = CreateObject("Scripting.FileSystemObject")

projectDir = fso.GetParentFolderName(WScript.ScriptFullName)
pythonw = projectDir & "\.venv\Scripts\pythonw.exe"
script = projectDir & "\wechat_gui_momo.py"

shell.CurrentDirectory = projectDir
app.ShellExecute pythonw, """" & script & """", projectDir, "runas", 1
