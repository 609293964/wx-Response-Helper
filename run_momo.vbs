Option Explicit

Dim shell, app, fso
Set shell = CreateObject("WScript.Shell")
Set app = CreateObject("Shell.Application")
Set fso = CreateObject("Scripting.FileSystemObject")

Dim projectDir, executable, pythonw, script
projectDir = fso.GetParentFolderName(WScript.ScriptFullName)
executable = projectDir & "\wechat_gui_momo.exe"
pythonw = projectDir & "\.venv\Scripts\pythonw.exe"
script = projectDir & "\wechat_gui_momo.py"

If fso.FileExists(executable) Then
    shell.CurrentDirectory = projectDir
    On Error Resume Next
    Err.Clear
    app.ShellExecute executable, "", projectDir, "runas", 1
    If Err.Number <> 0 Then
        MsgBox "Startup failed: " & Err.Description, vbCritical, "EasyChat Momo"
        WScript.Quit Err.Number
    End If
    On Error GoTo 0
    WScript.Quit 0
End If

If Not fso.FileExists(pythonw) Then
    MsgBox "wechat_gui_momo.exe and the Python environment were not found." & vbCrLf & vbCrLf & "Use the portable package, or run install_deps.bat first.", vbCritical, "EasyChat Momo"
    WScript.Quit 1
End If

If Not fso.FileExists(script) Then
    MsgBox "Application entry was not found:" & vbCrLf & script, vbCritical, "EasyChat Momo"
    WScript.Quit 1
End If

shell.CurrentDirectory = projectDir
On Error Resume Next
Err.Clear
app.ShellExecute pythonw, """" & script & """", projectDir, "runas", 1
If Err.Number <> 0 Then
    MsgBox "Startup failed: " & Err.Description, vbCritical, "EasyChat Momo"
    WScript.Quit Err.Number
End If
On Error GoTo 0
