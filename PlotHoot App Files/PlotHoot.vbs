Option Explicit

Dim shell, fso, appFolder, scriptPath, systemRoot, psExe, command, exitCode, extraArgs, file, env

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

appFolder = fso.GetParentFolderName(WScript.ScriptFullName)
scriptPath = fso.BuildPath(appFolder, "PlotHoot.ps1")
If Not fso.FileExists(scriptPath) Then
    scriptPath = ""
    For Each file In fso.GetFolder(appFolder).Files
        If LCase(fso.GetExtensionName(file.Name)) = "ps1" And LCase(Left(file.Name, 4)) = "plot" Then
            scriptPath = file.Path
            Exit For
        End If
    Next
End If
systemRoot = shell.ExpandEnvironmentStrings("%SystemRoot%")
psExe = systemRoot & "\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"

If Not fso.FileExists(psExe) Then
    psExe = systemRoot & "\System32\WindowsPowerShell\v1.0\powershell.exe"
End If

If Len(scriptPath) = 0 Or Not fso.FileExists(scriptPath) Then
    MsgBox "Could not find the PlotHoot app file in the app files folder.", vbCritical, "PlotHoot"
    WScript.Quit 1
End If

If Not fso.FileExists(psExe) Then
    MsgBox "Could not find Windows PowerShell.", vbCritical, "PlotHoot"
    WScript.Quit 1
End If

extraArgs = ""
If WScript.Arguments.Named.Exists("SelfTest") Then extraArgs = extraArgs & " -SelfTest"
If WScript.Arguments.Named.Exists("UiSmokeTest") Then extraArgs = extraArgs & " -UiSmokeTest"
If WScript.Arguments.Named.Exists("NoUi") Then extraArgs = extraArgs & " -NoUi"

shell.CurrentDirectory = appFolder
Set env = shell.Environment("PROCESS")
env("PLOTHOOT_SCRIPT") = scriptPath
env("PLOTHOOT_APPROOT") = appFolder
command = """" & psExe & """ -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -Command ""Add-Type -AssemblyName System.Windows.Forms; Add-Type -AssemblyName System.Drawing; $p=$env:PLOTHOOT_SCRIPT; $s=[System.IO.File]::ReadAllText($p); . ([scriptblock]::Create($s)) @args""" & extraArgs
exitCode = shell.Run(command, 0, True)

If exitCode <> 0 Then
    MsgBox "PlotHoot closed because of an error. Run PlotHoot-Clean.cmd or Run-PlotHoot-32bit.bat to see the full message.", vbExclamation, "PlotHoot"
End If

WScript.Quit exitCode
