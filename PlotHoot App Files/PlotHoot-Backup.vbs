Option Explicit

Dim shell, fso, folder, scriptPath, systemRoot, psExe, command, exitCode, extraArgs, env

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

folder = fso.GetParentFolderName(WScript.ScriptFullName)
scriptPath = fso.BuildPath(folder, "PlotHoot.ps1")
systemRoot = shell.ExpandEnvironmentStrings("%SystemRoot%")
psExe = systemRoot & "\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"

If Not fso.FileExists(psExe) Then
    psExe = systemRoot & "\System32\WindowsPowerShell\v1.0\powershell.exe"
End If

If Not fso.FileExists(scriptPath) Then
    MsgBox "Could not find PlotHoot.ps1 in the app files folder.", vbCritical, "PlotHoot"
    WScript.Quit 1
End If

extraArgs = ""
If WScript.Arguments.Named.Exists("SelfTest") Then extraArgs = extraArgs & " -SelfTest"
If WScript.Arguments.Named.Exists("UiSmokeTest") Then extraArgs = extraArgs & " -UiSmokeTest"
If WScript.Arguments.Named.Exists("NoUi") Then extraArgs = extraArgs & " -NoUi"

shell.CurrentDirectory = folder
Set env = shell.Environment("PROCESS")
env("PLOTHOOT_SCRIPT") = scriptPath
env("PLOTHOOT_APPROOT") = folder
command = """" & psExe & """ -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -Command ""Add-Type -AssemblyName System.Windows.Forms; Add-Type -AssemblyName System.Drawing; $p=$env:PLOTHOOT_SCRIPT; $s=[System.IO.File]::ReadAllText($p); . ([scriptblock]::Create($s)) @args""" & extraArgs
exitCode = shell.Run(command, 0, True)

If exitCode <> 0 Then
    MsgBox "PlotHoot closed because of an error. Run Run-PlotHoot-32bit.bat to see the full message.", vbExclamation, "PlotHoot"
End If

WScript.Quit exitCode
