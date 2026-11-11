' ConvertFlow - Silent Context Menu Launcher (Prevents CMD Window Flash)
Option Explicit

Dim args, shell, fso, scriptDir, nodeExe, cliScript, cmd, i
Set args = WScript.Arguments
If args.Count < 2 Then WScript.Quit 1

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
cliScript = fso.BuildPath(scriptDir, "convert-cli.js")

' Use explicit Node path if passed as 3rd arg, else default to "node.exe"
If args.Count >= 3 Then
  nodeExe = args(2)
Else
  nodeExe = "node.exe"
End If

cmd = """" & nodeExe & """ """ & cliScript & """ """ & args(0) & """ """ & args(1) & """"
shell.Run cmd, 0, False

