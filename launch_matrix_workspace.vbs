Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
helperPath = fso.BuildPath(scriptDir, "launch_matrix_workspace.cmd")
shell.CurrentDirectory = scriptDir
shell.Run Chr(34) & helperPath & Chr(34), 1, False
