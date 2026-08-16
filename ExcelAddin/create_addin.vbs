'================================================================================
' Exceller - Create Excel Add-in
'
' Builds ExcelAssistant.xlam from ExcelAssistant.bas and installs it into
' Excel automatically (compiles via Excel COM, registers it to auto-load).
'
' This is a thin wrapper around the shared one-click builder:
'   ..\ExcellerInstaller\build_addin.vbs
'
' You normally don't need to run this directly - just double-click install.bat.
'================================================================================

Option Explicit

Dim fso, shell, scriptDir, repoDir, builder, basPath, cmd, q, rc

Set fso   = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
repoDir   = fso.GetParentFolderName(scriptDir)

builder = fso.BuildPath(repoDir, "ExcellerInstaller\build_addin.vbs")
basPath = fso.BuildPath(scriptDir, "ExcelAssistant.bas")

If Not fso.FileExists(builder) Then
    MsgBox "Could not find the installer engine:" & vbCrLf & builder, _
           vbCritical, "Exceller"
    WScript.Quit 1
End If

q = Chr(34)   ' double quote
cmd = "cscript //nologo " & q & builder & q & " " & _
      q & basPath & q & " " & _
      q & q & " " & _                    ' no form file
      q & "ExcelAssistant" & q

rc = shell.Run(cmd, 1, True)
WScript.Quit rc
