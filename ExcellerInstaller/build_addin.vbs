'===============================================================================
' build_addin.vbs - One-click Excel add-in builder/installer
'
' Compiles a .xlam add-in from VBA source files using Excel COM automation,
' installs it into the user's AddIns folder, and registers it so Excel loads
' it automatically on startup. No manual VBA steps are required.
'
' Usage:
'   cscript //nologo build_addin.vbs "<path to .bas>" ["<path to .frm>"] ["<add-in name>"]
'
' Examples:
'   cscript //nologo build_addin.vbs "C:\repo\ExcelAddin\ExcelAssistant.bas" "" "ExcelAssistant"
'   cscript //nologo build_addin.vbs "C:\repo\ExcelAssistant\ExcelAssistant.bas" "C:\repo\ExcelAssistant\frmAssistant.frm" "ExcelAssistant"
'
' Exit codes: 0 = success, 1 = failure
'===============================================================================

Option Explicit

Dim fso, shell
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

'------------------------------------------------------------------------------
' Parse arguments
'------------------------------------------------------------------------------
If WScript.Arguments.Count < 1 Then
    WScript.Echo "Usage: cscript //nologo build_addin.vbs ""<bas file>"" [""<frm file>""] [""<add-in name>""]"
    WScript.Quit 1
End If

Dim basFile, frmFile, addinName
basFile   = WScript.Arguments(0)
frmFile   = ""
addinName = ""
If WScript.Arguments.Count > 1 Then frmFile = WScript.Arguments(1)
If WScript.Arguments.Count > 2 Then addinName = WScript.Arguments(2)

If Not fso.FileExists(basFile) Then
    WScript.Echo "ERROR: VBA source file not found: " & basFile
    WScript.Quit 1
End If

If Len(addinName) = 0 Then addinName = fso.GetBaseName(basFile)

'------------------------------------------------------------------------------
' Prepare the AddIns folder (a default trusted location for Excel)
'------------------------------------------------------------------------------
Dim addinsDir, xlamPath
addinsDir = shell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\AddIns"
If Not fso.FolderExists(addinsDir) Then fso.CreateFolder addinsDir
xlamPath = addinsDir & "\" & addinName & ".xlam"

'------------------------------------------------------------------------------
' Enable programmatic access to the VBA object model so the add-in can be
' compiled automatically (HKCU only - no administrator rights needed).
' Written for Office 2010/2013/2016+; extra keys are inert.
'------------------------------------------------------------------------------
Dim ver
For ver = 16 To 14 Step -1
    shell.RegWrite "HKCU\Software\Microsoft\Office\" & ver & ".0\Excel\Security\AccessVBOM", 1, "REG_DWORD"
Next

WScript.Echo "Building " & xlamPath & " ..."

'------------------------------------------------------------------------------
' Compile the .xlam through Excel COM automation
'------------------------------------------------------------------------------
Dim xlApp
On Error Resume Next
Set xlApp = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Could not start Microsoft Excel." & vbCrLf & _
                 "       Make sure Excel is installed, then try again."
    WScript.Quit 1
End If
Err.Clear
On Error GoTo 0

xlApp.Visible = False
xlApp.DisplayAlerts = False

Dim xlWB, vbProj, vbComp
On Error Resume Next
Set xlWB   = xlApp.Workbooks.Add
Set vbProj = xlWB.VBProject
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Cannot access the Excel VBA project (" & Err.Description & ")." & vbCrLf & _
                 "       Close all Excel windows, then run the installer again."
    xlApp.Quit
    WScript.Quit 1
End If
Err.Clear
On Error GoTo 0

' Import the standard module (.bas)
Set vbComp = vbProj.VBComponents.Add(1)   ' 1 = vbext_ct_StdModule
ImportModule vbComp, basFile

' Import the user form (.frm), if provided
If Len(frmFile) > 0 Then
    If Not fso.FileExists(frmFile) Then
        WScript.Echo "WARNING: Form file not found, skipping: " & frmFile
    Else
        Dim frmCopy
        frmCopy = NormalizeLineEndings(frmFile)
        On Error Resume Next
        vbProj.VBComponents.Import frmCopy
        If Err.Number <> 0 Then
            WScript.Echo "WARNING: Could not import form (" & Err.Description & ")"
            Err.Clear
        End If
        On Error GoTo 0
        fso.DeleteFile frmCopy, True
    End If
End If

' Save as an Excel add-in (FileFormat 51 = xlAddIn)
On Error Resume Next
If fso.FileExists(xlamPath) Then fso.DeleteFile xlamPath, True
xlWB.SaveAs xlamPath, 51
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Could not save the add-in (" & Err.Description & ")" & vbCrLf & _
                 "       Close any Excel window that may have the file open, then retry."
    xlApp.Quit
    WScript.Quit 1
End If
Err.Clear
On Error GoTo 0

xlWB.Close False
xlApp.Quit
Set xlApp = Nothing

'------------------------------------------------------------------------------
' Register the add-in so Excel loads it automatically on startup.
' This writes the same OPEN registry values that Excel's own Add-ins dialog
' creates when you check an add-in.
'------------------------------------------------------------------------------
RegisterAddIn addinName, xlamPath

WScript.Echo ""
WScript.Echo "SUCCESS: " & addinName & " installed to:"
WScript.Echo "  " & xlamPath
WScript.Echo "Excel will load it automatically the next time it starts."
WScript.Quit 0

'===============================================================================
' Helpers
'===============================================================================

Sub ImportModule(vbComp, filePath)
    ' Read the module, normalize line endings, and inject it into the VBA project.
    ' Uses AddFromString when small; falls back to line-by-line InsertLines.
    Dim text, normalized, ts
    Set ts = fso.OpenTextFile(filePath, 1, False, 0)   ' ForReading, ASCII
    text = ts.ReadAll
    ts.Close

    normalized = Replace(text, vbCrLf, vbLf)
    normalized = Replace(normalized, vbCr, vbLf)
    normalized = Replace(normalized, vbLf, vbCrLf)

    On Error Resume Next
    If Len(normalized) < 60000 Then
        vbComp.CodeModule.AddFromString normalized
        If Err.Number = 0 Then
            On Error GoTo 0
            Exit Sub
        End If
        Err.Clear
    End If
    On Error GoTo 0

    Dim lines, i, line
    lines = Split(normalized, vbCrLf)
    For i = 0 To UBound(lines)
        line = lines(i)
        If Len(line) > 0 Then
            vbComp.CodeModule.InsertLines i + 1, line
        End If
    Next
End Sub

Function NormalizeLineEndings(filePath) As String
    ' Write a CRLF copy of the file to %TEMP% (VBA import requires CRLF) and
    ' return its path.
    Dim text, normalized, tmpPath, tsIn, tsOut
    Set tsIn = fso.OpenTextFile(filePath, 1, False, 0)
    text = tsIn.ReadAll
    tsIn.Close

    normalized = Replace(text, vbCrLf, vbLf)
    normalized = Replace(normalized, vbCr, vbLf)
    normalized = Replace(normalized, vbLf, vbCrLf)

    tmpPath = shell.ExpandEnvironmentStrings("%TEMP%") & "\" & fso.GetBaseName(filePath) & ".tmp.frm"
    Set tsOut = fso.CreateTextFile(tmpPath, True, False)
    tsOut.Write normalized
    tsOut.Close
    NormalizeLineEndings = tmpPath
End Function

Sub RegisterAddIn(addinName, xlamPath)
    Dim regBase, i, v, valueName, existing, match
    match = LCase(addinName & ".xlam")

    For v = 16 To 14 Step -1
        regBase = "HKCU\Software\Microsoft\Office\" & v & ".0\Excel\Options"

        ' Remove any previous OPEN entries that point to this add-in,
        ' so it is never registered twice.
        For i = 0 To 99
            valueName = "OPEN"
            If i > 0 Then valueName = valueName & i
            On Error Resume Next
            existing = shell.RegRead(regBase & "\" & valueName)
            If Err.Number = 0 Then
                If InStr(LCase(existing), match) > 0 Then
                    shell.RegDelete regBase & "\" & valueName
                End If
            Else
                Err.Clear
            End If
            On Error GoTo 0
        Next

        ' Write the new entry into the first free OPEN slot.
        ' The value data includes the surrounding quotes, as Excel expects.
        For i = 0 To 99
            valueName = "OPEN"
            If i > 0 Then valueName = valueName & i
            On Error Resume Next
            shell.RegRead regBase & "\" & valueName
            If Err.Number <> 0 Then
                Err.Clear
                shell.RegWrite regBase & "\" & valueName, """" & xlamPath & """", "REG_SZ"
                On Error GoTo 0
                Exit For
            End If
            On Error GoTo 0
        Next
    Next
End Sub
