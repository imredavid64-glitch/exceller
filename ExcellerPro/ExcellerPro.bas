Attribute VB_Name = "ExcellerPro"
'==============================================================================
' Exceller Pro - AI-Powered Excel Assistant
' Version: 3.0.0
' License: Free to use
' Support: https://github.com/exceller/excel-addin
'==============================================================================

Option Explicit

'------------------------------------------------------------------------------
' Configuration
'------------------------------------------------------------------------------
Private Const APP_NAME As String = "Exceller Pro"
Private Const APP_VERSION As String = "3.0.0"
Private Const DEFAULT_API_URL As String = "https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent"
Private Const MAX_RESPONSE_CHARS As Long = 8000
Private Const TIMEOUT_SECONDS As Long = 30

'------------------------------------------------------------------------------
' State
'------------------------------------------------------------------------------
Private isInitialized As Boolean
Private cachedApiKey As String

'==============================================================================
' INITIALIZATION
'==============================================================================

Private Sub Auto_Open()
    On Error Resume Next
    Call Initialize
    On Error GoTo 0
End Sub

Private Sub Initialize()
    If isInitialized Then Exit Sub
    
    ' Load or create settings
    Call LoadSettings
    
    ' Create custom toolbar
    Call CreateToolbar
    
    isInitialized = True
End Sub

Private Sub LoadSettings()
    ' Load API key from registry
    cachedApiKey = GetSetting(APP_NAME, "Config", "ApiKey", "")
    
    ' First run - show setup
    If Len(cachedApiKey) = 0 Then
        Call ShowFirstRunSetup
    End If
End Sub

Private Sub ShowFirstRunSetup()
    Dim response As VbMsgBoxResult
    
    response = MsgBox("Welcome to " & APP_NAME & "!" & vbNewLine & vbNewLine & _
                      "This add-in helps you with:" & vbNewLine & _
                      Chr(149) & " AI-powered data analysis" & vbNewLine & _
                      Chr(149) & " Formula creation and fixes" & vbNewLine & _
                      Chr(149) & " Smart table formatting" & vbNewLine & _
                      Chr(149) & " Price calculations" & vbNewLine & vbNewLine & _
                      "Would you like to set up your API key now?" & vbNewLine & _
                      "(Required for AI features)", _
                      vbYesNo + vbQuestion, APP_NAME)
    
    If response = vbYes Then
        Call ExcellerPro_Settings
    End If
End Sub

'==============================================================================
' TOOLBAR
'==============================================================================

Private Sub CreateToolbar()
    Dim cmdBar As CommandBar
    Dim cmdBtn As CommandBarButton
    
    ' Remove existing toolbar
    On Error Resume Next
    Application.CommandBars(APP_NAME).Delete
    On Error GoTo 0
    
    ' Create new toolbar
    Set cmdBar = Application.CommandBars.Add( _
        Name:=APP_NAME, _
        Position:=msoBarTop, _
        Temporary:=True)
    
    ' Main chat button
    Set cmdBtn = cmdBar.Controls.Add(Type:=msoControlButton)
    With cmdBtn
        .Caption = "Ask AI"
        .OnAction = "=ExcellerPro_Chat()"
        .FaceId = 9854
        .TooltipText = "Chat with AI assistant"
        .Style = msoButtonCaption
    End With
    
    ' Analyze button
    Set cmdBtn = cmdBar.Controls.Add(Type:=msoControlButton)
    With cmdBtn
        .Caption = "Analyze"
        .OnAction = "=ExcellerPro_Analyze()"
        .FaceId = 629
        .TooltipText = "Analyze selected data"
        .Style = msoButtonCaption
    End With
    
    ' Separator
    cmdBar.Controls.Add Type:=msoControlSeparator
    
    ' Fix formula button
    Set cmdBtn = cmdBar.Controls.Add(Type:=msoControlButton)
    With cmdBtn
        .Caption = "Fix Formula"
        .OnAction = "=ExcellerPro_FixFormula()"
        .FaceId = 3146
        .TooltipText = "Fix broken formulas"
        .Style = msoButtonCaption
    End With
    
    ' Price calculator button
    Set cmdBtn = cmdBar.Controls.Add(Type:=msoControlButton)
    With cmdBtn
        .Caption = "Price Calc"
        .OnAction = "=ExcellerPro_PriceCalc()"
        .FaceId = 1636
        .TooltipText = "Price calculation help"
        .Style = msoButtonCaption
    End With
    
    ' Separator
    cmdBar.Controls.Add Type:=msoControlSeparator
    
    ' Settings button
    Set cmdBtn = cmdBar.Controls.Add(Type:=msoControlButton)
    With cmdBtn
        .Caption = "Settings"
        .OnAction = "=ExcellerPro_Settings()"
        .FaceId = 925
        .TooltipText = "Configure settings"
        .Style = msoButtonCaption
    End With
    
    cmdBar.Visible = True
End Sub

'==============================================================================
' PUBLIC COMMANDS
'==============================================================================

Public Sub ExcellerPro_Chat()
    Call ShowAssistant("chat")
End Sub

Public Sub ExcellerPro_Analyze()
    Call ShowAssistant("analyze")
End Sub

Public Sub ExcellerPro_FixFormula()
    Call ShowAssistant("fix")
End Sub

Public Sub ExcellerPro_VerifyFormula()
    Call ShowAssistant("verify")
End Sub

Public Sub ExcellerPro_CreateTable()
    Call ShowAssistant("table")
End Sub

Public Sub ExcellerPro_PriceCalc()
    Call ShowAssistant("price")
End Sub

Public Sub ExcellerPro_FormulaHelp()
    Call ShowAssistant("formula")
End Sub

Public Sub ExcellerPro_Settings()
    Call ShowSettingsDialog
End Sub

Public Sub ExcellerPro_Help()
    Call ShowHelp
End Sub

Public Sub ExcellerPro_About()
    Call ShowAbout
End Sub

Public Sub ExcellerPro_Uninstall()
    Call UninstallAddin
End Sub

'==============================================================================
' MAIN ASSISTANT
'==============================================================================

Private Sub ShowAssistant(mode As String)
    ' Check API key
    If Len(cachedApiKey) = 0 Then
        MsgBox "Please set your API key first." & vbNewLine & _
               "Go to: Exceller Pro > Settings", _
               vbExclamation, APP_NAME
        Call ExcellerPro_Settings
        Exit Sub
    End If
    
    Dim userInput As String
    Dim context As String
    
    ' Get context
    context = GetSelectedContext()
    
    ' Get user input based on mode
    Select Case mode
        Case "chat"
            userInput = ShowInputDialog( _
                "Ask anything about your data:" & vbNewLine & vbNewLine & _
                "Examples:" & vbNewLine & _
                Chr(149) & " Summarize this data" & vbNewLine & _
                Chr(149) & " Find duplicates" & vbNewLine & _
                Chr(149) & " Create a chart" & vbNewLine & _
                Chr(149) & " Write a formula", _
                "AI Chat")
                
        Case "analyze"
            If TypeName(Selection) <> "Range" Then
                MsgBox "Please select data first.", vbExclamation, APP_NAME
                Exit Sub
            End If
            userInput = "Analyze this data thoroughly. Provide: summary, key insights, statistics, anomalies, and recommendations."
            
        Case "fix"
            userInput = ShowInputDialog( _
                "Enter the broken formula and error:" & vbNewLine & vbNewLine & _
                "Example: =VLOOKUP(A1,Sheet2!A:B,2,FALSE)" & vbNewLine & _
                "Error: #REF!", _
                "Fix Formula")
            If Len(userInput) = 0 Then Exit Sub
            context = "BROKEN FORMULA: " & context
            
        Case "verify"
            If TypeName(Selection) = "Range" Then
                If Len(Selection.Formula) > 0 Then
                    userInput = "Verify this formula is correct and explain what it does: " & Selection.Formula
                Else
                    userInput = ShowInputDialog("Enter formula to verify:", "Verify Formula")
                End If
            Else
                userInput = ShowInputDialog("Enter formula to verify:", "Verify Formula")
            End If
            
        Case "table"
            If TypeName(Selection) <> "Range" Then
                MsgBox "Please select data first.", vbExclamation, APP_NAME
                Exit Sub
            End If
            userInput = "Create a professional table from this data. Suggest: headers, formatting, calculations, and improvements."
            
        Case "price"
            userInput = ShowInputDialog( _
                "Describe your price calculation:" & vbNewLine & vbNewLine & _
                "Examples:" & vbNewLine & _
                Chr(149) & " Calculate total with 10% discount" & vbNewLine & _
                Chr(149) & " Add 8% tax to subtotal" & vbNewLine & _
                Chr(149) & " Price lookup from another table", _
                "Price Calculator")
            
        Case "formula"
            userInput = ShowInputDialog( _
                "What do you want to calculate?" & vbNewLine & vbNewLine & _
                "Examples:" & vbNewLine & _
                Chr(149) & " Sum if date is in January" & vbNewLine & _
                Chr(149) & " Count unique items" & vbNewLine & _
                Chr(149) & " Look up value in another table", _
                "Formula Help")
    End Select
    
    ' Exit if no input
    If Len(userInput) = 0 Then Exit Sub
    
    ' Process
    Call ProcessRequest(userInput, context, mode)
End Sub

'==============================================================================
' PROCESS REQUEST
'==============================================================================

Private Sub ProcessRequest(prompt As String, context As String, mode As String)
    ' Show processing
    Application.StatusBar = "Exceller Pro is thinking..."
    Application.ScreenUpdating = False
    Application.Cursor = xlWait
    
    ' Get response
    Dim response As String
    response = CallGeminiAPI(prompt, context)
    
    ' Reset Excel state
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
    
    ' Display response
    Call DisplayResponse(response, mode)
End Sub

'==============================================================================
' API COMMUNICATION
'==============================================================================

Private Function CallGeminiAPI(prompt As String, context As String) As String
    Dim httpReq As Object
    Dim requestBody As String
    Dim fullPrompt As String
    
    ' Build full prompt
    fullPrompt = BuildPrompt(prompt, context)
    
    ' Create HTTP request
    Set httpReq = CreateObject("MSXML2.XMLHTTP")
    
    ' Build JSON body
    requestBody = BuildRequestBody(fullPrompt)
    
    ' Send request
    On Error GoTo ErrorHandler
    httpReq.Open "POST", DEFAULT_API_URL & "?key=" & cachedApiKey, False
    httpReq.setRequestHeader "Content-Type", "application/json"
    httpReq.send requestBody
    
    ' Check response
    If httpReq.Status = 200 Then
        CallGeminiAPI = ParseResponse(httpReq.responseText)
    Else
        CallGeminiAPI = "Error: API returned status " & httpReq.Status & vbNewLine & _
                       httpReq.responseText
    End If
    
    Set httpReq = Nothing
    Exit Function
    
ErrorHandler:
    CallGeminiAPI = "Error: " & Err.Description & vbNewLine & _
                   "Check your internet connection and API key."
    Set httpReq = Nothing
End Function

Private Function BuildPrompt(prompt As String, context As String) As String
    Dim parts As String
    
    ' System instructions
    parts = "You are Exceller Pro, an expert Excel assistant. " & _
            "Provide clear, actionable advice. Format responses nicely. " & _
            "When showing formulas, use code formatting." & vbNewLine & vbNewLine
    
    ' Add context if provided
    If Len(context) > 0 Then
        parts = parts & "=== DATA CONTEXT ===" & vbNewLine & context & vbNewLine & vbNewLine
    End If
    
    ' Add user request
    parts = parts & "=== USER REQUEST ===" & vbNewLine & prompt
    
    BuildPrompt = parts
End Function

Private Function BuildRequestBody(prompt As String) As String
    Dim json As String
    
    json = "{" & _
           """contents"":[{""parts"":[{""text"":""" & EscapeJson(prompt) & """}]}]," & _
           """generationConfig"":{" & _
           """maxOutputTokens"":2048," & _
           """temperature"":0.7," & _
           """topP"":0.8" & _
           "}," & _
           """safetySettings"":[{""category"":""HARM_CATEGORY_HARASSMENT"",""threshold"":""BLOCK_NONE""}]" & _
           "}"
    
    BuildRequestBody = json
End Function

Private Function EscapeJson(str As String) As String
    Dim result As String
    result = str
    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbNewLine, "\n")
    result = Replace(result, vbCr, "\n")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbTab, "\t")
    EscapeJson = result
End Function

Private Function ParseResponse(jsonResponse As String) As String
    Dim startPos As Long
    Dim endPos As Long
    
    ' Find text field
    startPos = InStr(jsonResponse, """text"":""")
    If startPos > 0 Then
        startPos = startPos + 8
        endPos = InStr(startPos, jsonResponse, """")
        If endPos > startPos Then
            Dim text As String
            text = Mid(jsonResponse, startPos, endPos - startPos)
            text = Replace(text, "\n", vbNewLine)
            text = Replace(text, "\""", """")
            text = Replace(text, "\\", "\")
            ParseResponse = text
        Else
            ParseResponse = "Error: Could not parse response"
        End If
    Else
        ' Check for error
        If InStr(jsonResponse, """error""") > 0 Then
            ParseResponse = "API Error: " & ExtractErrorMessage(jsonResponse)
        Else
            ParseResponse = "Error: No response received"
        End If
    End If
End Function

Private Function ExtractErrorMessage(json As String) As String
    Dim startPos As Long
    Dim endPos As Long
    
    startPos = InStr(json, """message"":""")
    If startPos > 0 Then
        startPos = startPos + 12
        endPos = InStr(startPos, json, """")
        If endPos > startPos Then
            ExtractErrorMessage = Mid(json, startPos, endPos - startPos)
        Else
            ExtractErrorMessage = "Unknown error"
        End If
    Else
        ExtractErrorMessage = "Unknown error"
    End If
End Function

'==============================================================================
' CONTEXT HELPERS
'==============================================================================

Private Function GetSelectedContext() As String
    If TypeName(Selection) <> "Range" Then
        GetSelectedContext = "No data selected"
        Exit Function
    End If
    
    Dim rng As Range
    Set rng = Selection
    
    Dim ctx As String
    ctx = "Range: " & rng.Address(True, True, xlA1, True) & vbNewLine
    ctx = ctx & "Sheet: " & rng.Worksheet.Name & vbNewLine
    ctx = ctx & "Size: " & rng.Rows.Count & " rows x " & rng.Columns.Count & " columns" & vbNewLine
    
    ' Add headers
    Dim headers As String
    Dim c As Long
    For c = 1 To rng.Columns.Count
        If c > 1 Then headers = headers & " | "
        headers = headers & rng.Cells(1, c).Value
    Next c
    ctx = ctx & "Headers: " & headers & vbNewLine
    
    ' Add sample data
    ctx = ctx & "Sample:" & vbNewLine
    Dim maxRows As Long
    maxRows = Application.Min(rng.Rows.Count, 15)
    Dim r As Long
    For r = 1 To maxRows
        Dim rowData As String
        For c = 1 To rng.Columns.Count
            If c > 1 Then rowData = rowData & " | "
            rowData = rowData & rng.Cells(r, c).Value
        Next c
        ctx = ctx & "Row " & r & ": " & rowData & vbNewLine
    Next r
    
    If rng.Rows.Count > maxRows Then
        ctx = ctx & "... (" & (rng.Rows.Count - maxRows) & " more rows)" & vbNewLine
    End If
    
    GetSelectedContext = ctx
End Function

'==============================================================================
' UI HELPERS
'==============================================================================

Private Function ShowInputDialog(prompt As String, title As String) As String
    ShowInputDialog = InputBox(prompt, APP_NAME & " - " & title)
End Function

Private Sub DisplayResponse(response As String, mode As String)
    ' Create new worksheet
    Dim ws As Worksheet
    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    
    ' Name it
    On Error Resume Next
    ws.Name = "Exceller " & Format(Now, "hhmmss")
    On Error GoTo 0
    
    ' Format header
    ws.Range("A1").Value = APP_NAME & " Response"
    ws.Range("A1").Font.Size = 18
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Color = RGB(76, 175, 80)
    
    ws.Range("A2").Value = "Mode: " & UCase(mode) & " | " & Format(Now, "mm/dd/yyyy hh:mm AM/PM")
    ws.Range("A2").Font.Size = 10
    ws.Range("A2").Font.Color = RGB(128, 128, 128)
    
    ' Add response
    ws.Range("A4").Value = response
    ws.Range("A4").WrapText = True
    ws.Range("A4").Font.Size = 11
    ws.Columns("A").ColumnWidth = 80
    ws.Rows("4").RowHeight = Application.Min(ws.Rows("4").RowHeight, 400)
    
    ' Select it
    ws.Range("A4").Select
    
    ' Also show message box
    If Len(response) <= 500 Then
        MsgBox response, vbInformation, APP_NAME
    End If
End Sub

'==============================================================================
' SETTINGS
'==============================================================================

Private Sub ShowSettingsDialog()
    Dim currentKey As String
    Dim newKey As String
    Dim info As String
    
    currentKey = GetSetting(APP_NAME, "Config", "ApiKey", "")
    
    info = APP_NAME & " Settings" & vbNewLine & vbNewLine
    info = info & "Version: " & APP_VERSION & vbNewLine
    info = info & "API Key: " & IIf(Len(currentKey) > 0, String(Len(currentKey) - 8, "*") & Right(currentKey, 8), "Not set") & vbNewLine & vbNewLine
    info = info & "Enter your Gemini API key below:" & vbNewLine
    info = info & "(Get free key at: https://makersuite.google.com/app/apikey)"
    
    newKey = InputBox(info, APP_NAME & " - Settings", currentKey)
    
    If Len(newKey) > 0 Then
        If newKey <> currentKey Then
            cachedApiKey = newKey
            SaveSetting APP_NAME, "Config", "ApiKey", newKey
            MsgBox "Settings saved!", vbInformation, APP_NAME
        End If
    End If
End Sub

'==============================================================================
' HELP
'==============================================================================

Private Sub ShowHelp()
    Dim helpText As String
    
    helpText = APP_NAME & " Help" & vbNewLine & _
               String(50, "=") & vbNewLine & vbNewLine
    helpText = helpText & "COMMANDS:" & vbNewLine & vbNewLine
    helpText = helpText & Chr(149) & " Ask AI - Chat with AI about your data" & vbNewLine
    helpText = helpText & Chr(149) & " Analyze - Get insights from selected data" & vbNewLine
    helpText = helpText & Chr(149) & " Fix Formula - Repair broken formulas" & vbNewLine
    helpText = helpText & Chr(149) & " Price Calc - Price calculation help" & vbNewLine
    helpText = helpText & Chr(149) & " Verify Formula - Check if formula works" & vbNewLine
    helpText = helpText & Chr(149) & " Create Table - Format data as table" & vbNewLine
    helpText = helpText & Chr(149) & " Formula Help - Get formula suggestions" & vbNewLine & vbNewLine
    helpText = helpText & "HOW TO USE:" & vbNewLine & vbNewLine
    helpText = helpText & "1. Select data in Excel" & vbNewLine
    helpText = helpText & "2. Click a button on the toolbar" & vbNewLine
    helpText = helpText & "3. Or press Alt+F8 and choose a command" & vbNewLine & vbNewLine
    helpText = helpText & "SUPPORT:" & vbNewLine & vbNewLine
    helpText = helpText & "Email: support@exceller.app" & vbNewLine
    helpText = helpText & "Web: https://exceller.app/help"
    
    MsgBox helpText, vbInformation, APP_NAME & " Help"
End Sub

Private Sub ShowAbout()
    Dim aboutText As String
    
    aboutText = APP_NAME & vbNewLine & _
                "Version " & APP_VERSION & vbNewLine & vbNewLine & _
                "AI-Powered Excel Assistant" & vbNewLine & vbNewLine & _
                "Features:" & vbNewLine & _
                Chr(149) & " Natural language data analysis" & vbNewLine & _
                Chr(149) & " Formula creation and debugging" & vbNewLine & _
                Chr(149) & " Smart table formatting" & vbNewLine & _
                Chr(149) & " Price calculation helpers" & vbNewLine & _
                Chr(149) & " Data quality checks" & vbNewLine & vbNewLine & _
                "Powered by Google Gemini AI" & vbNewLine & vbNewLine & _
                Chr(169) & " 2024 Exceller. Free to use."
    
    MsgBox aboutText, vbInformation, "About " & APP_NAME
End Sub

'==============================================================================
' UNINSTALL
'==============================================================================

Private Sub UninstallAddin()
    Dim response As VbMsgBoxResult
    
    response = MsgBox("Are you sure you want to uninstall " & APP_NAME & "?" & vbNewLine & vbNewLine & _
                      "This will:" & vbNewLine & _
                      Chr(149) & " Remove the toolbar" & vbNewLine & _
                      Chr(149) & " Delete saved settings" & vbNewLine & _
                      Chr(149) & " Keep your Excel files unchanged", _
                      vbYesNo + vbQuestion, APP_NAME & " - Uninstall")
    
    If response = vbYes Then
        ' Delete toolbar
        On Error Resume Next
        Application.CommandBars(APP_NAME).Delete
        On Error GoTo 0
        
        ' Delete settings
        DeleteSetting APP_NAME
        
        ' Remove from add-ins
        Dim addin As AddIn
        For Each addin In Application.AddIns
            If InStr(addin.Name, "Exceller") > 0 Then
                addin.Installed = False
                Exit For
            End If
        Next addin
        
        MsgBox APP_NAME & " has been uninstalled." & vbNewLine & vbNewLine & _
               "To completely remove:" & vbNewLine & _
               "1. Close Excel" & vbNewLine & _
               "2. Delete the .xlam file from your AddIns folder", _
               vbInformation, APP_NAME
    End If
End Sub
