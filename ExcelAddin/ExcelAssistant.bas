Attribute VB_Name = "ExcelAssistant"
'================================================================================
' Exceller - AI Excel Assistant Add-in
' Version: 2.0
' Use: Paste this into Excel VBA, save as .xlam
'================================================================================

Option Explicit

'API Configuration
' Leave empty - set your own key via Exceller_Settings (Alt+F8).
Private Const API_KEY As String = ""
Private Const API_URL As String = "https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent"

'================================================================================
' PUBLIC MACROS - These appear in Excel
'================================================================================

Public Sub Exceller_Chat()
    Call ShowAssistant("chat")
End Sub

Public Sub Exceller_Analyze()
    Call ShowAssistant("analyze")
End Sub

Public Sub Exceller_FixFormula()
    Call ShowAssistant("fix")
End Sub

Public Sub Exceller_VerifyFormula()
    Call ShowAssistant("verify")
End Sub

Public Sub Exceller_CreateTable()
    Call ShowAssistant("table")
End Sub

Public Sub Exceller_SuggestPrice()
    Call ShowAssistant("price")
End Sub

Public Sub Exceller_GetFormulaHelp()
    Call ShowAssistant("formula")
End Sub

Public Sub Exceller_Settings()
    Call ShowAssistant("settings")
End Sub

'================================================================================
' MAIN ASSISTANT
'================================================================================

Private Sub ShowAssistant(mode As String)
    Dim ws As Worksheet
    Dim userInput As String
    Dim context As String
    Dim response As String
    
    'Get context from selection
    context = GetSelectedContext()
    
    Select Case mode
        Case "chat"
            userInput = InputBox("Ask Exceller anything:" & vbNewLine & vbNewLine & _
                               "Examples:" & vbNewLine & _
                               "- Summarize this data" & vbNewLine & _
                               "- Find duplicates" & vbNewLine & _
                               "- Create a chart" & vbNewLine & _
                               "- Write a formula", _
                               "Exceller - AI Chat")
            
        Case "analyze"
            If TypeName(Selection) <> "Range" Then
                MsgBox "Select data first!", vbExclamation, "Exceller"
                Exit Sub
            End If
            userInput = "Analyze this data and provide insights, statistics, and recommendations"
            
        Case "fix"
            userInput = InputBox("Enter the broken formula and error:" & vbNewLine & vbNewLine & _
                               "Example: =VLOOKUP(A1,Sheet2!A:B,2,FALSE) - Error: #REF!", _
                               "Exceller - Fix Formula")
            context = "FORMULA FIX REQUEST: " & context
            
        Case "verify"
            If TypeName(Selection) = "Range" Then
                userInput = "Verify this formula is correct: " & Selection.Formula
            Else
                userInput = InputBox("Enter formula to verify:", "Exceller - Verify")
            End If
            
        Case "table"
            If TypeName(Selection) <> "Range" Then
                MsgBox "Select data first!", vbExclamation, "Exceller"
                Exit Sub
            End If
            userInput = "Create a professional table from this data. Suggest headers, formatting, and any calculations to add"
            
        Case "price"
            userInput = InputBox("Describe your price calculation:" & vbNewLine & vbNewLine & _
                               "Example: Calculate total with 10% discount and 8% tax", _
                               "Exceller - Price Calculator")
            context = "PRICE CALCULATION: " & context
            
        Case "formula"
            userInput = InputBox("What do you want to calculate?" & vbNewLine & vbNewLine & _
                               "Example: Sum values if date is in January", _
                               "Exceller - Formula Help")
            
        Case "settings"
            Dim newKey As String
            newKey = InputBox("Enter your Gemini API Key:" & vbNewLine & vbNewLine & _
                            "Get key at: https://makersuite.google.com/app/apikey", _
                            "Exceller - Settings", API_KEY)
            If Len(newKey) > 0 Then
                SaveSetting "Exceller", "Config", "ApiKey", newKey
                MsgBox "API key saved!", vbInformation, "Exceller"
            End If
            Exit Sub
            
        Case Else
            userInput = InputBox("Ask Exceller:", "Exceller")
    End Select
    
    'Exit if no input
    If Len(userInput) = 0 Then Exit Sub
    
    'Show processing
    Application.StatusBar = "Exceller is thinking..."
    Application.ScreenUpdating = False
    
    'Call API
    response = CallGeminiAPI(userInput, context)
    
    'Reset Excel
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    'Show response
    Call ShowResponse(response, mode)
End Sub

'================================================================================
' GEMINI API
'================================================================================

Private Function CallGeminiAPI(prompt As String, context As String) As String
    Dim httpReq As Object
    Dim requestBody As String
    Dim fullPrompt As String
    Dim apiKey As String
    
    'Get API key from settings
    apiKey = GetSetting("Exceller", "Config", "ApiKey", API_KEY)
    
    'Build prompt
    If Len(context) > 0 Then
        fullPrompt = "Context: " & context & vbNewLine & vbNewLine & "Question: " & prompt
    Else
        fullPrompt = prompt
    End If
    
    'Create HTTP request
    Set httpReq = CreateObject("MSXML2.XMLHTTP")
    
    'Build JSON body
    requestBody = "{""contents"":[{""parts"":[{""text"":""" & _
                  EscapeJson(fullPrompt) & """}]}],""generationConfig"":{""maxOutputTokens"":2048}}"
    
    'Send request
    On Error GoTo ErrorHandler
    httpReq.Open "POST", API_URL & "?key=" & apiKey, False
    httpReq.setRequestHeader "Content-Type", "application/json"
    httpReq.send requestBody
    
    'Check response
    If httpReq.Status = 200 Then
        CallGeminiAPI = ExtractText(httpReq.responseText)
    Else
        CallGeminiAPI = "Error: API returned status " & httpReq.Status
    End If
    
    Set httpReq = Nothing
    Exit Function
    
ErrorHandler:
    CallGeminiAPI = "Error: " & Err.Description
    Set httpReq = Nothing
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

Private Function ExtractText(jsonResponse As String) As String
    Dim startPos As Long
    Dim endPos As Long
    
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
            ExtractText = text
        Else
            ExtractText = "Error parsing response"
        End If
    Else
        ExtractText = "No response received"
    End If
End Function

'================================================================================
' CONTEXT HELPERS
'================================================================================

Private Function GetSelectedContext() As String
    If TypeName(Selection) = "Range" Then
        Dim rng As Range
        Set rng = Selection
        
        GetSelectedContext = "Range: " & rng.Address & vbNewLine & _
                            "Sheet: " & rng.Worksheet.Name & vbNewLine & _
                            "Size: " & rng.Rows.Count & " rows x " & rng.Columns.Count & " columns" & vbNewLine & _
                            "Data: " & GetRangeText(rng, 20)
    Else
        GetSelectedContext = "No data selected"
    End If
End Function

Private Function GetRangeText(rng As Range, maxRows As Long) As String
    Dim result As String
    Dim r As Long, c As Long
    Dim rowCount As Long
    
    rowCount = Application.Min(rng.Rows.Count, maxRows)
    
    'Headers
    For c = 1 To rng.Columns.Count
        If c > 1 Then result = result & " | "
        result = result & rng.Cells(1, c).Value
    Next c
    result = result & vbNewLine
    
    'Data
    For r = 2 To rowCount
        For c = 1 To rng.Columns.Count
            If c > 1 Then result = result & " | "
            result = result & rng.Cells(r, c).Value
        Next c
        result = result & vbNewLine
    Next r
    
    If rng.Rows.Count > maxRows Then
        result = result & "... (" & (rng.Rows.Count - maxRows) & " more rows)"
    End If
    
    GetRangeText = result
End Function

'================================================================================
' RESPONSE DISPLAY
'================================================================================

Private Sub ShowResponse(response As String, mode As String)
    Dim ws As Worksheet
    Dim cell As Range
    
    'Create new sheet for response
    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    ws.Name = "Exceller " & Format(Now, "hhmmss")
    
    'Add title
    ws.Range("A1").Value = "Exceller Response"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Color = RGB(76, 175, 80)
    
    'Add mode indicator
    ws.Range("A2").Value = "Mode: " & UCase(mode)
    ws.Range("A2").Font.Size = 10
    ws.Range("A2").Font.Color = RGB(128, 128, 128)
    
    'Add response
    ws.Range("A4").Value = response
    ws.Range("A4").WrapText = True
    ws.Range("A4").ColumnWidth = 80
    ws.Columns("A").AutoFit
    
    'Select response
    ws.Range("A4").Select
    
    'Show in message box too
    MsgBox response, vbInformation, "Exceller"
End Sub

'================================================================================
' AUTO-RUN (When add-in loads)
'================================================================================

Private Sub Auto_Open()
    'Runs when Excel opens with this add-in
    Call InitializeSettings
    
    'Show welcome (optional - remove next line to silent start)
    ' MsgBox "Exceller is loaded!" & vbNewLine & "Use Alt+F8 to see all commands.", vbInformation, "Exceller"
End Sub

Private Sub Auto_Close()
    'Runs when Excel closes
    'Clean up if needed
End Sub

Private Sub InitializeSettings()
    'Set default API key only if one is configured in code
    If Len(API_KEY) > 0 Then
        If Len(GetSetting("Exceller", "Config", "ApiKey", "")) = 0 Then
            SaveSetting "Exceller", "Config", "ApiKey", API_KEY
        End If
    End If
End Sub
