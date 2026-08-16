Attribute VB_Name = "ExcelAssistant"
'Excel Assistant - Gemini AI Powered
'Version: 1.0

Option Explicit

'API Configuration
' Leave empty - set your own key via the Settings button / ShowSettings.
Private Const GEMINI_API_KEY As String = ""
Private Const GEMINI_API_URL As String = "https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent"
Private Const MAX_TOKENS As Long = 2048

'Global Variables
Private assistantPane As Object
Private isInitialized As Boolean

'================================================================================
'INITIALIZATION
'================================================================================

Public Sub SetupAssistant()
    'Initializes the Excel Assistant add-in
    
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Setting up Excel Assistant..."
    
    'Create ribbon button
    Call CreateRibbonInterface
    
    'Initialize settings
    Call InitializeSettings
    
    'Show welcome message
    MsgBox "Excel Assistant is ready!" & vbNewLine & vbNewLine & _
           "Click 'AI Assistant' in the ribbon to get started." & vbNewLine & vbNewLine & _
           "You can:" & vbNewLine & _
           "- Ask questions about your data" & vbNewLine & _
           "- Analyze selected cells" & vbNewLine & _
           "- Create tables and charts" & vbNewLine & _
           "- Get formula help" & vbNewLine & _
           "- Clean and transform data", vbInformation, "Excel Assistant"
    
    Application.StatusBar = False
    Exit Sub
    
ErrorHandler:
    MsgBox "Error setting up assistant: " & Err.Description, vbCritical, "Error"
    Application.StatusBar = False
End Sub

Private Sub InitializeSettings()
    'Initialize default settings
    
    'Store API key only if one is configured in code
    If Len(GEMINI_API_KEY) > 0 Then
        SaveSetting "ExcelAssistant", "Config", "ApiKey", GEMINI_API_KEY
    End If
    
    'Set default preferences
    SaveSetting "ExcelAssistant", "Config", "AutoAnalyze", "True"
    SaveSetting "ExcelAssistant", "Config", "MaxTokens", CStr(MAX_TOKENS)
    
    isInitialized = True
End Sub

Private Sub Auto_Open()
    'Runs automatically when the add-in loads - no manual setup needed
    On Error Resume Next
    Call InitializeSettings
    Call CreateRibbonInterface
    On Error GoTo 0
End Sub

Private Sub CreateRibbonInterface()
    'Creates custom ribbon interface (requires ribbon XML)
    'Note: This is a simplified version - full ribbon requires XML customization
    
    'For now, we'll use the Quick Access Toolbar approach
    'Users can run macros directly or create custom toolbar
    
    Dim cmdBar As CommandBar
    Dim cmdBtn As CommandBarButton
    
    'Create custom toolbar
    On Error Resume Next
    Application.CommandBars("Excel Assistant").Delete
    On Error GoTo 0
    
    Set cmdBar = Application.CommandBars.Add(Name:="Excel Assistant", _
                                              Position:=msoBarTop, _
                                              Temporary:=True)
    
    'Add buttons
    Set cmdBtn = cmdBar.Controls.Add(Type:=msoControlButton)
    With cmdBtn
        .Caption = "AI Chat"
        .OnAction = "=ShowChat()"
        .FaceId = 463
        .TooltipText = "Open AI Chat Assistant"
    End With
    
    Set cmdBtn = cmdBar.Controls.Add(Type:=msoControlButton)
    With cmdBtn
        .Caption = "Analyze"
        .OnAction = "=AnalyzeSelection()"
        .FaceId = 295
        .TooltipText = "Analyze selected data"
    End With
    
    Set cmdBtn = cmdBar.Controls.Add(Type:=msoControlButton)
    With cmdBtn
        .Caption = "Create Table"
        .OnAction = "=CreateSmartTable()"
        .FaceId = 293
        .TooltipText = "Create formatted table from data"
    End With
    
    Set cmdBtn = cmdBar.Controls.Add(Type:=msoControlButton)
    With cmdBtn
        .Caption = "Check Files"
        .OnAction = "=CheckFiles()"
        .FaceId = 2072
        .TooltipText = "Analyze files in folder"
    End With
    
    Set cmdBtn = cmdBar.Controls.Add(Type:=msoControlButton)
    With cmdBtn
        .Caption = "Settings"
        .OnAction = "=ShowSettings()"
        .FaceId = 925
        .TooltipText = "Configure Excel Assistant"
    End With
    
    cmdBar.Visible = True
End Sub

'================================================================================
'AI COMMUNICATION
'================================================================================

Public Function CallGeminiAPI(prompt As String, Optional context As String = "") As String
    'Calls Gemini API with the given prompt
    
    Dim httpReq As Object
    Dim requestBody As String
    Dim response As String
    Dim fullPrompt As String
    
    On Error GoTo ErrorHandler
    
    'Build prompt with context
    If Len(context) > 0 Then
        fullPrompt = "Context: " & context & vbNewLine & vbNewLine & "Question: " & prompt
    Else
        fullPrompt = prompt
    End If
    
    'Get API key from settings (set via ShowSettings)
    Dim apiKey As String
    apiKey = GetSetting("ExcelAssistant", "Config", "ApiKey", GEMINI_API_KEY)
    
    'Create HTTP request
    Set httpReq = CreateObject("MSXML2.XMLHTTP")
    
    'Build request body
    requestBody = "{""contents"":[{""parts"":[{""text"":""" & EscapeJson(fullPrompt) & """}]}],""generationConfig"":{""maxOutputTokens"":" & MAX_TOKENS & "}}"
    
    'Send request
    httpReq.Open "POST", GEMINI_API_URL & "?key=" & apiKey, False
    httpReq.setRequestHeader "Content-Type", "application/json"
    httpReq.send requestBody
    
    'Check response
    If httpReq.Status = 200 Then
        response = httpReq.responseText
        CallGeminiAPI = ExtractResponseText(response)
    Else
        CallGeminiAPI = "Error: API request failed with status " & httpReq.Status
    End If
    
    Set httpReq = Nothing
    Exit Function
    
ErrorHandler:
    CallGeminiAPI = "Error: " & Err.Description
    Set httpReq = Nothing
End Function

Private Function ExtractResponseText(jsonResponse As String) As String
    'Extracts text from Gemini API JSON response
    
    Dim startPos As Long
    Dim endPos As Long
    Dim text As String
    
    'Find the text field in response
    startPos = InStr(jsonResponse, """text"":""")
    If startPos > 0 Then
        startPos = startPos + 8 'Skip past ""text"":"
        endPos = InStr(startPos, jsonResponse, """")
        If endPos > startPos Then
            text = Mid(jsonResponse, startPos, endPos - startPos)
            'Unescape JSON string
            text = Replace(text, "\n", vbNewLine)
            text = Replace(text, "\""", """")
            text = Replace(text, "\\", "\")
            ExtractResponseText = text
        Else
            ExtractResponseText = "Error: Could not parse response"
        End If
    Else
        ExtractResponseText = "Error: No text in response"
    End If
End Function

Private Function EscapeJson(str As String) As String
    'Escapes special characters for JSON
    
    Dim result As String
    Dim i As Long
    
    result = str
    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbNewLine, "\n")
    result = Replace(result, vbCr, "\n")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbTab, "\t")
    
    EscapeJson = result
End Function

'================================================================================
'UI FUNCTIONS
'================================================================================

Public Function ShowChat() As String
    'Shows the AI chat interface
    
    Dim userPrompt As String
    Dim context As String
    Dim response As String
    
    'Get context from selection
    context = GetDataContext()
    
    'Show input dialog
    userPrompt = InputBox("Ask Excel Assistant anything:" & vbNewLine & vbNewLine & _
                          "Examples:" & vbNewLine & _
                          "- Summarize this data" & vbNewLine & _
                          "- Find duplicates" & vbNewLine & _
                          "- Create a chart" & vbNewLine & _
                          "- Write a SUMIF formula" & vbNewLine & _
                          "- Clean this data", _
                          "Excel Assistant - AI Chat")
    
    If Len(userPrompt) > 0 Then
        Application.StatusBar = "Thinking..."
        response = CallGeminiAPI(userPrompt, context)
        Application.StatusBar = False
        
        'Display response
        ShowResponse response
    End If
    
    ShowChat = response
End Function

Public Function AnalyzeSelection() As String
    'Analyzes the selected data range
    
    Dim dataRange As Range
    Dim dataContext As String
    Dim prompt As String
    Dim response As String
    
    'Check if selection exists
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range of data first.", vbExclamation, "Excel Assistant"
        AnalyzeSelection = ""
        Exit Function
    End If
    
    Set dataRange = Selection
    
    'Build context
    dataContext = "Analyze this Excel data:" & vbNewLine & vbNewLine
    dataContext = dataContext & "Range: " & dataRange.Address & vbNewLine
    dataContext = dataContext & "Dimensions: " & dataRange.Rows.Count & " rows x " & dataRange.Columns.Count & " columns" & vbNewLine
    dataContext = dataContext & "Data:" & vbNewLine & GetDataAsString(dataRange)
    
    'Create analysis prompt
    prompt = "Please analyze this data and provide:" & vbNewLine & _
             "1. Summary statistics" & vbNewLine & _
             "2. Key insights" & vbNewLine & _
             "3. Data quality observations" & vbNewLine & _
             "4. Recommendations" & vbNewLine & vbNewLine & _
             "Format your response clearly with headers."
    
    Application.StatusBar = "Analyzing data..."
    response = CallGeminiAPI(prompt, dataContext)
    Application.StatusBar = False
    
    'Display results
    ShowResponse response
    
    AnalyzeSelection = response
End Function

Public Function CreateSmartTable() As String
    'Creates a smart table from selected data
    
    Dim dataRange As Range
    Dim prompt As String
    Dim context As String
    Dim response As String
    
    'Check selection
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select data first.", vbExclamation, "Excel Assistant"
        CreateSmartTable = ""
        Exit Function
    End If
    
    Set dataRange = Selection
    
    'Build context
    context = "Data to format as table:" & vbNewLine & _
              GetDataAsString(dataRange)
    
    'Ask for table creation help
    prompt = "Help me create a professional table from this data:" & vbNewLine & _
             "- What headers should I use?" & vbNewLine & _
             "- How should I format it?" & vbNewLine & _
             "- Any sorting or grouping suggestions?" & vbNewLine & _
             "- What calculations should I add?"
    
    Application.StatusBar = "Creating table..."
    response = CallGeminiAPI(prompt, context)
    Application.StatusBar = False
    
    'Create actual table
    Call FormatAsTable(dataRange)
    
    ShowResponse response
    
    CreateSmartTable = response
End Function

Public Function CheckFiles() As String
    'Analyzes files in a selected folder
    
    Dim folderPath As String
    Dim fileName As String
    Dim fileList As String
    Dim fileContent As String
    Dim response As String
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    
    'Get folder from user
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select folder to analyze"
        .AllowMultiSelect = False
        If .Show = -1 Then
            folderPath = .SelectedItems(1)
        Else
            CheckFiles = ""
            Exit Function
        End If
    End With
    
    'Create file system object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'Check if folder exists
    If Not fso.FolderExists(folderPath) Then
        MsgBox "Folder not found.", vbExclamation, "Excel Assistant"
        CheckFiles = ""
        Exit Function
    End If
    
    Set folder = fso.GetFolder(folderPath)
    
    'List files
    fileList = "Files in " & folderPath & ":" & vbNewLine & vbNewLine
    
    For Each file In folder.Files
        fileList = fileList & "- " & file.Name & " (" & FormatFileSize(file.Size) & ")" & vbNewLine
    Next file
    
    'Analyze file content (sample)
    fileContent = ""
    For Each file In folder.Files
        If LCase(Right(file.Name, 4)) = ".csv" Or LCase(Right(file.Name, 4)) = ".txt" Then
            fileContent = fileContent & vbNewLine & "File: " & file.Name & vbNewLine
            fileContent = fileContent & ReadFileSample(file.Path, 20) & vbNewLine
        End If
    Next file
    
    'Get analysis
    Dim prompt As String
    prompt = "Analyze these files and provide:" & vbNewLine & _
             "1. File summary" & vbNewLine & _
             "2. Data quality assessment" & vbNewLine & _
             "3. Recommendations for processing" & vbNewLine & _
             "4. Any issues detected"
    
    Application.StatusBar = "Analyzing files..."
    response = CallGeminiAPI(prompt, fileList & vbNewLine & "Sample content:" & vbNewLine & fileContent)
    Application.StatusBar = False
    
    ShowResponse response
    
    CheckFiles = response
End Function

'================================================================================
'HELPER FUNCTIONS
'================================================================================

Private Function GetDataContext() As String
    'Gets context from current selection
    
    Dim ctx As String
    
    If TypeName(Selection) = "Range" Then
        ctx = "Current selection: " & Selection.Address & vbNewLine
        ctx = ctx & "Dimensions: " & Selection.Rows.Count & " rows x " & Selection.Columns.Count & " columns" & vbNewLine
        ctx = ctx & "Data sample: " & GetDataAsString(Selection, 10)
    Else
        ctx = "No specific data selected."
    End If
    
    GetDataContext = ctx
End Function

Private Function GetDataAsString(rng As Range, Optional maxRows As Long = 20) As String
    'Converts range data to string representation
    
    Dim result As String
    Dim row As Long, col As Long
    Dim rowCount As Long, colCount As Long
    
    rowCount = Application.WorksheetFunction.Min(rng.Rows.Count, maxRows)
    colCount = rng.Columns.Count
    
    'Add headers if first row looks like headers
    For col = 1 To colCount
        If col > 1 Then result = result & vbTab
        result = result & rng.Cells(1, col).Value
    Next col
    result = result & vbNewLine
    
    'Add data rows
    For row = 2 To rowCount
        For col = 1 To colCount
            If col > 1 Then result = result & vbTab
            result = result & rng.Cells(row, col).Value
        Next col
        result = result & vbNewLine
    Next row
    
    'Add note if truncated
    If rng.Rows.Count > maxRows Then
        result = result & "... (" & (rng.Rows.Count - maxRows) & " more rows)" & vbNewLine
    End If
    
    GetDataAsString = result
End Function

Private Sub FormatAsTable(rng As Range)
    'Formats range as Excel table
    
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set ws = rng.Worksheet
    
    'Remove existing table if any
    For Each tbl In ws.ListObjects
        If Not Intersect(tbl.Range, rng) Is Nothing Then
            tbl.Unlist
            Exit For
        End If
    Next tbl
    
    'Create new table
    Set tbl = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    tbl.Name = "Table_" & Format(Now, "hhmmss")
    tbl.TableStyle = "TableStyleMedium9"
    
    'Auto-fit columns
    tbl.Range.Columns.AutoFit
    
    On Error GoTo 0
End Sub

Private Function ReadFileSample(filePath As String, maxLines As Long) As String
    'Reads first N lines from a file
    
    Dim fso As Object
    Dim file As Object
    Dim line As String
    Dim content As String
    Dim lineCount As Long
    
    On Error Resume Next
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.OpenTextFile(filePath, 1)
    
    If Err.Number = 0 Then
        lineCount = 0
        Do While Not file.AtEndOfStream And lineCount < maxLines
            line = file.ReadLine
            content = content & line & vbNewLine
            lineCount = lineCount + 1
        Loop
        file.Close
    End If
    
    ReadFileSample = content
    
    On Error GoTo 0
End Function

Private Function FormatFileSize(bytes As Long) As String
    'Formats file size to human-readable string
    
    If bytes < 1024 Then
        FormatFileSize = bytes & " B"
    ElseIf bytes < 1024 * 1024 Then
        FormatFileSize = Format(bytes / 1024, "#,##0") & " KB"
    Else
        FormatFileSize = Format(bytes / (1024 * 1024), "#,##0.0") & " MB"
    End If
End Function

Private Sub ShowResponse(response As String)
    'Displays AI response in a new worksheet
    
    Dim ws As Worksheet
    Dim cell As Range
    
    'Create new sheet
    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    ws.Name = "AI Response " & Format(Now, "hhmmss")
    
    'Add response
    ws.Range("A1").Value = "Excel Assistant Response"
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True
    
    ws.Range("A3").Value = response
    ws.Range("A3").WrapText = True
    ws.Range("A3").ColumnWidth = 80
    
    'Auto-fit
    ws.Columns("A").AutoFit
    
    'Select the response
    ws.Range("A3").Select
End Sub

'================================================================================
'FORMULA HELPER
'================================================================================

Public Function GetFormulaHelp() As String
    'Gets help with Excel formulas
    
    Dim userQuestion As String
    Dim response As String
    
    userQuestion = InputBox("Describe what you want to calculate:" & vbNewLine & vbNewLine & _
                           "Examples:" & vbNewLine & _
                           "- Sum values if date is in January" & vbNewLine & _
                           "- Count unique items in a list" & vbNewLine & _
                           "- Look up value in another table" & vbNewLine & _
                           "- Calculate percentage change", _
                           "Formula Helper")
    
    If Len(userQuestion) > 0 Then
        Dim prompt As String
        prompt = "Help me write an Excel formula for: " & userQuestion & vbNewLine & vbNewLine & _
                 "Please provide:" & vbNewLine & _
                 "1. The formula" & vbNewLine & _
                 "2. Explanation of how it works" & vbNewLine & _
                 "3. Example use case" & vbNewLine & _
                 "4. Any alternatives"
        
        Application.StatusBar = "Getting formula help..."
        response = CallGeminiAPI(prompt)
        Application.StatusBar = False
        
        ShowResponse response
        GetFormulaHelp = response
    End If
End Function

'================================================================================
'DATA CLEANING
'================================================================================

Public Function CleanData() As String
    'Helps clean and transform data
    
    Dim dataRange As Range
    Dim context As String
    Dim prompt As String
    Dim response As String
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select data to clean.", vbExclamation, "Excel Assistant"
        CleanData = ""
        Exit Function
    End If
    
    Set dataRange = Selection
    
    context = "Data to clean:" & vbNewLine & GetDataAsString(dataRange, 30)
    
    prompt = "Analyze this data and help me clean it:" & vbNewLine & _
             "- Identify issues (missing values, duplicates, formatting)" & vbNewLine & _
             "- Suggest VBA code or formulas to fix issues" & vbNewLine & _
             "- Provide step-by-step cleaning instructions"
    
    Application.StatusBar = "Analyzing data quality..."
    response = CallGeminiAPI(prompt, context)
    Application.StatusBar = False
    
    ShowResponse response
    CleanData = response
End Function

'================================================================================
'CHART HELPER
'================================================================================

Public Function SuggestChart() As String
    'Suggests and creates charts
    
    Dim dataRange As Range
    Dim context As String
    Dim prompt As String
    Dim response As String
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select data for chart.", vbExclamation, "Excel Assistant"
        SuggestChart = ""
        Exit Function
    End If
    
    Set dataRange = Selection
    
    context = "Data for chart:" & vbNewLine & GetDataAsString(dataRange, 20)
    
    prompt = "Suggest the best chart type for this data and explain why:" & vbNewLine & _
             "Also provide VBA code to create the chart."
    
    Application.StatusBar = "Suggesting chart..."
    response = CallGeminiAPI(prompt, context)
    Application.StatusBar = False
    
    'Optionally create chart
    If MsgBox("Would you like me to create a chart based on the suggestions?", _
              vbYesNo + vbQuestion, "Create Chart") = vbYes Then
        Call CreateChartFromData(dataRange)
    End If
    
    ShowResponse response
    SuggestChart = response
End Function

Private Sub CreateChartFromData(rng As Range)
    'Creates a chart from selected data
    
    Dim ws As Worksheet
    Dim cht As ChartObject
    
    Set ws = rng.Worksheet
    
    'Create chart
    Set cht = ws.ChartObjects.Add( _
        Left:=ws.Cells(rng.Row + rng.Rows.Count + 2, rng.Column).Left, _
        Top:=ws.Cells(rng.Row, rng.Column + rng.Columns.Count + 2).Top, _
        Width:=400, Height:=300)
    
    'Configure chart
    With cht.Chart
        .SetSourceData Source:=rng
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Data Visualization"
        .HasLegend = rng.Columns.Count > 1
    End With
End Sub

'================================================================================
'SETTINGS
'================================================================================

Public Function ShowSettings() As String
    'Shows settings dialog
    
    Dim currentKey As String
    Dim newKey As String
    
    currentKey = GetSetting("ExcelAssistant", "Config", "ApiKey", "")
    
    newKey = InputBox("Enter your Gemini API Key:" & vbNewLine & vbNewLine & _
                     "Current key: " & IIf(Len(currentKey) > 0, Left(currentKey, 8) & "...", "Not set") & vbNewLine & vbNewLine & _
                     "Get your key at: https://makersuite.google.com/app/apikey", _
                     "Settings - API Key", currentKey)
    
    If Len(newKey) > 0 Then
        SaveSetting "ExcelAssistant", "Config", "ApiKey", newKey
        MsgBox "API key saved successfully!", vbInformation, "Settings"
    End If
    
    ShowSettings = "Settings updated"
End Function

'================================================================================
'QUICK ACTIONS
'================================================================================

Public Function QuickSummarize() As String
    'Quick summary of selected data
    
    If TypeName(Selection) = "Range" Then
        Dim prompt As String
        prompt = "Provide a brief summary of this data in 3-5 bullet points:" & vbNewLine & _
                 GetDataAsString(Selection, 15)
        
        QuickSummarize = CallGeminiAPI(prompt)
        ShowResponse QuickSummarize
    End If
End Function

Public Function FindDuplicates() As String
    'Finds and highlights duplicates
    
    Dim rng As Range
    Dim cell As Range
    Dim dict As Object
    Dim dupCount As Long
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select data.", vbExclamation, "Excel Assistant"
        FindDuplicates = ""
        Exit Function
    End If
    
    Set rng = Selection
    Set dict = CreateObject("Scripting.Dictionary")
    dupCount = 0
    
    'Find duplicates
    For Each cell In rng
        If cell.Value <> "" Then
            If dict.Exists(cell.Value) Then
                dict(cell.Value) = dict(cell.Value) + 1
                cell.Interior.Color = RGB(255, 200, 200) 'Highlight duplicate
                dupCount = dupCount + 1
            Else
                dict.Add cell.Value, 1
            End If
        End If
    Next cell
    
    FindDuplicates = "Found " & dupCount & " duplicate values."
    MsgBox FindDuplicates, vbInformation, "Duplicates Found"
End Function

Public Function GetStatistics() As String
    'Gets statistics for numeric data
    
    Dim rng As Range
    Dim stats As String
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select numeric data.", vbExclamation, "Excel Assistant"
        GetStatistics = ""
        Exit Function
    End If
    
    Set rng = Selection
    
    stats = "Statistics:" & vbNewLine & _
            "Count: " & Application.WorksheetFunction.Count(rng) & vbNewLine & _
            "Sum: " & Application.WorksheetFunction.Sum(rng) & vbNewLine & _
            "Average: " & Format(Application.WorksheetFunction.Average(rng), "#,##0.00") & vbNewLine & _
            "Min: " & Application.WorksheetFunction.Min(rng) & vbNewLine & _
            "Max: " & Application.WorksheetFunction.Max(rng) & vbNewLine & _
            "Std Dev: " & Format(Application.WorksheetFunction.StDev(rng), "#,##0.00")
    
    ShowResponse stats
    GetStatistics = stats
End Function

'================================================================================
'EXPORT/IMPORT
'================================================================================

Public Function ExportToFile() As String
    'Exports selected data to a file
    
    Dim rng As Range
    Dim filePath As String
    Dim fileNum As Integer
    Dim row As Long, col As Long
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select data to export.", vbExclamation, "Excel Assistant"
        ExportToFile = ""
        Exit Function
    End If
    
    Set rng = Selection
    
    'Get save location
    With Application.FileDialog(msoFileDialogSaveAs)
        .Title = "Export Data"
        .FilterIndex = 1
        If .Show = -1 Then
            filePath = .SelectedItems(1)
        Else
            ExportToFile = ""
            Exit Function
        End If
    End With
    
    'Write to file
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    
    For row = 1 To rng.Rows.Count
        For col = 1 To rng.Columns.Count
            If col > 1 Then Print #fileNum, vbTab;
            Print #fileNum, rng.Cells(row, col).Value;
        Next col
        Print #fileNum,
    Next row
    
    Close #fileNum
    
    MsgBox "Data exported successfully!", vbInformation, "Export Complete"
    ExportToFile = "Exported to " & filePath
End Function
