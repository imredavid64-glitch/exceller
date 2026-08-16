VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAssistant 
   Caption         =   "Excel Assistant - AI Chat"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8400
   StartUpPosition =   1  'CenterOwner
   Begin MSForms.TextBox txtInput 
      Height          =   1200
      Left            =   120
      TabIndex        =   2
      Top             =   4680
      Width           =   8160
      WordWrap        =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
   End
   Begin MSForms.TextBox txtOutput 
      Height          =   4320
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   8160
      WordWrap        =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      Locked          =   -1  'True
   End
   Begin MSForms.CommandButton btnSend 
      Caption         =   "Send"
      Height          =   480
      Left            =   6960
      TabIndex        =   0
      Top             =   5520
      Width           =   1320
   End
   Begin MSForms.CommandButton btnClear 
      Caption         =   "Clear"
      Height          =   480
      Left            =   5520
      TabIndex        =   3
      Top             =   5520
      Width           =   1320
   End
   Begin MSForms.CommandButton btnAnalyze 
      Caption         =   "Analyze Selection"
      Height          =   480
      Left            =   120
      TabIndex        =   4
      Top             =   5520
      Width           =   1800
   End
   Begin MSForms.Label lblStatus 
      Caption         =   "Ready"
      Height          =   240
      Left            =   2040
      TabIndex        =   5
      Top             =   5640
      Width           =   3360
   End
End
Attribute VB_Name = "frmAssistant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Excel Assistant Chat Form

Option Explicit

Private conversationHistory As String

Private Sub UserForm_Initialize()
    'Initialize the form
    
    Me.Caption = "Excel Assistant - AI Chat"
    txtOutput.Value = "Welcome to Excel Assistant!" & vbNewLine & vbNewLine & _
                     "I can help you with:" & vbNewLine & _
                     "- Analyzing data" & vbNewLine & _
                     "- Creating tables" & vbNewLine & _
                     "- Writing formulas" & vbNewLine & _
                     "- Cleaning data" & vbNewLine & _
                     "- And more!" & vbNewLine & vbNewLine & _
                     "Type your question below or click 'Analyze Selection' to start."
    
    conversationHistory = ""
    lblStatus.Caption = "Ready"
End Sub

Private Sub btnSend_Click()
    'Send message to AI
    
    Dim userMessage As String
    Dim aiResponse As String
    Dim context As String
    
    userMessage = Trim(txtInput.Value)
    
    If Len(userMessage) = 0 Then
        MsgBox "Please enter a message.", vbExclamation, "Input Required"
        Exit Sub
    End If
    
    'Update UI
    lblStatus.Caption = "Thinking..."
    btnSend.Enabled = False
    Me.MousePointer = vbHourglass
    
    'Build context
    context = GetDataContext()
    
    'Add to conversation history
    conversationHistory = conversationHistory & "User: " & userMessage & vbNewLine & vbNewLine
    
    'Get AI response
    aiResponse = CallGeminiAPI(userMessage, context & vbNewLine & "Conversation:" & vbNewLine & conversationHistory)
    
    'Update conversation history
    conversationHistory = conversationHistory & "Assistant: " & aiResponse & vbNewLine & vbNewLine
    
    'Display response
    txtOutput.Value = txtOutput.Value & vbNewLine & vbNewLine & "---" & vbNewLine & _
                      "You: " & userMessage & vbNewLine & vbNewLine & _
                      "Assistant: " & aiResponse
    
    'Clear input and scroll to bottom
    txtInput.Value = ""
    txtOutput.SelStart = Len(txtOutput.Value)
    
    'Reset UI
    lblStatus.Caption = "Ready"
    btnSend.Enabled = True
    Me.MousePointer = vbDefault
End Sub

Private Sub btnAnalyze_Click()
    'Analyze selected data
    
    lblStatus.Caption = "Analyzing..."
    btnAnalyze.Enabled = False
    Me.MousePointer = vbHourglass
    
    Dim response As String
    response = AnalyzeSelection()
    
    If Len(response) > 0 Then
        txtOutput.Value = txtOutput.Value & vbNewLine & vbNewLine & "---" & vbNewLine & _
                         "Analysis Results:" & vbNewLine & response
        txtOutput.SelStart = Len(txtOutput.Value)
    End If
    
    lblStatus.Caption = "Ready"
    btnAnalyze.Enabled = True
    Me.MousePointer = vbDefault
End Sub

Private Sub btnClear_Click()
    'Clear conversation
    
    txtOutput.Value = "Conversation cleared. How can I help you?"
    conversationHistory = ""
    txtInput.Value = ""
End Sub

Private Sub txtInput_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    'Send on Enter key (without Shift)
    
    If KeyCode = vbKeyReturn And Shift = 0 Then
        KeyCode = 0
        btnSend_Click
    End If
End Sub

Private Function GetDataContext() As String
    'Gets context from Excel selection
    
    Dim ctx As String
    
    On Error Resume Next
    
    If TypeName(Excel.Selection) = "Range" Then
        Dim rng As Range
        Set rng = Excel.Selection
        
        ctx = "Current selection in Excel:" & vbNewLine
        ctx = ctx & "Range: " & rng.Address & vbNewLine
        ctx = ctx & "Sheet: " & rng.Worksheet.Name & vbNewLine
        ctx = ctx & "Dimensions: " & rng.Rows.Count & " rows x " & rng.Columns.Count & " columns" & vbNewLine
        ctx = ctx & "Data sample:" & vbNewLine
        
        'Add sample data
        Dim row As Long, col As Long
        Dim maxRows As Long
        maxRows = Application.WorksheetFunction.Min(rng.Rows.Count, 10)
        
        'Headers
        For col = 1 To rng.Columns.Count
            If col > 1 Then ctx = ctx & vbTab
            ctx = ctx & rng.Cells(1, col).Value
        Next col
        ctx = ctx & vbNewLine
        
        'Data
        For row = 2 To maxRows
            For col = 1 To rng.Columns.Count
                If col > 1 Then ctx = ctx & vbTab
                ctx = ctx & rng.Cells(row, col).Value
            Next col
            ctx = ctx & vbNewLine
        Next row
    Else
        ctx = "No specific data selected in Excel."
    End If
    
    On Error GoTo 0
    
    GetDataContext = ctx
End Function
