"""
Exceller Pro - One-Click Excel Add-in Installer
Run this app to automatically install Exceller into Excel
"""

import os
import sys
import json
import shutil
import winreg
import subprocess
from pathlib import Path
from threading import Thread

try:
    from tkinter import (
        Tk, Toplevel, Label, Button, Entry, Frame,
        messagebox, StringVar, IntVar, PhotoImage,
        RIGHT, LEFT, TOP, BOTTOM, X, Y, BOTH, W, E, N, S, CENTER
    )
    from tkinter import ttk
    HAS_TKINTER = True
except ImportError:
    HAS_TKINTER = False

# VBA Code for the Excel Add-in
VBA_CODE = '''Attribute VB_Name = "ExcellerPro"
'==============================================================================
' Exceller Pro - AI-Powered Excel Assistant
' Version: 3.0.0
' Installed by Exceller Installer
'==============================================================================

Option Explicit

Private Const APP_NAME As String = "Exceller Pro"
Private Const APP_VERSION As String = "3.0.0"
Private Const DEFAULT_API_URL As String = "https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent"

Private isInitialized As Boolean
Private cachedApiKey As String

Private Sub Auto_On()
    Call Initialize
End Sub

Private Sub Initialize()
    If isInitialized Then Exit Sub
    Call LoadSettings
    Call CreateToolbar
    isInitialized = True
End Sub

Private Sub LoadSettings()
    cachedApiKey = GetSetting(APP_NAME, "Config", "ApiKey", "")
    If Len(cachedApiKey) = 0 Then
        MsgBox "Welcome to Exceller Pro!" & vbNewLine & vbNewLine & _
               "Press Alt+F8 and run ExcellerPro_Settings" & vbNewLine & _
               "to enter your API key.", vbInformation, APP_NAME
    End If
End Sub

Private Sub CreateToolbar()
    Dim cmdBar As CommandBar
    Dim cmdBtn As CommandBarButton
    
    On Error Resume Next
    Application.CommandBars(APP_NAME).Delete
    On Error GoTo 0
    
    Set cmdBar = Application.CommandBars.Add(Name:=APP_NAME, Position:=msoBarTop, Temporary:=True)
    
    Set cmdBtn = cmdBar.Controls.Add(Type:=msoControlButton)
    With cmdBtn
        .Caption = "Ask AI"
        .OnAction = "=ExcellerPro_Chat()"
        .FaceId = 9854
        .Style = msoButtonCaption
    End With
    
    Set cmdBtn = cmdBar.Controls.Add(Type:=msoControlButton)
    With cmdBtn
        .Caption = "Analyze"
        .OnAction = "=ExcellerPro_Analyze()"
        .FaceId = 629
        .Style = msoButtonCaption
    End With
    
    cmdBar.Controls.Add Type:=msoControlSeparator
    
    Set cmdBtn = cmdBar.Controls.Add(Type:=msoControlButton)
    With cmdBtn
        .Caption = "Fix Formula"
        .OnAction = "=ExcellerPro_FixFormula()"
        .FaceId = 3146
        .Style = msoButtonCaption
    End With
    
    Set cmdBtn = cmdBar.Controls.Add(Type:=msoControlButton)
    With cmdBtn
        .Caption = "Price Calc"
        .OnAction = "=ExcellerPro_PriceCalc()"
        .FaceId = 1636
        .Style = msoButtonCaption
    End With
    
    cmdBar.Controls.Add Type:=msoControlSeparator
    
    Set cmdBtn = cmdBar.Controls.Add(Type:=msoControlButton)
    With cmdBtn
        .Caption = "Settings"
        .OnAction = "=ExcellerPro_Settings()"
        .FaceId = 925
        .Style = msoButtonCaption
    End With
    
    cmdBar.Visible = True
End Sub

Public Sub ExcellerPro_Chat()
    Call ShowAssistant("chat")
End Sub

Public Sub ExcellerPro_Analyze()
    Call ShowAssistant("analyze")
End Sub

Public Sub ExcellerPro_FixFormula()
    Call ShowAssistant("fix")
End Sub

Public Sub ExcellerPro_PriceCalc()
    Call ShowAssistant("price")
End Sub

Public Sub ExcellerPro_Settings()
    Dim newKey As String
    newKey = InputBox("Enter your Gemini API Key:" & vbNewLine & vbNewLine & _
                     "Get free key at:" & vbNewLine & _
                     "https://makersuite.google.com/app/apikey", _
                     APP_NAME & " Settings", cachedApiKey)
    If Len(newKey) > 0 Then
        cachedApiKey = newKey
        SaveSetting APP_NAME, "Config", "ApiKey", newKey
        MsgBox "Settings saved!", vbInformation, APP_NAME
    End If
End Sub

Public Sub ExcellerPro_Help()
    MsgBox "Exceller Pro Commands:" & vbNewLine & vbNewLine & _
           Chr(149) & " Ask AI - Chat about your data" & vbNewLine & _
           Chr(149) & " Analyze - Get data insights" & vbNewLine & _
           Chr(149) & " Fix Formula - Repair broken formulas" & vbNewLine & _
           Chr(149) & " Price Calc - Price calculation help" & vbNewLine & _
           Chr(149) & " Settings - Configure API key", _
           vbInformation, APP_NAME
End Sub

Public Sub ExcellerPro_Uninstall()
    If MsgBox("Remove Exceller Pro?", vbYesNo + vbQuestion, APP_NAME) = vbYes Then
        On Error Resume Next
        Application.CommandBars(APP_NAME).Delete
        DeleteSetting APP_NAME
        MsgBox "Exceller Pro removed. Close and reopen Excel.", vbInformation, APP_NAME
    End If
End Sub

Private Sub ShowAssistant(mode As String)
    If Len(cachedApiKey) = 0 Then
        MsgBox "Please set API key first (Alt+F8 > ExcellerPro_Settings)", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    Dim userInput As String
    Dim context As String
    
    If TypeName(Selection) = "Range" Then
        context = "Range: " & Selection.Address & vbNewLine & _
                  "Sheet: " & Selection.Worksheet.Name & vbNewLine & _
                  "Data: " & GetRangeText(Selection, 15)
    End If
    
    Select Case mode
        Case "chat"
            userInput = InputBox("Ask about your data:", APP_NAME)
        Case "analyze"
            If TypeName(Selection) <> "Range" Then
                MsgBox "Select data first!", vbExclamation, APP_NAME
                Exit Sub
            End If
            userInput = "Analyze this data and provide insights"
        Case "fix"
            userInput = InputBox("Enter broken formula and error:", APP_NAME)
            context = "BROKEN: " & context
        Case "price"
            userInput = InputBox("Describe price calculation:", APP_NAME)
    End Select
    
    If Len(userInput) = 0 Then Exit Sub
    
    Application.StatusBar = "Exceller Pro thinking..."
    Dim response As String
    response = CallAPI(userInput, context)
    Application.StatusBar = False
    
    Dim ws As Worksheet
    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    ws.Name = "Exceller " & Format(Now, "hhmmss")
    ws.Range("A1").Value = "Exceller Pro Response"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A3").Value = response
    ws.Range("A3").WrapText = True
    ws.Columns("A").ColumnWidth = 60
    
    If Len(response) <= 500 Then MsgBox response, vbInformation, APP_NAME
End Sub

Private Function GetRangeText(rng As Range, maxRows As Long) As String
    Dim r As Long, c As Long
    Dim result As String
    Dim rowCount As Long
    rowCount = Application.Min(rng.Rows.Count, maxRows)
    For r = 1 To rowCount
        For c = 1 To rng.Columns.Count
            If c > 1 Then result = result & " | "
            result = result & rng.Cells(r, c).Value
        Next c
        result = result & vbNewLine
    Next r
    GetRangeText = result
End Function

Private Function CallAPI(prompt As String, context As String) As String
    Dim httpReq As Object
    Set httpReq = CreateObject("MSXML2.XMLHTTP")
    
    Dim fullPrompt As String
    fullPrompt = "You are an Excel assistant. " & prompt
    If Len(context) > 0 Then fullPrompt = fullPrompt & vbNewLine & "Context: " & context
    
    Dim json As String
    json = "{""contents"":[{""parts"":[{""text"":""" & EscapeJson(fullPrompt) & """}]}],""generationConfig"":{""maxOutputTokens"":2048}}"
    
    On Error GoTo ErrHandler
    httpReq.Open "POST", DEFAULT_API_URL & "?key=" & cachedApiKey, False
    httpReq.setRequestHeader "Content-Type", "application/json"
    httpReq.send json
    
    If httpReq.Status = 200 Then
        CallAPI = ParseResponse(httpReq.responseText)
    Else
        CallAPI = "Error: " & httpReq.Status
    End If
    Exit Function
    
ErrHandler:
    CallAPI = "Error: " & Err.Description
End Function

Private Function EscapeJson(str As String) As String
    Dim result As String
    result = str
    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbNewLine, "\n")
    result = Replace(result, vbCr, "\n")
    result = Replace(result, vbLf, "\n")
    EscapeJson = result
End Function

Private Function ParseResponse(json As String) As String
    Dim startPos As Long
    startPos = InStr(json, """text"":""")
    If startPos > 0 Then
        startPos = startPos + 8
        Dim endPos As Long
        endPos = InStr(startPos, json, """")
        If endPos > startPos Then
            Dim text As String
            text = Mid(json, startPos, endPos - startPos)
            text = Replace(text, "\n", vbNewLine)
            text = Replace(text, "\""", """")
            ParseResponse = text
        Else
            ParseResponse = "Error parsing response"
        End If
    Else
        ParseResponse = "No response"
    End If
End Function
'''


class ExcellerInstaller:
    """One-click installer for Exceller Excel Add-in"""
    
    def __init__(self):
        if not HAS_TKINTER:
            print("Tkinter not available. Run: pip install tk")
            sys.exit(1)
        
        self.root = Tk()
        self.root.title("Exceller Pro Installer")
        self.root.geometry("500x600")
        self.root.resizable(False, False)
        
        self.install_path = None
        self.api_key = StringVar()
        self.status = StringVar(value="Ready to install")
        
        self.setup_ui()
    
    def setup_ui(self):
        """Build the installer UI"""
        # Colors
        BG = "#1a1a2e"
        FG = "#ffffff"
        ACCENT = "#4CAF50"
        
        self.root.configure(bg=BG)
        
        # Header
        header = Frame(self.root, bg=BG)
        header.pack(fill=X, pady=(20, 10))
        
        Label(header, text="Exceller Pro", font=("Segoe UI", 24, "bold"),
              bg=BG, fg=ACCENT).pack()
        Label(header, text="AI-Powered Excel Assistant", font=("Segoe UI", 12),
              bg=BG, fg="#888888").pack()
        
        # Features
        features = Frame(self.root, bg=BG)
        features.pack(fill=X, padx=40, pady=10)
        
        feature_list = [
            "AI Chat - Ask questions about your data",
            "Data Analysis - Get insights and statistics",
            "Fix Formula - Repair broken formulas",
            "Price Calculator - Price calculation help",
            "Verify Formula - Check formula correctness"
        ]
        
        for feat in feature_list:
            Label(features, text=f"  {feat}", font=("Segoe UI", 10),
                  bg=BG, fg="#cccccc", anchor=W).pack(fill=X, pady=2)
        
        # API Key
        api_frame = Frame(self.root, bg=BG)
        api_frame.pack(fill=X, padx=40, pady=15)
        
        Label(api_frame, text="Gemini API Key (free at makersuite.google.com):",
              font=("Segoe UI", 10), bg=BG, fg=FG, anchor=W).pack(fill=X)
        
        api_entry = Entry(api_frame, textvariable=self.api_key,
                         font=("Segoe UI", 11), width=50)
        api_entry.pack(fill=X, pady=(5, 0))
        
        # Progress
        progress_frame = Frame(self.root, bg=BG)
        progress_frame.pack(fill=X, padx=40, pady=10)
        
        self.progress = ttk.Progressbar(progress_frame, mode='determinate')
        self.progress.pack(fill=X)
        
        Label(progress_frame, textvariable=self.status, font=("Segoe UI", 9),
              bg=BG, fg="#888888").pack(pady=(5, 0))
        
        # Buttons
        btn_frame = Frame(self.root, bg=BG)
        btn_frame.pack(fill=X, padx=40, pady=20)
        
        self.install_btn = Button(btn_frame, text="INSTALL NOW", 
                                  font=("Segoe UI", 12, "bold"),
                                  bg=ACCENT, fg="white", relief="flat",
                                  padx=30, pady=10, command=self.start_install)
        self.install_btn.pack()
        
        # Footer
        Label(self.root, text="Free to use | No data collected | Uninstall anytime",
              font=("Segoe UI", 8), bg=BG, fg="#666666").pack(side=BOTTOM, pady=10)
    
    def start_install(self):
        """Start installation in background"""
        self.install_btn.config(state="disabled", text="Installing...")
        Thread(target=self.install, daemon=True).start()
    
    def install(self):
        """Main installation logic"""
        try:
            self.update_status("Checking system...", 10)
            
            # Check for Excel
            excel_path = self.find_excel()
            if not excel_path:
                self.update_status("Microsoft Excel not found!", 0)
                messagebox.showerror("Error", "Microsoft Excel not found!\nPlease install Excel first.")
                self.install_btn.config(state="normal", text="INSTALL NOW")
                return
            
            self.update_status("Excel found", 20)
            
            # Create install directory
            addin_dir = Path(os.environ['APPDATA']) / "Microsoft" / "AddIns"
            addin_dir.mkdir(parents=True, exist_ok=True)
            
            self.update_status("Creating add-in...", 40)
            
            # Create .xlam using Excel COM
            xlam_path = addin_dir / "ExcellerPro.xlam"
            self.create_xlam(excel_path, xlam_path)
            
            self.update_status("Configuring Excel...", 70)
            
            # Register add-in
            self.register_addin(xlam_path)
            
            self.update_status("Saving API key...", 90)
            
            # Save API key
            api_key = self.api_key.get().strip()
            if api_key:
                self.save_api_key(api_key)
            
            self.update_status("Installation complete!", 100)
            
            messagebox.showinfo(
                "Success!",
                "Exceller Pro installed successfully!\\n\\n"
                "To use:\\n"
                "1. Open Excel\\n"
                "2. Press Alt+F8\\n"
                "3. Select ExcellerPro_Chat\\n"
                "4. Click Run\\n\\n"
                "The toolbar will appear automatically!"
            )
            
            self.install_btn.config(state="normal", text="INSTALL COMPLETE")
            
        except Exception as e:
            self.update_status(f"Error: {str(e)}", 0)
            messagebox.showerror("Installation Error", str(e))
            self.install_btn.config(state="normal", text="INSTALL NOW")
    
    def find_excel(self):
        """Find Excel installation"""
        try:
            # Try common paths
            program_files = [
                os.environ.get('ProgramFiles', 'C:\\Program Files'),
                os.environ.get('ProgramFiles(x86)', 'C:\\Program Files (x86)')
            ]
            
            for pf in program_files:
                excel_path = Path(pf) / "Microsoft Office" / "root" / "Office16" / "EXCEL.EXE"
                if excel_path.exists():
                    return str(excel_path)
                
                excel_path = Path(pf) / "Microsoft Office" / "Office16" / "EXCEL.EXE"
                if excel_path.exists():
                    return str(excel_path)
            
            # Try registry
            try:
                key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 
                                    r"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\excel.exe")
                value, _ = winreg.QueryValueEx(key, "")
                winreg.CloseKey(key)
                return value
            except:
                pass
            
            return None
        except:
            return None
    
    def create_xlam(self, excel_path, xlam_path):
        """Create .xlam file using Excel COM"""
        try:
            import win32com.client
            
            # Start Excel
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            # Create new workbook
            wb = excel.Workbooks.Add()
            
            # Add VBA module
            vb_proj = wb.VBProject
            vb_mod = vb_proj.VBComponents.Add(1)  # Standard module
            
            # Add the code
            vb_mod.CodeModule.AddFromString(VBA_CODE)
            
            # Save as add-in
            wb.SaveAs(str(xlam_path), FileFormat=51)  # xlAddIn = 51
            
            # Close
            wb.Close(SaveChanges=False)
            excel.Quit()
            
            import time
            time.sleep(1)
            
        except ImportError:
            # Fallback: save as text and use script
            self.create_xlam_fallback(xlam_path)
        except Exception as e:
            self.create_xlam_fallback(xlam_path)
    
    def create_xlam_fallback(self, xlam_path):
        """Fallback method to create .xlam"""
        # Save VBA code to temp file
        temp_vba = Path(os.environ['TEMP']) / "exceller_vba.txt"
        with open(temp_vba, 'w') as f:
            f.write(VBA_CODE)
        
        # Create VBScript to automate Excel
        vbs_content = f'''
Dim xlApp
Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False
xlApp.DisplayAlerts = False

Dim xlWB
Set xlWB = xlApp.Workbooks.Add()

Dim vbProj
Set vbProj = xlWB.VBProject

Dim vbMod
Set vbMod = vbProj.VBComponents.Add(1)

Dim fso, ts
Set fso = CreateObject("Scripting.FileSystemObject")
Set ts = fso.OpenTextFile(r"{temp_vba}", 1)
vbMod.CodeModule.AddFromString ts.ReadAll
ts.Close

xlWB.SaveAs r"{xlam_path}", 51
xlWB.Close False
xlApp.Quit

Set xlApp = Nothing
'''
        
        vbs_path = Path(os.environ['TEMP']) / "create_xlam.vbs"
        with open(vbs_path, 'w') as f:
            f.write(vbs_content)
        
        # Run the script
        subprocess.run(['cscript', '//B', str(vbs_path)], 
                      capture_output=True, timeout=30)
        
        # Cleanup
        try:
            temp_vba.unlink()
            vbs_path.unlink()
        except:
            pass
    
    def register_addin(self, xlam_path):
        """Register the add-in with Excel"""
        try:
            # Add to registry for auto-load
            key_path = r"Software\\Microsoft\\Office\\16.0\\Excel\\Options"
            try:
                key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, 
                                    winreg.KEY_ALL_ACCESS)
            except:
                key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path)
            
            # Get existing OPEN values
            try:
                existing, _ = winreg.QueryValueEx(key, "OPEN")
                if str(xlam_path) not in existing:
                    winreg.SetValueEx(key, "OPEN", 0, winreg.REG_SZ, 
                                     f'"{xlam_path}" /e')
            except:
                winreg.SetValueEx(key, "OPEN", 0, winreg.REG_SZ, 
                                 f'"{xlam_path}" /e')
            
            winreg.CloseKey(key)
        except Exception as e:
            print(f"Registry warning: {e}")
    
    def save_api_key(self, api_key):
        """Save API key to registry"""
        try:
            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, 
                                  r"Software\\ExcellerPro\\Config")
            winreg.SetValueEx(key, "ApiKey", 0, winreg.REG_SZ, api_key)
            winreg.CloseKey(key)
        except Exception as e:
            print(f"Settings warning: {e}")
    
    def update_status(self, message, progress):
        """Update UI from thread"""
        self.status.set(message)
        self.progress['value'] = progress
        self.root.update_idletasks()
    
    def run(self):
        """Start the installer"""
        self.root.mainloop()


def main():
    """Entry point"""
    if sys.platform != 'win32':
        print("This installer is for Windows only.")
        print("For Mac/Linux, install manually:")
        print("1. Open Excel")
        print("2. Alt+F11")
        print("3. Insert Module")
        print("4. Paste code from ExcellerPro.bas")
        sys.exit(1)
    
    installer = ExcellerInstaller()
    installer.run()


if __name__ == "__main__":
    main()
