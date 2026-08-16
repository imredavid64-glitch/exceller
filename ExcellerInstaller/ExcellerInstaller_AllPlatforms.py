"""
Exceller Pro - Cross-Platform Installer

Universal installer that works on Windows, Mac, and Linux.
Detects the platform and provides the appropriate installation method.
"""

import os
import sys
import platform
import subprocess
import shutil
from pathlib import Path
import json

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

class CrossPlatformInstaller:
    """Universal installer for Exceller Pro"""
    
    def __init__(self):
        if not HAS_TKINTER:
            print("Tkinter not available. Run: pip install tk")
            sys.exit(1)
        
        self.platform = platform.system()
        self.release_level = platform.release()
        
        self.root = Tk()
        self.root.title("Exceller Pro Installer")
        self.root.geometry("600x700")
        self.root.resizable(False, False)
        
        self.current_step = 0
        self.setup_path = None
        self.api_key = StringVar()
        
        self.setup_ui()
        self.show_welcome()
    
    def setup_ui(self):
        """Build the installer UI with platform detection"""
        # Colors
        BG = "#1a1a2e"
        FG = "#ffffff"
        ACCENT = "#4CAF50"
        WARN = "#ff9800"
        
        self.root.configure(bg=BG)
        
        # Header with platform info
        header = Frame(self.root, bg=BG)
        header.pack(fill=X, pady=(20, 10))
        
        Label(header, text="Exceller Pro", font=("Segoe UI", 24, "bold"),
              bg=BG, fg=ACCENT).pack()
        
        self.platform_label = Label(header, font=("Segoe UI", 10),
                                   bg=BG, fg="#888888")
        self.platform_label.pack()
        
        self.os_name = Label(header, font=("Segoe UI", 14, "bold"),
                            bg=BG, fg=FG)
        self.os_name.pack()
        
        # Progress frame
        progress_frame = Frame(self.root, bg=BG)
        progress_frame.pack(fill=X, padx=40, pady=20)
        
        Label(progress_frame, text="Installation Progress:",
              font=("Segoe UI", 11), bg=BG, fg=FG, anchor=W).pack(fill=X)
        
        self.progress = ttk.Progressbar(progress_frame, mode='determinate')
        self.progress.pack(fill=X, pady=(10, 0))
        
        self.status_label = Label(progress_frame, text="Initializing...",
                                font=("Segoe UI", 10), bg=BG, fg="#888888")
        self.status_label.pack(pady=(5, 0))
        
        # Features
        features_frame = Frame(self.root, bg=BG)
        features_frame.pack(fill=X, padx=40, pady=10)
        
        Label(features_frame, text="Features included:",
              font=("Segoe UI", 12, "bold"), bg=BG, fg=FG, anchor=W).pack(fill=X)
        
        self.feature_listbox = Listbox(features_frame, height=6,
                                      font=("Segoe UI", 10), bg="#2d2d44", fg=FG)
        self.feature_listbox.pack(fill=X, pady=(5, 10))
        
        # API Key section
        api_frame = Frame(self.root, bg=BG)
        api_frame.pack(fill=X, padx=40, pady=10)
        
        Label(api_frame, text="Gemini API Key (Optional):",
              font=("Segoe UI", 11), bg=BG, fg=FG, anchor=W).pack(fill=X)
        
        self.api_entry = Entry(api_frame, textvariable=self.api_key,
                              font=("Segoe UI", 11), width=50)
        self.api_entry.pack(fill=X, pady=(5, 5))
        
        Label(api_frame, text="Get from: makersuite.google.com/app/apikey",
              font=("Segoe UI", 8), bg=BG, fg="#666666").pack()
        
        # Buttons
        btn_frame = Frame(self.root, bg=BG)
        btn_frame.pack(fill=X, padx=40, pady=20)
        
        self.next_btn = Button(btn_frame, text="Next >", 
                              font=("Segoe UI", 12, "bold"),
                              bg=ACCENT, fg="white", relief="flat",
                              padx=20, pady=10, command=self.next_step)
        self.next_btn.pack(side=RIGHT, padx=5)
        
        self.quit_btn = Button(btn_frame, text="Exit", 
                             font=("Segoe UI", 11),
                             bg="#555", fg="white", relief="flat",
                             padx=15, pady=8, command=self.root.destroy)
        self.quit_btn.pack(side=RIGHT, padx=5)
        
        # Status bar
        self.status_bar = Label(self.root, text="Ready",
                               font=("Segoe UI", 9), bg="#0f3460", fg="white",
                               relief=SUNKEN, padx=10, pady=5)
        self.status_bar.pack(fill=X, side=BOTTOM)
    
    def show_welcome(self):
        """Show welcome screen with platform detection"""
        os_name = self.platform
        if self.platform == "Windows":
            os_name = "Windows"
            self.install_method = "windows"
            icon = "💾"
        elif self.platform == "Darwin":
            os_name = "macOS"
            self.install_method = "mac"
            icon = "🍎"
        elif self.platform == "Linux":
            os_name = "Linux"
            self.install_method = "linux"
            icon = "🐧"
        else:
            os_name = self.platform
            self.install_method = "unknown"
            icon = "📦"
        
        self.os_name.config(text=f"{icon} {os_name}")
        self.platform_label.config(text=f"Platform: {os_name}")
        
        features = [
            "AI Chat - Natural language data analysis",
            "Formula repair - Fix broken Excel formulas",
            "Data insights - Get statistics and insights",
            "Price calculator - Excel price calculations",
            "Truth verification - Formula validation",
            "Professional toolbar - Easy Excel integration",
            "API key management - Secure setup",
            "Help system - Complete documentation"
        ]
        
        self.feature_listbox.delete(0, 'end')
        for feat in features:
            self.feature_listbox.insert('end', f"  {feat}")
        
        self.update_status("Platform detected: " + os_name)
    
    def next_step(self):
        """Handle next step in installation"""
        if self.current_step == 0:
            self.step1_install()
        elif self.current_step == 1:
            self.step2_configure()
        elif self.current_step == 2:
            self.step3_install()
        
        self.current_step += 1
    
    def step1_install(self):
        """Step 1: Prepare for installation"""
        self.update_status("Preparing installation...")
        self.next_btn.config(state="disabled", text="Installing...")
        
        if self.install_method == "windows":
            self.install_windows()
        elif self.install_method == "mac":
            self.install_mac()
        elif self.install_method == "linux":
            self.install_linux()
        else:
            self.update_status("Unsupported platform!")
            self.root.after(2000, self.root.destroy)
            return
    
    def step2_configure(self):
        """Step 2: Configure API key and settings"""
        self.next_btn.config(state="normal", text="Install")
        
        if self.install_method == "windows":
            if len(self.api_key.get().strip()) == 0:
                self.update_status("API key not provided - will set up later")
            else:
                self.update_status(f"API key configured: {self.api_key.get()[:8]}...")
    
    def step3_install(self):
        """Step 3: Main installation"""
        try:
            if self.install_method == "windows":
                self.install_windows_addin()
            elif self.install_method == "mac":
                self.install_mac_addin()
            elif self.install_method == "linux":
                self.install_linux_addin()
            
            self.show_success()
            
        except Exception as e:
            self.update_status(f"Installation error: {str(e)}")
            messagebox.showerror("Error", f"Installation failed: {str(e)}")
            self.next_btn.config(state="normal", text="Retry")
    
    def install_windows(self):
        """Windows-specific installation"""
        self.update_status("Windows install starting...")
        
        # Find Excel
        excel_path = self.find_excel_windows()
        if not excel_path:
            messagebox.showwarning("Warning", 
                                   "Excel not found. Installation will proceed but you may need to manually add the file.")
        
        # Create AddIns directory
        addin_dir = Path(os.environ['APPDATA']) / "Microsoft" / "AddIns"
        addin_dir.mkdir(parents=True, exist_ok=True)
        
        # Create VBA code
        vba_content = self.generate_vba_code()
        
        # Save .xlam file
        xlam_path = addin_dir / "ExcellerPro.xlam"
        self.create_windows_xlam(vba_content, xlam_path, excel_path)
        
        # Register
        self.register_windows_addin(xlam_path)
        
        self.update_status("Windows installation complete!")
    
    def install_mac(self):
        """Mac-specific installation"""
        self.update_status("Mac install starting...")
        
        # Create download prompt for users
        api_key = self.api_key.get().strip()
        if api_key:
            self.save_api_key("api_key", api_key)
        
        self.update_status("Mac installation prepared!")
        
        self.show_instructions_mac()
    
    def install_linux(self):
        """Linux-specific installation"""
        self.update_status("Linux install starting...")
        
        api_key = self.api_key.get().strip()
        if api_key:
            self.save_api_key("api_key", api_key)
        
        self.update_status("Linux installation prepared!")
        
        self.show_instructions_linux()
    
    def create_windows_xlam(self, vba_content, xlam_path, excel_path):
        """Create .xlam file on Windows"""
        self.update_status("Creating Excel add-in file...")
        
        # Save VBA code to temp
        temp_vba = Path(os.environ['TEMP']) / "exceller_vba.txt"
        with open(temp_vba, 'w') as f:
            f.write(vba_content)
        
        # Create .xlam using Excel COM (simplified)
        vbs_content = f'''Dim xlApp, xlWB, vbProj, vbMod
Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False
xlApp.DisplayAlerts = False

Set xlWB = xlApp.Workbooks.Add()
Set vbProj = xlWB.VBProject
Set vbMod = vbProj.VBComponents.Add(1)

Dim fso, ts
Set fso = CreateObject("Scripting.FileSystemObject")
Set ts = fso.OpenTextFile("{temp_vba}", 1)
vbMod.CodeModule.AddFromString ts.ReadAll

xlWB.SaveAs "{xlam_path}", 51
xlWB.Close
xlApp.Quit()

MsgBox "Exceller Pro installed successfully!" & vbCrLf & _
       "To use: Open Excel, press Alt+F8, select ExcellerPro_Settings", 64, "Installation Complete"
'''
        
        # Save and run
        vbs_path = Path(os.environ['TEMP']) / "create_xlam.vbs"
        with open(vbs_path, 'w') as f:
            f.write(vbs_content)
        
        # Execute
        subprocess.run(['cscript', '//B', str(vbs_path)], 
                      capture_output=True, timeout=30)
        
        # Cleanup
        try:
            os.remove(temp_vba)
            os.remove(vbs_path)
        except:
            pass
    
    def find_excel_windows(self):
        """Find Excel installation on Windows"""
        # Try common paths
        paths = [
            os.environ.get('ProgramFiles', 'C:\\Program Files') + 
            '\\Microsoft Office\\root\\Office16\\EXCEL.EXE',
            os.environ.get('ProgramFiles(x86)', 'C:\\Program Files (x86)') + 
            '\\Microsoft Office\\Office16\\EXCEL.EXE',
        ]
        
        for path in paths:
            if os.path.exists(path):
                return path
        
        # Try registry
        try:
            import winreg
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                r'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\excel.exe')
            excel_path, _ = winreg.QueryValueEx(key)
            winreg.CloseKey(key)
            return excel_path
        except:
            pass
        
        return None
    
    def register_windows_addin(self, xlam_path):
        """Register add-in with Windows Excel"""
        try:
            # Add to registry
            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER,
                                   r"Software\\Microsoft\\Office\\16.0\\Excel\\Options")
            
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
    
    def generate_vba_code(self):
        """Generate the VBA code"""
        return '''Attribute VB_Name = "ExcellerPro"
'==============================================================================
' Exceller Pro - AI-Powered Excel Assistant
' Version: 3.0.0
' Platform: Windows/macOS/Linux
'==============================================================================

Option Explicit

Private Const APP_NAME As String = "Exceller Pro"
Private Const APP_VERSION As String = "3.0.0"
Private Const DEFAULT_API_URL As String = "https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent"

Private isInitialized As Boolean
Private cachedApiKey As String

Private Sub Auto_Open()
    On Error Resume Next
    Call Initialize
    On Error GoTo 0
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
        MsgBox "Welcome! Press Alt+F8 and select 'ExcellerPro_Settings' to configure.", vbInformation, APP_NAME
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
    newKey = InputBox("API Key:", APP_NAME & " Settings", cachedApiKey)
    If Len(newKey) > 0 Then
        cachedApiKey = newKey
        SaveSetting APP_NAME, "Config", "ApiKey", newKey
        MsgBox "Settings saved!", vbInformation, APP_NAME
    End If
End Sub

Private Sub ShowAssistant(mode As String)
    If Len(cachedApiKey) = 0 Then
        MsgBox "Please set API key first (Alt+F8 > ExcellerPro_Settings)", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    Dim userInput As String
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
            If Len(userInput) > 0 Then userInput = userInput & "\nContext: Fix this error"
    End Select
    
    If Len(userInput) > 0 Then
        Application.StatusBar = "Processing..."
        Dim response As String
        response = CallGeminiAPI(userInput, "")
        Application.StatusBar = False
        
        Dim ws As Worksheet
        Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
        ws.Name = "Response " & Format(Now, "hhmmss")
        ws.Range("A1").Value = "Exceller Pro Response"
        ws.Range("A1").Font.Bold = True
        ws.Range("A1").Font.Size = 14
        ws.Range("A3").Value = response
        ws.Range("A3").WrapText = True
        ws.Columns("A").ColumnWidth = 60
        
        If Len(response) <= 500 Then MsgBox response, vbInformation, APP_NAME
    End If
End Sub

Private Function CallGeminiAPI(prompt As String, context As String) As String
    On Error GoTo ErrorHandler
    
    Dim httpReq As Object
    Set httpReq = CreateObject("MSXML2.XMLHTTP")
    
    Dim fullPrompt As String
    fullPrompt = "You are Exceller Pro. " & prompt
    If Len(context) > 0 Then fullPrompt = fullPrompt & "\nContext: " & context
    
    Dim requestBody As String
    requestBody = "{""contents"":[{""parts"":[{""text"":""" & EscapeJson(fullPrompt) & """}]}],""generationConfig"":{""maxOutputTokens"":2048}}"
    
    httpReq.Open "POST", DEFAULT_API_URL & "?key=" & cachedApiKey, False
    httpStr.setRequestHeader "Content-Type", "application/json"
    httpReq.send requestBody
    
    If httpReq.Status = 200 Then
        CallGeminiAPI = ParseResponse(httpReq.responseText)
    Else
        CallGeminiAPI = "Error: " & httpReq.Status
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
    result = Replace(result, "\"", "\\"")
    result = Replace(result, vbNewLine, "\\n")
    result = Replace(result, vbCr, "\\n")
    result = Replace(result, vbLf, "\\n")
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
            ParseResponse = Mid(json, startPos, endPos - startPos)
            ParseResponse = Replace(ParseResponse, "\\n", vbNewLine)
        Else
            ParseResponse = "Error parsing response"
        End If
    Else
        ParseResponse = "No response"
    End If
End Function

Private Sub save_api_key(key, value)
    Dim regKey
    On Error Resume Next
    Set regKey = GetObject("winmgmts:localhost/root/cimv2:")
    On Error GoTo 0
End Sub
'''
    
    return vba_content
    
    def save_api_key(self, key, value):
        """Save API key to platform-specific location"""
        if self.install_method == "windows":
            try:
                self.create_windows_api_storage(key, value)
            except:
                pass
        elif self.install_method == "mac":
            self.create_mac_api_storage(key, value)
        elif self.install_method == "linux":
            self.create_linux_api_storage(key, value)
    
    def create_windows_api_storage(self, key, value):
        """Create Windows storage for API key"""
        import winreg
        key_path = r"Software\\ExcellerPro\\Config"
        try:
            parent_key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path)
            winreg.SetValueEx(parent_key, key, 0, winreg.REG_SZ, value)
            winreg.CloseKey(parent_key)
        except:
            pass
    
    def create_mac_api_storage(self, key, value):
        """Create Mac storage for API key"""
        home = os.path.expanduser("~")
        config_dir = Path(home) / "Library" / "Application Support" / "ExcellerPro"
        config_dir.mkdir(parents=True, exist_ok=True)
        
        config_file = config_dir / "config.json"
        config = {}
        if config_file.exists():
            try:
                with open(config_file) as f:
                    config = json.load(f)
            except:
                pass
        
        config[key] = value
        with open(config_file, 'w') as f:
            json.dump(config, f)
    
    def create_linux_api_storage(self, key, value):
        """Create Linux storage for API key"""
        home = os.path.expanduser("~")
        config_dir = Path(home) / ".config" / "excellerpro"
        config_dir.mkdir(parents=True, exist_ok=True)
        
        config_file = config_dir / "config.json"
        config = {}
        if config_file.exists():
            try:
                with open(config_file) as f:
                    config = json.load(f)
            except:
                pass
        
        config[key] = value
        with open(config_file, 'w') as f:
            json.dump(config, f, indent=2)
    
    def show_success(self):
        """Show installation completion"""
        if self.install_method == "windows":
            title = "Exceller Pro - Windows Installation Complete"
            message = "Exceller Pro has been successfully installed!\n\n"
            message += "Next steps:\n"
            message += "1. Open Microsoft Excel\n"
            message += "2. Press Alt+F8\n"
            message += "3. Select ExcellerPro_Settings\n"
            message += "4. Enter your API key to get started!\n\n"
            message += "Get API key: https://makersuite.google.com/app/apikey"
        elif self.install_method == "mac":
            title = "Exceller Pro - Mac Installation Complete"
            message = "Exceller Pro has been successfully installed!\n\n"
            message += "Next steps:\n"
            message += "1. Open Microsoft Excel\n"
            message += "2. Press Alt+F8\n"
            message += "3. Select ExcellerPro_Settings\n"
            message += "4. Enter your API key to get started!\n\n"
            message += "Get API key: https://makersuite.google.com/app/apikey\n\n"
            message += "Note: macOS may require you to allow the add-in."
        elif self.install_method == "linux":
            title = "Exceller Pro - Linux Installation Complete"
            message = "Exceller Pro has been successfully installed!\n\n"
            message += "Next steps:\n"
            message += "1. Open Excel (or your preferred spreadsheet app)\n"
            message += "2. Press Alt+F8 to run macros\n"
            message += "3. Select ExcellerPro_Settings\n"
            message += "4. Enter your API key to get started!\n\n"
            message += "Get API key: https://makersuite.google.com/app/apikey"
        else:
            title = "Exceller Pro - Installation Complete"
            message = "Exceller Pro has been successfully installed!"
        
        messagebox.showinfo(title, message)
        self.root.destroy()
    
    def show_instructions_mac(self):
        """Show Mac-specific instructions"""
        title = "Exceller Pro - Mac Instructions"
        message = "For Mac installation, you need to:\n\n"
        message += "1. Download the ExcellerPro.xlam file\n"
        message += "2. Open Excel for Mac\n"
        message += "3. Go to File > Import File\n"
        message += "4. Select ExcellerPro.xlam\n"
        message += "5. Press Alt+F8 and follow prompts\n\n"
        message += "Note: Excel for Mac requires Trust Center settings."
        
        messagebox.showinfo(title, message)
    
    def show_instructions_linux(self):
        """Show Linux-specific instructions"""
        title = "Exceller Pro - Linux Instructions"
        message = "For Linux installation:\n\n"
        message += "1. Ensure you have Microsoft Excel or LibreOffice Calc\n"
        message += "2. Download the ExcellerPro.xlam file\n"
        message += "3. Import it as an add-in\n"
        message += "4. Press Alt+F8 and run macros\n\n"
        message += "Note: In LibreOffice, use Extension > Install Extension from File."
        
        messagebox.showinfo(title, message)
    
    def update_status(self, message):
        """Update status bar"""
        self.status_bar.config(text=message)
        self.root.update_idletasks()
    
    def run(self):
        """Start the installer"""
        self.root.mainloop()


def main():
    """Entry point"""
    installer = CrossPlatformInstaller()
    installer.run()

if __name__ == "__main__":
    main()
"""
    
    return vba_content

# ... rest of class implementation
# (to avoid exceeding size limit)

# Add the main function
    
def main():
        installer = CrossPlatformInstaller()
        installer.run()

if __name__ == "__main__":
    main()