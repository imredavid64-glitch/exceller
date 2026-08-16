"""
Exceller - AI-Powered Excel Assistant Desktop App
A Grammarly-like assistant for Microsoft Excel
"""

import sys
import os
import json
import threading
from pathlib import Path
from datetime import datetime

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTextEdit, QPushButton, QLabel, QSystemTrayIcon, QMenu,
    QMessageBox, QFileDialog, QFrame, QGraphicsDropShadowEffect,
    QLineEdit, QComboBox, QProgressBar, QSplitter, QDialog,
    QDialogButtonBox, QTabWidget, QTableWidget, QTableWidgetItem,
    QHeaderView, QGroupBox, QRadioButton, QButtonGroup, QTextBrowser
)
from PyQt6.QtCore import Qt, QTimer, pyqtSignal, QObject, QSize, QPropertyAnimation, QEasingCurve
from PyQt6.QtGui import (
    QFont, QColor, QIcon, QPixmap, QPainter, QLinearGradient,
    QBrush, QPen, QCursor, QAction
)

from modules.ai_engine import AIEngine
from modules.excel_parser import ExcelParser
from modules.file_monitor import FileMonitor
from modules.theme import ThemeManager


class AssistantSignals(QObject):
    """Signals for thread communication"""
    response_ready = pyqtSignal(str)
    status_update = pyqtSignal(str)
    file_detected = pyqtSignal(str)
    error_occurred = pyqtSignal(str)
    json_response_ready = pyqtSignal(dict)
    verification_complete = pyqtSignal(dict)


class FormulaFixerDialog(QDialog):
    """Dialog for fixing formula errors"""
    
    def __init__(self, ai_engine, parent=None):
        super().__init__(parent)
        self.ai_engine = ai_engine
        self.result_data = None
        self.init_ui()
        
    def init_ui(self):
        self.setWindowTitle("Formula Error Fixer")
        self.setMinimumSize(600, 500)
        self.setStyleSheet("""
            QDialog { background-color: #1a1a2e; }
            QLabel { color: white; }
            QLineEdit, QTextEdit { 
                background-color: #16213e; color: white; 
                border: 1px solid #2d2d44; border-radius: 6px;
                padding: 8px;
            }
            QPushButton {
                background-color: #4CAF50; color: white;
                border: none; border-radius: 6px;
                padding: 8px 16px; font-weight: bold;
            }
            QPushButton:hover { background-color: #45a049; }
            QGroupBox {
                color: white; border: 1px solid #2d2d44;
                border-radius: 8px; margin-top: 12px;
                padding-top: 20px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px; padding: 0 8px;
            }
        """)
        
        layout = QVBoxLayout(self)
        
        # Title
        title = QLabel("🔧 Formula Error Fixer")
        title.setFont(QFont("Segoe UI", 16, QFont.Weight.Bold))
        layout.addWidget(title)
        
        # Input section
        input_group = QGroupBox("Enter Formula Details")
        input_layout = QVBoxLayout(input_group)
        
        # Formula input
        formula_label = QLabel("Formula with error:")
        input_layout.addWidget(formula_label)
        
        self.formula_input = QLineEdit()
        self.formula_input.setPlaceholderText("e.g., =VLOOKUP(A1, Sheet2!B:C, 2, FALSE)")
        input_layout.addWidget(self.formula_input)
        
        # Error message input
        error_label = QLabel("Error message (if any):")
        input_layout.addWidget(error_label)
        
        self.error_input = QLineEdit()
        self.error_input.setPlaceholderText("e.g., #REF!, #VALUE!, #N/A")
        input_layout.addWidget(self.error_input)
        
        # Context input
        context_label = QLabel("Additional context (optional):")
        input_layout.addWidget(context_label)
        
        self.context_input = QTextEdit()
        self.context_input.setPlaceholderText("Describe what the formula should do, column names, etc.")
        self.context_input.setMaximumHeight(80)
        input_layout.addWidget(self.context_input)
        
        layout.addWidget(input_group)
        
        # Fix button
        self.fix_btn = QPushButton("🔍 Find Fixes")
        self.fix_btn.clicked.connect(self.fix_formula)
        layout.addWidget(self.fix_btn)
        
        # Results section
        results_group = QGroupBox("Suggested Fixes")
        results_layout = QVBoxLayout(results_group)
        
        self.results_browser = QTextBrowser()
        self.results_browser.setOpenExternalLinks(True)
        results_layout.addWidget(self.results_browser)
        
        layout.addWidget(results_group)
        
        # Apply button
        self.apply_btn = QPushButton("📋 Copy Selected Fix")
        self.apply_btn.clicked.connect(self.copy_fix)
        layout.addWidget(self.apply_btn)
        
    def fix_formula(self):
        """Fix the formula using AI"""
        formula = self.formula_input.text().strip()
        error = self.error_input.text().strip()
        context = self.context_input.toPlainText().strip()
        
        if not formula:
            QMessageBox.warning(self, "Input Required", "Please enter a formula.")
            return
        
        self.fix_btn.setEnabled(False)
        self.fix_btn.setText("Analyzing...")
        self.results_browser.setPlainText("Analyzing formula error...")
        
        # Run in thread
        def run_fix():
            result = self.ai_engine.fix_formula_error(formula, error, context)
            self.result_data = result
            self.display_results(result)
            self.fix_btn.setEnabled(True)
            self.fix_btn.setText("🔍 Find Fixes")
        
        thread = threading.Thread(target=run_fix, daemon=True)
        thread.start()
    
    def display_results(self, result):
        """Display the fix results"""
        if "raw_response" in result:
            self.results_browser.setPlainText(result["raw_response"])
            return
        
        html = f"""
        <h3 style="color: #4CAF50;">Error Type: {result.get('error_type', 'Unknown')}</h3>
        <p><strong>Diagnosis:</strong> {result.get('diagnosis', 'N/A')}</p>
        <hr style="border-color: #2d2d44;">
        <h4 style="color: #2196F3;">Suggested Fixes:</h4>
        """
        
        fixes = result.get('fixes', [])
        for i, fix in enumerate(fixes, 1):
            html += f"""
            <div style="background: #16213e; padding: 12px; margin: 8px 0; border-radius: 8px; border-left: 4px solid #4CAF50;">
                <strong>Option {i}:</strong> {fix.get('best_for', '')}<br>
                <code style="background: #0f3460; padding: 4px 8px; border-radius: 4px; display: block; margin: 8px 0;">
                    {fix.get('formula', '')}
                </code>
                <small>{fix.get('explanation', '')}</small>
            </div>
            """
        
        tips = result.get('prevention_tips', [])
        if tips:
            html += '<h4 style="color: #ff9800;">Prevention Tips:</h4><ul>'
            for tip in tips:
                html += f"<li>{tip}</li>"
            html += "</ul>"
        
        self.results_browser.setHtml(html)
    
    def copy_fix(self):
        """Copy the first fix to clipboard"""
        if self.result_data and 'fixes' in self.result_data:
            fixes = self.result_data['fixes']
            if fixes:
                formula = fixes[0].get('formula', '')
                QApplication.clipboard().setText(formula)
                QMessageBox.information(self, "Copied", "Formula copied to clipboard!")


class PriceCalculatorDialog(QDialog):
    """Dialog for alternative price calculations"""
    
    def __init__(self, ai_engine, parent=None):
        super().__init__(parent)
        self.ai_engine = ai_engine
        self.result_data = None
        self.init_ui()
        
    def init_ui(self):
        self.setWindowTitle("Price Calculation Helper")
        self.setMinimumSize(650, 550)
        self.setStyleSheet("""
            QDialog { background-color: #1a1a2e; }
            QLabel { color: white; }
            QLineEdit, QTextEdit { 
                background-color: #16213e; color: white; 
                border: 1px solid #2d2d44; border-radius: 6px;
                padding: 8px;
            }
            QPushButton {
                background-color: #2196F3; color: white;
                border: none; border-radius: 6px;
                padding: 8px 16px; font-weight: bold;
            }
            QPushButton:hover { background-color: #1976D2; }
            QGroupBox {
                color: white; border: 1px solid #2d2d44;
                border-radius: 8px; margin-top: 12px;
                padding-top: 20px;
            }
        """)
        
        layout = QVBoxLayout(self)
        
        # Title
        title = QLabel("💰 Price Calculation Helper")
        title.setFont(QFont("Segoe UI", 16, QFont.Weight.Bold))
        layout.addWidget(title)
        
        # Input section
        input_group = QGroupBox("Describe Your Price Calculation")
        input_layout = QVBoxLayout(input_group)
        
        # Current formula
        current_label = QLabel("Current formula (if any):")
        input_layout.addWidget(current_label)
        
        self.formula_input = QLineEdit()
        self.formula_input.setPlaceholderText("e.g., =A2*B2")
        input_layout.addWidget(self.formula_input)
        
        # Columns
        columns_label = QLabel("Available columns (comma-separated):")
        input_layout.addWidget(columns_label)
        
        self.columns_input = QLineEdit()
        self.columns_input.setPlaceholderText("e.g., Quantity, UnitPrice, Discount, Tax, Total")
        input_layout.addWidget(self.columns_input)
        
        # Description
        desc_label = QLabel("What should the calculation do?")
        input_layout.addWidget(desc_label)
        
        self.desc_input = QTextEdit()
        self.desc_input.setPlaceholderText("e.g., Calculate final price with quantity discount and tax")
        self.desc_input.setMaximumHeight(80)
        input_layout.addWidget(self.desc_input)
        
        layout.addWidget(input_group)
        
        # Calculate button
        self.calc_btn = QPushButton("🔄 Find Alternative Calculations")
        self.calc_btn.clicked.connect(self.find_calculations)
        layout.addWidget(self.calc_btn)
        
        # Results section with tabs
        self.results_tabs = QTabWidget()
        self.results_tabs.setStyleSheet("""
            QTabWidget::pane { border: 1px solid #2d2d44; border-radius: 8px; }
            QTabBar::tab { 
                background: #16213e; color: white; 
                padding: 8px 16px; border-top-left-radius: 6px;
                border-top-right-radius: 6px;
            }
            QTabBar::tab:selected { background: #4CAF50; }
        """)
        
        # All approaches tab
        self.approaches_browser = QTextBrowser()
        self.results_tabs.addTab(self.approaches_browser, "All Methods")
        
        # Comparison tab
        self.comparison_browser = QTextBrowser()
        self.results_tabs.addTab(self.comparison_browser, "Comparison")
        
        layout.addWidget(self.results_tabs)
        
        # Copy button
        self.copy_btn = QPushButton("📋 Copy Formula")
        self.copy_btn.clicked.connect(self.copy_formula)
        layout.addWidget(self.copy_btn)
    
    def find_calculations(self):
        """Find alternative price calculations"""
        formula = self.formula_input.text().strip()
        columns = self.columns_input.text().strip()
        desc = self.desc_input.toPlainText().strip()
        
        if not columns:
            QMessageBox.warning(self, "Input Required", "Please enter column names.")
            return
        
        self.calc_btn.setEnabled(False)
        self.calc_btn.setText("Calculating...")
        
        columns_list = [c.strip() for c in columns.split(',')]
        context = f"Description: {desc}" if desc else ""
        
        def run_calc():
            result = self.ai_engine.suggest_price_calculation(
                formula, columns_list, context
            )
            self.result_data = result
            self.display_results(result)
            self.calc_btn.setEnabled(False)
            self.calc_btn.setText("🔄 Find Alternative Calculations")
        
        thread = threading.Thread(target=run_calc, daemon=True)
        thread.start()
    
    def display_results(self, result):
        """Display calculation results"""
        if "raw_response" in result:
            self.approaches_browser.setPlainText(result["raw_response"])
            return
        
        approaches = result.get('approaches', [])
        
        # All approaches tab
        html = "<h3 style='color: #4CAF50;'>Alternative Price Calculations</h3>"
        for i, approach in enumerate(approaches, 1):
            pros = ", ".join(approach.get('pros', []))
            cons = ", ".join(approach.get('cons', []))
            
            html += f"""
            <div style="background: #16213e; padding: 12px; margin: 10px 0; border-radius: 8px;">
                <h4 style="color: #2196F3;">{i}. {approach.get('name', 'Method ' + str(i))}</h4>
                <p><strong>Best for:</strong> {approach.get('best_for', 'N/A')}</p>
                <code style="background: #0f3460; padding: 8px; border-radius: 4px; display: block; margin: 8px 0;">
                    {approach.get('formula', '')}
                </code>
                <p>{approach.get('description', '')}</p>
                <p style="color: #4CAF50;"><small>Pros: {pros}</small></p>
                <p style="color: #f44336;"><small>Cons: {cons}</small></p>
                <p><em>Example: {approach.get('example', 'N/A')}</em></p>
            </div>
            """
        
        rec = result.get('recommendation', '')
        if rec:
            html += f"<div style='background: #0f3460; padding: 12px; border-radius: 8px; border-left: 4px solid #4CAF50;'><strong>Recommendation:</strong> {rec}</div>"
        
        self.approaches_browser.setHtml(html)
        
        # Comparison tab
        comp_html = "<h3 style='color: #FF9800;'>Quick Comparison</h3>"
        comp_html += "<table style='width:100%; border-collapse: collapse;'>"
        comp_html += "<tr style='background: #16213e;'><th style='padding: 8px; border: 1px solid #2d2d44;'>Method</th><th style='padding: 8px; border: 1px solid #2d2d44;'>Best For</th><th style='padding: 8px; border: 1px solid #2d2d44;'>Complexity</th></tr>"
        
        for i, approach in enumerate(approaches, 1):
            bg = "#0f3460" if i % 2 == 0 else "#16213e"
            comp_html += f"<tr style='background: {bg};'><td style='padding: 8px; border: 1px solid #2d2d44;'>{approach.get('name', '')}</td><td style='padding: 8px; border: 1px solid #2d2d44;'>{approach.get('best_for', '')}</td><td style='padding: 8px; border: 1px solid #2d2d44;'>{approach.get('complexity', 'Medium')}</td></tr>"
        
        comp_html += "</table>"
        self.comparison_browser.setHtml(comp_html)
    
    def copy_formula(self):
        """Copy selected formula"""
        if self.result_data and 'approaches' in self.result_data:
            approaches = self.result_data['approaches']
            if approaches:
                formula = approaches[0].get('formula', '')
                QApplication.clipboard().setText(formula)
                QMessageBox.information(self, "Copied", "Formula copied to clipboard!")


class VerificationDialog(QDialog):
    """Dialog for verifying formulas and code"""
    
    def __init__(self, ai_engine, parent=None):
        super().__init__(parent)
        self.ai_engine = ai_engine
        self.result_data = None
        self.init_ui()
        
    def init_ui(self):
        self.setWindowTitle("Formula & Code Verifier")
        self.setMinimumSize(600, 500)
        self.setStyleSheet("""
            QDialog { background-color: #1a1a2e; }
            QLabel { color: white; }
            QLineEdit, QTextEdit { 
                background-color: #16213e; color: white; 
                border: 1px solid #2d2d44; border-radius: 6px;
                padding: 8px;
            }
            QPushButton {
                background-color: #9C27B0; color: white;
                border: none; border-radius: 6px;
                padding: 8px 16px; font-weight: bold;
            }
            QPushButton:hover { background-color: #7B1FA2; }
            QGroupBox {
                color: white; border: 1px solid #2d2d44;
                border-radius: 8px; margin-top: 12px;
                padding-top: 20px;
            }
            QRadioButton { color: white; }
        """)
        
        layout = QVBoxLayout(self)
        
        # Title
        title = QLabel("✅ Formula & Code Verifier")
        title.setFont(QFont("Segoe UI", 16, QFont.Weight.Bold))
        layout.addWidget(title)
        
        # Mode selection
        mode_group = QGroupBox("Verification Mode")
        mode_layout = QHBoxLayout(mode_group)
        
        self.mode_group = QButtonGroup()
        
        self.formula_mode = QRadioButton("Verify Formula")
        self.formula_mode.setChecked(True)
        self.mode_group.addButton(self.formula_mode, 1)
        mode_layout.addWidget(self.formula_mode)
        
        self.vba_mode = QRadioButton("Verify VBA Code")
        self.mode_group.addButton(self.vba_mode, 2)
        mode_layout.addWidget(self.vba_mode)
        
        self.output_mode = QRadioButton("Validate Output")
        self.mode_group.addButton(self.output_mode, 3)
        mode_layout.addWidget(self.output_mode)
        
        layout.addWidget(mode_group)
        
        # Input section
        input_group = QGroupBox("Enter Code to Verify")
        input_layout = QVBoxLayout(input_group)
        
        code_label = QLabel("Formula or VBA code:")
        input_layout.addWidget(code_label)
        
        self.code_input = QTextEdit()
        self.code_input.setPlaceholderText("Enter your formula or VBA code here...")
        self.code_input.setMaximumHeight(100)
        input_layout.addWidget(self.code_input)
        
        # Expected result
        expected_label = QLabel("Expected result (optional):")
        input_layout.addWidget(expected_label)
        
        self.expected_input = QLineEdit()
        self.expected_input.setPlaceholderText("What should the formula return?")
        input_layout.addWidget(self.expected_input)
        
        layout.addWidget(input_group)
        
        # Verify button
        self.verify_btn = QPushButton("🔍 Verify Now")
        self.verify_btn.clicked.connect(self.verify_code)
        layout.addWidget(self.verify_btn)
        
        # Results section
        results_group = QGroupBox("Verification Results")
        results_layout = QVBoxLayout(results_group)
        
        # Confidence score
        self.confidence_label = QLabel("Confidence Score: --")
        self.confidence_label.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
        results_layout.addWidget(self.confidence_label)
        
        # Results browser
        self.results_browser = QTextBrowser()
        self.results_browser.setOpenExternalLinks(True)
        results_layout.addWidget(self.results_browser)
        
        layout.addWidget(results_group)
    
    def verify_code(self):
        """Verify the code using AI"""
        code = self.code_input.toPlainText().strip()
        expected = self.expected_input.text().strip()
        
        if not code:
            QMessageBox.warning(self, "Input Required", "Please enter code to verify.")
            return
        
        mode = self.mode_group.checkedId()
        
        self.verify_btn.setEnabled(False)
        self.verify_btn.setText("Verifying...")
        self.results_browser.setPlainText("Analyzing code...")
        
        def run_verify():
            if mode == 1:  # Formula
                result = self.ai_engine.verify_formula(code, expected)
            elif mode == 2:  # VBA
                result = self.ai_engine.verify_vba_code(code, expected)
            else:  # Output validation
                result = self.ai_engine.validate_output(code, expected)
            
            self.result_data = result
            self.display_results(result)
            self.verify_btn.setEnabled(True)
            self.verify_btn.setText("🔍 Verify Now")
        
        thread = threading.Thread(target=run_verify, daemon=True)
        thread.start()
    
    def display_results(self, result):
        """Display verification results"""
        if "raw_response" in result:
            self.results_browser.setPlainText(result["raw_response"])
            self.confidence_label.setText("Confidence Score: N/A")
            return
        
        # Update confidence score
        confidence = result.get('confidence_score', 0)
        is_valid = result.get('is_valid', result.get('is_correct', False))
        
        if confidence >= 80:
            color = "#4CAF50"
            status = "PASSED"
        elif confidence >= 50:
            color = "#FF9800"
            status = "WARNING"
        else:
            color = "#f44336"
            status = "FAILED"
        
        self.confidence_label.setText(f"Confidence Score: {confidence}% - {status}")
        self.confidence_label.setStyleSheet(f"color: {color}; font-size: 14px; font-weight: bold;")
        
        # Build results HTML
        html = f"<h3 style='color: {color};'>Verification {status}</h3>"
        
        # Syntax check
        syntax = result.get('syntax_check', result.get('syntax_errors', {}))
        if isinstance(syntax, dict):
            syntax_valid = syntax.get('valid', True)
            syntax_icon = "✅" if syntax_valid else "❌"
            html += f"<p>{syntax_icon} <strong>Syntax:</strong> {'Valid' if syntax_valid else 'Issues found'}</p>"
            
            issues = syntax.get('issues', [])
            if issues:
                html += "<ul style='color: #f44336;'>"
                for issue in issues:
                    html += f"<li>{issue}</li>"
                html += "</ul>"
        
        # Logic check
        logic = result.get('logic_check', {})
        if logic:
            logic_valid = logic.get('valid', True)
            logic_icon = "✅" if logic_valid else "❌"
            html += f"<p>{logic_icon} <strong>Logic:</strong> {logic.get('explanation', 'N/A')}</p>"
        
        # Edge cases
        edge_cases = result.get('edge_cases', {})
        if edge_cases:
            html += "<p><strong>Edge Cases:</strong></p><ul>"
            for case, handled in edge_cases.items():
                icon = "✅" if handled else "⚠️"
                html += f"<li>{icon} {case.replace('_', ' ').title()}</li>"
            html += "</ul>"
        
        # Issues
        issues = result.get('issues', result.get('syntax_errors', []))
        if issues:
            html += "<p style='color: #f44336;'><strong>Issues Found:</strong></p><ul>"
            for issue in issues:
                if isinstance(issue, dict):
                    html += f"<li>{issue.get('error', issue.get('description', str(issue)))}</li>"
                else:
                    html += f"<li>{issue}</li>"
            html += "</ul>"
        
        # Warnings
        warnings = result.get('warnings', [])
        if warnings:
            html += "<p style='color: #FF9800;'><strong>Warnings:</strong></p><ul>"
            for warning in warnings:
                if isinstance(warning, dict):
                    html += f"<li>{warning.get('warning', warning.get('description', str(warning)))}</li>"
                else:
                    html += f"<li>{warning}</li>"
            html += "</ul>"
        
        # Verified/Fixed code
        verified = result.get('verified_code', result.get('verified_formula', ''))
        if verified:
            html += f"""
            <div style='background: #0f3460; padding: 12px; border-radius: 8px; margin-top: 12px;'>
                <strong>Corrected Version:</strong><br>
                <code style='background: #16213e; padding: 8px; border-radius: 4px; display: block; margin-top: 8px;'>
                    {verified}
                </code>
            </div>
            """
        
        # Verdict
        verdict = result.get('final_verdict', result.get('verdict', ''))
        if verdict:
            verdict_color = "#4CAF50" if "PASS" in verdict.upper() else "#f44336"
            html += f"<div style='background: {verdict_color}20; padding: 12px; border-radius: 8px; border-left: 4px solid {verdict_color}; margin-top: 12px;'><strong>Final Verdict:</strong> {verdict}</div>"
        
        self.results_browser.setHtml(html)


class FloatingAssistant(QMainWindow):
    """Main floating assistant window"""
    
    def __init__(self):
        super().__init__()
        
        self.signals = AssistantSignals()
        self.signals.response_ready.connect(self.on_response_ready)
        self.signals.status_update.connect(self.on_status_update)
        self.signals.file_detected.connect(self.on_file_detected)
        self.signals.error_occurred.connect(self.on_error)
        
        self.ai_engine = AIEngine()
        self.excel_parser = ExcelParser()
        self.theme = ThemeManager()
        
        self.current_file = None
        self.is_processing = False
        
        self.init_ui()
        self.init_system_tray()
        self.init_file_monitor()
        
    def init_ui(self):
        """Initialize the user interface"""
        self.setWindowTitle("Exceller - AI Excel Assistant")
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint |
            Qt.WindowType.WindowStaysOnTopHint |
            Qt.WindowType.Tool
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setMinimumSize(380, 520)
        self.setMaximumSize(380, 700)
        
        # Main container
        self.container = QWidget()
        self.container.setObjectName("mainContainer")
        self.setCentralWidget(self.container)
        
        # Main layout
        main_layout = QVBoxLayout(self.container)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # Title bar
        title_bar = self.create_title_bar()
        main_layout.addWidget(title_bar)
        
        # Content area
        content = QWidget()
        content.setObjectName("contentArea")
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(12, 8, 12, 12)
        content_layout.setSpacing(8)
        
        # File info section
        self.file_info = self.create_file_info()
        content_layout.addWidget(self.file_info)
        
        # Quick actions
        quick_actions = self.create_quick_actions()
        content_layout.addWidget(quick_actions)
        
        # Chat area
        self.chat_area = QTextEdit()
        self.chat_area.setObjectName("chatArea")
        self.chat_area.setPlaceholderText("Ask anything about your Excel data...")
        self.chat_area.setReadOnly(True)
        content_layout.addWidget(self.chat_area)
        
        # Input area
        input_area = self.create_input_area()
        content_layout.addWidget(input_area)
        
        main_layout.addWidget(content)
        
        # Apply styles
        self.setStyleSheet(self.theme.get_main_stylesheet())
        
        # Enable drag
        self.drag_position = None
        
    def create_title_bar(self):
        """Create custom title bar"""
        title_bar = QWidget()
        title_bar.setObjectName("titleBar")
        title_bar.setFixedHeight(40)
        
        layout = QHBoxLayout(title_bar)
        layout.setContentsMargins(12, 0, 8, 0)
        
        # App icon and title
        icon_label = QLabel("📊")
        icon_label.setFont(QFont("Segoe UI", 14))
        layout.addWidget(icon_label)
        
        title = QLabel("Exceller")
        title.setObjectName("appTitle")
        layout.addWidget(title)
        
        layout.addStretch()
        
        # Pin button (always on top)
        self.pin_btn = QPushButton("📌")
        self.pin_btn.setObjectName("pinBtn")
        self.pin_btn.setFixedSize(28, 28)
        self.pin_btn.setToolTip("Pin to top")
        self.pin_btn.clicked.connect(self.toggle_pin)
        layout.addWidget(self.pin_btn)
        
        # Minimize button
        min_btn = QPushButton("−")
        min_btn.setObjectName("minBtn")
        min_btn.setFixedSize(28, 28)
        min_btn.clicked.connect(self.minimize_to_tray)
        layout.addWidget(min_btn)
        
        # Close button
        close_btn = QPushButton("×")
        close_btn.setObjectName("closeBtn")
        close_btn.setFixedSize(28, 28)
        close_btn.clicked.connect(self.close)
        layout.addWidget(close_btn)
        
        return title_bar
    
    def create_file_info(self):
        """Create file info section"""
        frame = QFrame()
        frame.setObjectName("fileInfoFrame")
        frame.setFixedHeight(60)
        
        layout = QHBoxLayout(frame)
        layout.setContentsMargins(10, 8, 10, 8)
        
        # File icon
        file_icon = QLabel("📁")
        file_icon.setFont(QFont("Segoe UI", 18))
        layout.addWidget(file_icon)
        
        # File details
        details = QVBoxLayout()
        details.setSpacing(2)
        
        self.file_name = QLabel("No file loaded")
        self.file_name.setObjectName("fileName")
        details.addWidget(self.file_name)
        
        self.file_stats = QLabel("Drop an Excel file or click Browse")
        self.file_stats.setObjectName("fileStats")
        details.addWidget(self.file_stats)
        
        layout.addLayout(details)
        layout.addStretch()
        
        # Browse button
        browse_btn = QPushButton("Browse")
        browse_btn.setObjectName("browseBtn")
        browse_btn.clicked.connect(self.browse_file)
        layout.addWidget(browse_btn)
        
        return frame
    
    def create_quick_actions(self):
        """Create quick action buttons"""
        # Main actions frame
        main_frame = QWidget()
        main_layout = QVBoxLayout(main_frame)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(6)
        
        # Row 1: Data actions
        row1 = QFrame()
        row1.setObjectName("actionsFrame")
        layout1 = QHBoxLayout(row1)
        layout1.setContentsMargins(4, 4, 4, 4)
        layout1.setSpacing(6)
        
        actions_row1 = [
            ("📈", "Analyze", self.analyze_data),
            ("🔍", "Find Issues", self.find_issues),
            ("📊", "Stats", self.get_statistics),
            ("🧹", "Clean", self.clean_data),
        ]
        
        for icon, text, callback in actions_row1:
            btn = QPushButton(f"{icon}\n{text}")
            btn.setObjectName("actionBtn")
            btn.setFixedSize(75, 50)
            btn.clicked.connect(callback)
            layout1.addWidget(btn)
        
        main_layout.addWidget(row1)
        
        # Row 2: Formula tools
        row2 = QFrame()
        row2.setObjectName("actionsFrame")
        layout2 = QHBoxLayout(row2)
        layout2.setContentsMargins(4, 4, 4, 4)
        layout2.setSpacing(6)
        
        actions_row2 = [
            ("🔧", "Fix Formula", self.open_formula_fixer),
            ("💰", "Price Calc", self.open_price_calculator),
            ("✅", "Verify", self.open_verifier),
            ("📝", "Formula", self.get_formula_help),
        ]
        
        for icon, text, callback in actions_row2:
            btn = QPushButton(f"{icon}\n{text}")
            btn.setObjectName("actionBtn")
            btn.setFixedSize(75, 50)
            btn.clicked.connect(callback)
            layout2.addWidget(btn)
        
        main_layout.addWidget(row2)
        
        return main_frame
    
    def create_input_area(self):
        """Create input area with text field and send button"""
        frame = QWidget()
        layout = QHBoxLayout(frame)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        
        # Input field
        self.input_field = QLineEdit()
        self.input_field.setObjectName("inputField")
        self.input_field.setPlaceholderText("Ask about your data...")
        self.input_field.returnPressed.connect(self.send_message)
        layout.addWidget(self.input_field)
        
        # Send button
        self.send_btn = QPushButton("→")
        self.send_btn.setObjectName("sendBtn")
        self.send_btn.setFixedSize(40, 40)
        self.send_btn.clicked.connect(self.send_message)
        layout.addWidget(self.send_btn)
        
        return frame
    
    def init_system_tray(self):
        """Initialize system tray icon"""
        self.tray_icon = QSystemTrayIcon(self)
        
        # Create tray icon
        pixmap = QPixmap(32, 32)
        pixmap.fill(QColor("#4CAF50"))
        painter = QPainter(pixmap)
        painter.setPen(QPen(QColor("white"), 2))
        painter.setFont(QFont("Arial", 16, QFont.Weight.Bold))
        painter.drawText(pixmap.rect(), Qt.AlignmentFlag.AlignCenter, "E")
        painter.end()
        
        self.tray_icon.setIcon(QIcon(pixmap))
        self.tray_icon.setToolTip("Exceller - AI Excel Assistant")
        
        # Create tray menu
        tray_menu = QMenu()
        tray_menu.setStyleSheet("""
            QMenu {
                background-color: #2b2b2b;
                border: 1px solid #3d3d3d;
                border-radius: 8px;
                padding: 4px;
            }
            QMenu::item {
                color: white;
                padding: 8px 24px;
                border-radius: 4px;
            }
            QMenu::item:selected {
                background-color: #4CAF50;
            }
        """)
        
        show_action = QAction("Show Assistant", self)
        show_action.triggered.connect(self.show)
        tray_menu.addAction(show_action)
        
        tray_menu.addSeparator()
        
        quit_action = QAction("Quit", self)
        quit_action.triggered.connect(self.quit_app)
        tray_menu.addAction(quit_action)
        
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.tray_activated)
        self.tray_icon.show()
    
    def init_file_monitor(self):
        """Initialize file system monitor"""
        self.file_monitor = FileMonitor()
        self.file_monitor.file_changed.connect(self.on_file_changed)
        
    def toggle_pin(self):
        """Toggle always on top"""
        current_flags = self.windowFlags()
        if current_flags & Qt.WindowType.WindowStaysOnTopHint:
            self.setWindowFlags(current_flags & ~Qt.WindowType.WindowStaysOnTopHint)
            self.pin_btn.setText("📌")
        else:
            self.setWindowFlags(current_flags | Qt.WindowType.WindowStaysOnTopHint)
            self.pin_btn.setText("📍")
        self.show()
    
    def minimize_to_tray(self):
        """Minimize to system tray"""
        self.hide()
        self.tray_icon.showMessage(
            "Exceller",
            "Assistant is running in the background",
            QSystemTrayIcon.MessageIcon.Information,
            2000
        )
    
    def tray_activated(self, reason):
        """Handle tray icon activation"""
        if reason == QSystemTrayIcon.ActivationReason.DoubleClick:
            self.show()
            self.raise_()
            self.activateWindow()
    
    def browse_file(self):
        """Open file browser dialog"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Excel File",
            "",
            "Excel Files (*.xlsx *.xls *.csv);;All Files (*)"
        )
        
        if file_path:
            self.load_file(file_path)
    
    def load_file(self, file_path):
        """Load an Excel file"""
        self.current_file = file_path
        file_name = Path(file_path).name
        
        self.file_name.setText(file_name)
        
        # Get file stats
        try:
            stats = self.excel_parser.get_file_stats(file_path)
            self.file_stats.setText(
                f"{stats['sheets']} sheets | {stats['rows']} rows | {stats['columns']} columns"
            )
            
            # Start monitoring file
            self.file_monitor.watch_file(file_path)
            
            # Add to chat
            self.add_message("system", f"Loaded: {file_name}")
            
        except Exception as e:
            self.file_stats.setText("Error reading file")
            self.add_message("error", f"Error loading file: {str(e)}")
    
    def send_message(self):
        """Send message to AI"""
        message = self.input_field.text().strip()
        
        if not message:
            return
        
        if self.is_processing:
            return
        
        self.input_field.clear()
        self.add_message("user", message)
        
        self.is_processing = True
        self.send_btn.setEnabled(False)
        
        # Process in background thread
        thread = threading.Thread(
            target=self.process_message,
            args=(message,),
            daemon=True
        )
        thread.start()
    
    def process_message(self, message):
        """Process message with AI"""
        try:
            context = ""
            
            if self.current_file:
                # Get file context
                file_data = self.excel_parser.get_file_context(self.current_file)
                context = f"Current file: {Path(self.current_file).name}\n{file_data}"
            
            response = self.ai_engine.chat(message, context)
            self.signals.response_ready.emit(response)
            
        except Exception as e:
            self.signals.error_occurred.emit(str(e))
        finally:
            self.is_processing = False
    
    def on_response_ready(self, response):
        """Handle AI response"""
        self.add_message("assistant", response)
        self.send_btn.setEnabled(True)
    
    def on_status_update(self, status):
        """Handle status update"""
        self.file_stats.setText(status)
    
    def on_file_detected(self, file_path):
        """Handle file detection"""
        self.load_file(file_path)
    
    def on_error(self, error):
        """Handle error"""
        self.add_message("error", f"Error: {error}")
        self.send_btn.setEnabled(True)
    
    def on_file_changed(self, file_path):
        """Handle file changes"""
        if file_path == self.current_file:
            self.add_message("system", "File changed, refreshing...")
            self.load_file(file_path)
    
    def add_message(self, role, content):
        """Add message to chat area"""
        timestamp = datetime.now().strftime("%H:%M")
        
        if role == "user":
            html = f"""
            <div style="margin: 8px 0; text-align: right;">
                <span style="background: #4CAF50; color: white; padding: 8px 12px; 
                       border-radius: 12px 12px 4px 12px; display: inline-block; 
                       max-width: 85%; text-align: left;">
                    {content}
                </span>
                <br><small style="color: #888;">{timestamp}</small>
            </div>
            """
        elif role == "assistant":
            html = f"""
            <div style="margin: 8px 0;">
                <span style="background: #3d3d3d; color: white; padding: 8px 12px; 
                       border-radius: 12px 12px 12px 4px; display: inline-block; 
                       max-width: 85%;">
                    {content.replace(chr(10), '<br>')}
                </span>
                <br><small style="color: #888;">{timestamp}</small>
            </div>
            """
        elif role == "system":
            html = f"""
            <div style="margin: 8px 0; text-align: center;">
                <span style="background: #1a1a2e; color: #888; padding: 4px 12px; 
                       border-radius: 8px; font-size: 11px;">
                    {content}
                </span>
            </div>
            """
        elif role == "error":
            html = f"""
            <div style="margin: 8px 0; text-align: center;">
                <span style="background: #5c2020; color: #ff6b6b; padding: 4px 12px; 
                       border-radius: 8px; font-size: 11px;">
                    {content}
                </span>
            </div>
            """
        
        self.chat_area.append(html)
        
        # Scroll to bottom
        scrollbar = self.chat_area.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
    
    # Quick action handlers
    def analyze_data(self):
        """Analyze current file data"""
        if not self.current_file:
            self.add_message("system", "Please load a file first")
            return
        
        self.input_field.setText("Analyze this data and provide insights")
        self.send_message()
    
    def find_issues(self):
        """Find issues in current data"""
        if not self.current_file:
            self.add_message("system", "Please load a file first")
            return
        
        self.input_field.setText("Find any issues, errors, or inconsistencies in this data")
        self.send_message()
    
    def get_statistics(self):
        """Get statistics for current data"""
        if not self.current_file:
            self.add_message("system", "Please load a file first")
            return
        
        self.input_field.setText("Calculate and display statistics for this data")
        self.send_message()
    
    def clean_data(self):
        """Clean current data"""
        if not self.current_file:
            self.add_message("system", "Please load a file first")
            return
        
        self.input_field.setText("Suggest how to clean and improve this data")
        self.send_message()
    
    def open_formula_fixer(self):
        """Open the formula error fixer dialog"""
        dialog = FormulaFixerDialog(self.ai_engine, self)
        dialog.exec()
    
    def open_price_calculator(self):
        """Open the price calculation helper dialog"""
        dialog = PriceCalculatorDialog(self.ai_engine, self)
        dialog.exec()
    
    def open_verifier(self):
        """Open the verification dialog"""
        dialog = VerificationDialog(self.ai_engine, self)
        dialog.exec()
    
    def get_formula_help(self):
        """Get formula help"""
        self.input_field.setText("Help me write an Excel formula")
        self.send_message()
    
    def mousePressEvent(self, event):
        """Handle mouse press for dragging"""
        if event.button() == Qt.MouseButton.LeftButton:
            self.drag_position = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
            event.accept()
    
    def mouseMoveEvent(self, event):
        """Handle mouse move for dragging"""
        if event.buttons() & Qt.MouseButton.LeftButton and self.drag_position:
            self.move(event.globalPosition().toPoint() - self.drag_position)
            event.accept()
    
    def closeEvent(self, event):
        """Handle close event"""
        event.ignore()
        self.minimize_to_tray()
    
    def quit_app(self):
        """Quit the application"""
        self.tray_icon.hide()
        QApplication.quit()


def main():
    """Main entry point"""
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    
    # Set app style
    app.setStyle("Fusion")
    
    # Create and show assistant
    assistant = FloatingAssistant()
    assistant.show()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
