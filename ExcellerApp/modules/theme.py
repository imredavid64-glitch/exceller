"""
Theme Manager - Dark theme styling
Provides consistent dark theme for the application
"""


class ThemeManager:
    """Manages application theme and styling"""
    
    # Color palette
    COLORS = {
        'background': '#1a1a2e',
        'surface': '#16213e',
        'card': '#0f3460',
        'primary': '#4CAF50',
        'primary_hover': '#45a049',
        'secondary': '#e94560',
        'text': '#ffffff',
        'text_secondary': '#b0b0b0',
        'text_muted': '#666666',
        'border': '#2d2d44',
        'success': '#4CAF50',
        'warning': '#ff9800',
        'error': '#f44336',
        'info': '#2196F3',
        'chat_user': '#4CAF50',
        'chat_assistant': '#2d2d44',
        'chat_system': '#1a1a2e',
    }
    
    def get_main_stylesheet(self) -> str:
        """Get the main application stylesheet"""
        return f"""
            /* Main Container */
            #mainContainer {{
                background-color: {self.COLORS['background']};
                border: 1px solid {self.COLORS['border']};
                border-radius: 12px;
            }}
            
            /* Title Bar */
            #titleBar {{
                background-color: {self.COLORS['surface']};
                border-top-left-radius: 12px;
                border-top-right-radius: 12px;
                border-bottom: 1px solid {self.COLORS['border']};
            }}
            
            #appTitle {{
                color: {self.COLORS['text']};
                font-size: 14px;
                font-weight: bold;
            }}
            
            /* Title bar buttons */
            #pinBtn, #minBtn, #closeBtn {{
                background-color: transparent;
                border: none;
                border-radius: 4px;
                color: {self.COLORS['text_secondary']};
                font-size: 14px;
            }}
            
            #pinBtn:hover {{
                background-color: {self.COLORS['card']};
            }}
            
            #minBtn:hover {{
                background-color: {self.COLORS['card']};
            }}
            
            #closeBtn:hover {{
                background-color: {self.COLORS['error']};
                color: white;
            }}
            
            /* Content Area */
            #contentArea {{
                background-color: transparent;
            }}
            
            /* File Info */
            #fileInfoFrame {{
                background-color: {self.COLORS['surface']};
                border: 1px solid {self.COLORS['border']};
                border-radius: 8px;
            }}
            
            #fileName {{
                color: {self.COLORS['text']};
                font-size: 13px;
                font-weight: bold;
            }}
            
            #fileStats {{
                color: {self.COLORS['text_secondary']};
                font-size: 11px;
            }}
            
            #browseBtn {{
                background-color: {self.COLORS['primary']};
                color: white;
                border: none;
                border-radius: 6px;
                padding: 6px 12px;
                font-size: 11px;
                font-weight: bold;
            }}
            
            #browseBtn:hover {{
                background-color: {self.COLORS['primary_hover']};
            }}
            
            /* Quick Actions */
            #actionsFrame {{
                background-color: transparent;
            }}
            
            #actionBtn {{
                background-color: {self.COLORS['surface']};
                border: 1px solid {self.COLORS['border']};
                border-radius: 8px;
                color: {self.COLORS['text']};
                font-size: 11px;
            }}
            
            #actionBtn:hover {{
                background-color: {self.COLORS['card']};
                border-color: {self.COLORS['primary']};
            }}
            
            /* Chat Area */
            #chatArea {{
                background-color: {self.COLORS['background']};
                border: 1px solid {self.COLORS['border']};
                border-radius: 8px;
                color: {self.COLORS['text']};
                font-size: 13px;
                padding: 8px;
            }}
            
            #chatArea::selection {{
                background-color: {self.COLORS['primary']};
            }}
            
            /* Input Area */
            #inputField {{
                background-color: {self.COLORS['surface']};
                border: 1px solid {self.COLORS['border']};
                border-radius: 20px;
                color: {self.COLORS['text']};
                padding: 10px 16px;
                font-size: 13px;
            }}
            
            #inputField:focus {{
                border-color: {self.COLORS['primary']};
            }}
            
            #inputField::placeholder {{
                color: {self.COLORS['text_muted']};
            }}
            
            #sendBtn {{
                background-color: {self.COLORS['primary']};
                border: none;
                border-radius: 20px;
                color: white;
                font-size: 16px;
                font-weight: bold;
            }}
            
            #sendBtn:hover {{
                background-color: {self.COLORS['primary_hover']};
            }}
            
            #sendBtn:disabled {{
                background-color: {self.COLORS['text_muted']};
            }}
            
            /* Scrollbar */
            QScrollBar:vertical {{
                border: none;
                background-color: {self.COLORS['background']};
                width: 8px;
                margin: 0;
            }}
            
            QScrollBar::handle:vertical {{
                background-color: {self.COLORS['border']};
                min-height: 30px;
                border-radius: 4px;
            }}
            
            QScrollBar::handle:vertical:hover {{
                background-color: {self.COLORS['text_muted']};
            }}
            
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
                height: 0;
            }}
            
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{
                background: none;
            }}
        """
    
    def get_chat_bubble_style(self, role: str) -> str:
        """Get chat bubble style for different roles"""
        styles = {
            'user': f"""
                background-color: {self.COLORS['chat_user']};
                color: white;
                border-radius: 12px 12px 4px 12px;
                padding: 8px 12px;
            """,
            'assistant': f"""
                background-color: {self.COLORS['chat_assistant']};
                color: white;
                border-radius: 12px 12px 12px 4px;
                padding: 8px 12px;
            """,
            'system': f"""
                background-color: {self.COLORS['chat_system']};
                color: {self.COLORS['text_muted']};
                border-radius: 8px;
                padding: 4px 12px;
                font-size: 11px;
            """
        }
        return styles.get(role, styles['system'])
    
    @property
    def colors(self):
        """Get color palette"""
        return self.COLORS.copy()
