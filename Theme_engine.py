import tkinter as tk
from tkinter import ttk

# ==============================================================================
# GLOBAL DESIGN STYLE PALETTE CONSTANTS
# ==============================================================================
COLOR_PRIMARY_DARK  = "#2B4C7E"  # Deep slate blue header
COLOR_WORKSPACE_BG  = "#F4F7FA"  # Soft blue-gray canvas workspace
COLOR_HIGHLIGHT_ALT = "#E1EBF5"  # Muted highlight accent hue
COLOR_WHITE         = "#FFFFFF"
COLOR_CONSOLE_BG    = "#1E1E1E"  # Dark console backing
COLOR_CONSOLE_FG    = "#F0F0F0"  # Light gray terminal typography

FONT_FAMILY = "Segoe UI"
FONT_FIXED  = "Courier"

# Console Dynamic Logging Tag Color Matrix
CONSOLE_TAGS = {
    "PASS": {"foreground": "#4CAF50"},  # Crisp Green
    "FAIL": {"foreground": "#F44336"},  # Vivid Red
    "INFO": {"foreground": "#00BCD4"}   # Bright Cyan Info
}

def apply_global_theme(root):
    """
    Sets up the window base configuration and applies the unified global 
    TTK theme definitions across elements.
    """
    style = ttk.Style()
    
    # Force cleaner 'clam' rendering engine rules if available natively
    if 'clam' in style.theme_names():
        style.theme_use('clam')
        
    root.configure(bg=COLOR_WORKSPACE_BG)
    
    # Frame Styles
    style.configure("TFrame", background=COLOR_WORKSPACE_BG)
    style.configure("Header.TFrame", background=COLOR_PRIMARY_DARK)
    
    # Typography Styles
    style.configure("HeaderLabel.TLabel", 
                    background=COLOR_PRIMARY_DARK, 
                    foreground=COLOR_WHITE, 
                    font=(FONT_FAMILY, 12, "bold"))
    
    style.configure("Standard.TLabel", 
                    background=COLOR_WORKSPACE_BG, 
                    font=(FONT_FAMILY, 10))
    
    style.configure("Path.TLabel", 
                    background=COLOR_WHITE, 
                    font=(FONT_FAMILY, 9, "italic"), 
                    relief="solid", 
                    borderwidth=1)
    
    # Button Layout Formats
    style.configure("Action.TButton", 
                    font=(FONT_FAMILY, 10, "bold"), 
                    padding=5)
    
    style.configure("Run.TButton", 
                    font=(FONT_FAMILY, 11, "bold"), 
                    background=COLOR_PRIMARY_DARK, 
                    foreground=COLOR_WHITE, 
                    padding=6)
    
    return style

def style_console_widget(text_widget):
    """
    Applies custom styling properties and color tagging maps to non-TTK Text components.
    """
    text_widget.configure(
        bg=COLOR_CONSOLE_BG, 
        fg=COLOR_CONSOLE_FG, 
        font=(FONT_FIXED, 10), 
        wrap="word", 
        relief="sunken", 
        bd=2
    )
    
    # Inject text styling configurations programmatically
    for tag_name, configurations in CONSOLE_TAGS.items():
        text_widget.tag_config(tag_name, **configurations)