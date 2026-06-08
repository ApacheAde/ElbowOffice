================================================================================
                          ELBOWOFFICE WRITER
                     A Simple Word Processor in Python
================================================================================

Welcome to ElbowOffice!

ElbowOffice is a clean, fully-featured word processor written entirely in 
Python using PyQt6. It provides a comfortable writing experience similar to 
OpenOffice Writer, LibreOffice Writer, or Microsoft Word, but lightweight and 
easy to run from source.

It supports rich text formatting, multiple document tabs, images, tables, 
PDF export, printing, and more.

--------------------------------------------------------------------------------
INSTALLATION (Windows)
--------------------------------------------------------------------------------

1. Make sure you have Python installed.
   Recommended (one command in PowerShell):
        winget install Python.Python.3.12

   Alternative options:
   - Use Scoop:   scoop install python
   - Or download from https://www.python.org/downloads/
     (Check "Add python.exe to PATH" during installation)

2. Install the required library:
        pip install PyQt6

3. (Optional but recommended) For full .docx (Microsoft Word) file support:
        pip install python-docx

4. Run the program:
        cd Desktop
        python elbowoffice.py

   You can also double-click elbowoffice.py once the dependencies are installed.

--------------------------------------------------------------------------------
MAIN FEATURES
--------------------------------------------------------------------------------

• Multiple document tabs (open and edit several files at once)
• Full rich-text formatting:
    - Bold, Italic, Underline, Strikethrough
    - Font family and size selection
    - Text color and highlight color
    - Paragraph alignment (Left, Center, Right, Justify)
    - Bullet lists and numbered lists
    - Indentation controls
• Insert images from your computer
• Insert and edit tables
• Find & Replace with options (case sensitive, whole words)
• Zoom in / out / reset
• Export to PDF
• Direct printing with print preview
• Save / Load in multiple formats:
    - HTML (.html) - best for preserving formatting
    - Plain Text (.txt)
    - Microsoft Word (.docx) - requires python-docx
• Automatic word and character count
• Cursor position display
• Unsaved changes protection when closing tabs or exiting
• Recent files menu
• Keyboard shortcuts for almost everything (Ctrl+B, Ctrl+I, Ctrl+S, etc.)

--------------------------------------------------------------------------------
KEYBOARD SHORTCUTS
--------------------------------------------------------------------------------

File:
    Ctrl + N          New document
    Ctrl + O          Open document
    Ctrl + S          Save
    Ctrl + Shift + S  Save As
    Ctrl + P          Print
    Ctrl + W          Close current tab
    Ctrl + Q          Quit ElbowOffice

Edit:
    Ctrl + Z          Undo
    Ctrl + Y          Redo
    Ctrl + X          Cut
    Ctrl + C          Copy
    Ctrl + V          Paste
    Ctrl + F          Find & Replace
    Ctrl + A          Select All

Formatting:
    Ctrl + B          Bold
    Ctrl + I          Italic
    Ctrl + U          Underline
    Ctrl + ]          Indent more
    Ctrl + [          Indent less

View:
    Ctrl + +          Zoom In
    Ctrl + -          Zoom Out
    Ctrl + 0          Reset Zoom

--------------------------------------------------------------------------------
FILE FORMATS
--------------------------------------------------------------------------------

When saving or opening, you can choose:

- .html / .htm   : Recommended for rich text (preserves fonts, colors, lists, 
                   images, tables, etc.)
- .txt           : Plain text only (no formatting)
- .docx          : Microsoft Word format (requires "python-docx" package)

ElbowOffice uses HTML internally for the highest fidelity when working with 
formatted documents.

--------------------------------------------------------------------------------
TIPS
--------------------------------------------------------------------------------

- The first time you open ElbowOffice, it shows a short welcome document 
  with examples. Feel free to edit or delete it.

- Use HTML format when you want to keep complex formatting (recommended).

- For compatibility with Microsoft Word users, save as .docx when 
  python-docx is installed.

- You can insert images of almost any common format (PNG, JPG, GIF, etc.).

- Tables can be created via Insert > Table. You can type inside cells 
  just like a normal document.

- The status bar at the bottom always shows word count, character count, 
  and your current cursor position.

- ElbowOffice remembers your recent files and window size between sessions.

--------------------------------------------------------------------------------
FILES ON YOUR DESKTOP
--------------------------------------------------------------------------------

elbowoffice.py           - The main ElbowOffice word processor application
ElbowOffice_README.txt   - This instruction file
(High scores or settings are stored separately via Qt's QSettings)

--------------------------------------------------------------------------------
TECHNICAL NOTES
--------------------------------------------------------------------------------

- Built with Python + PyQt6
- Uses QTextEdit as the rich text engine (very powerful)
- All code is contained in a single .py file (no external assets required)
- Graceful startup messages if dependencies are missing
- Optional DOCX support via the python-docx library

--------------------------------------------------------------------------------
ENJOY WRITING!
--------------------------------------------------------------------------------

ElbowOffice was created as a fun, educational project to demonstrate building 
a real word processor from scratch in Python.

Type away, format to your heart's content, and export when you're done!

If you find any issues or have feature ideas, feel free to improve the script.

Happy writing!

================================================================================
