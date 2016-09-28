Imports System.Text

Public Module Win32API

    'Button Control Messages
    Public Const BM_CLICK = &HF5        'Button Click message
    Public Const BM_GETCHECK = &HF0     'Radio button check
    Public Const CB_GETCURSEL = &H147   'Combo box get selected item index
    Public Const CB_GETCOUNT = &H146    'Combo box get count of items
    Public Const CB_GETLBTEXT = &H148   'Combo box get selected item string
    Public Const WM_GETTEXT = 13        'Text box get text


    Public Declare Auto Function FindWindow Lib "user32.dll" _
        Alias "FindWindow" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer

    Public Declare Auto Function FindWindowEx Lib "user32.dll" _
        Alias "FindWindowEx" (ByVal hwndParent As Integer, _
            ByVal hwndChildAfter As Integer, _
            ByVal lpClassName As String, _
            ByVal lpWindowName As String) As Integer

    Public Declare Sub GetWindowText Lib "user32.dll" _
            Alias "GetWindowTextA" (ByVal hWnd As Integer, ByVal lpString As StringBuilder, ByVal nMaxCount As Integer)

    Public Declare Function GetClassName Lib "user32.dll" _
        Alias "GetClassNameA" (ByVal hwnd As Integer, ByVal lpClassName As StringBuilder, ByVal cch As Integer) As Integer

    Public Declare Function SendMessage Lib "user32.dll" _
        Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer

    Public Declare Function SendMessageS Lib "user32.dll" _
        Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As StringBuilder) As Integer

    Public Declare Function PostMessage Lib "user32.dll" _
        Alias "PostMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer

    Public Delegate Function EnumWindowsCallback(ByVal hWnd As Integer, ByVal lParam As Integer) As Boolean

    Public Declare Function EnumWindows Lib "user32.dll" _
        Alias "EnumWindows" (ByVal callback As EnumWindowsCallback, ByVal lParam As Integer) As Integer

    Public Declare Function EnumChildWindows Lib "user32.dll" _
        Alias "EnumChildWindows" (ByVal hWndParent As Integer, ByVal callback As EnumWindowsCallback, ByVal lParam As Integer) As Integer
End Module
