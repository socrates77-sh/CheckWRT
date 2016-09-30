Imports System.Text

Public Module Win32API

    Public Const BM_CLICK As Integer = &HF5        'Button Click message
    Public Const BM_GETCHECK As Integer = &HF0     'Radio button check
    Public Const CB_GETCURSEL As Integer = &H147   'Combo box get selected item index
    Public Const CB_GETCOUNT As Integer = &H146    'Combo box get count of items
    Public Const CB_GETLBTEXT As Integer = &H148   'Combo box get selected item string
    Public Const WM_GETTEXT As Integer = &HD       'Text box get text
    Public Const WM_SYSCOMMAND As Integer = &H112
    Public Const WM_COMMAND As Integer = &H111
    Public Const WM_KEYDOWN As Integer = &H100
    Public Const VK_RETURN As Integer = &HD
    'TTreeView const
    Public Const TV_FIRST As Integer = &H1100
    Public Const TVM_SETBKCOLOR As Integer = TV_FIRST + 29
    Public Const TVM_GETITEM As Integer = TV_FIRST + 12
    Public Const TVIF_HANDLE As Integer = &H10
    Public Const TVIS_SELECTED As Integer = &H2
    Public Const TVIF_STATE As Integer = &H8



    Public Structure TVITEM
        Dim mask As Long
        Dim HTreeItem As Integer
        Dim state As Integer
        Dim stateMask As Integer
        Dim pszText As Integer
        Dim cchTextMax As Integer
        Dim iImage As Integer
        Dim iSelectedImage As Integer
        Dim cChildren As Integer
        Dim lParam As Integer
    End Structure

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

    Public Declare Function SendMessageTV Lib "user32.dll" _
       Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As TVITEM) As Integer

    Public Declare Function PostMessage Lib "user32.dll" _
        Alias "PostMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer

    Public Delegate Function EnumWindowsCallback(ByVal hWnd As Integer, ByVal lParam As Integer) As Boolean

    Public Declare Function EnumWindows Lib "user32.dll" _
        Alias "EnumWindows" (ByVal callback As EnumWindowsCallback, ByVal lParam As Integer) As Integer

    Public Declare Function EnumChildWindows Lib "user32.dll" _
        Alias "EnumChildWindows" (ByVal hWndParent As Integer, ByVal callback As EnumWindowsCallback, ByVal lParam As Integer) As Integer

    Public Declare Function GetMenu Lib "user32.dll" Alias "GetMenu" (ByVal hwnd As Integer) As Integer

    Public Declare Function GetMenuItemID Lib "user32.dll" Alias "GetMenuItemID" (ByVal hMenu As Integer, ByVal nPos As Integer) As Integer

    Public Declare Function GetSubMenu Lib "user32" Alias "GetSubMenu" (ByVal hMenu As Integer, ByVal nPos As Integer) As Integer

    Public Declare Function BringWindowToTop Lib "user32" Alias "BringWindowToTop" (ByVal hwnd As Integer) As Integer

End Module
