Imports System.IO
Imports System.Text

Module General

    '全局常量 
    Public Const LOG_FILE_NAME As String = "result.log"             'LOG文件名
    Public Const CHKLIST_FILE_NAME As String = "checklist.txt"      'Checklist文件名
    Public Const STR_BUFFER_LEN As Integer = 256                    'StringBuffer类默认长度
    Public Const EZPRO_TITLE As String = "EZPro100"                 'EZPro100程序主窗口标题
    'Public Const ABOUT_TITLE As String = "About Writer"             'About窗口标题
    Public Const TOOL_GROUP_TEXT As String = "快捷工具栏"           '快捷工具Group标题
    Public Const CONFIG_BUTTON_TEXT As String = "配置芯片"          '配置芯片Button标题
    Public Const CHIP_BUTTON_TEXT As String = "选择芯片"           '选择芯片Button标题
    Public Const INFO_GROUP_TEXT As String = "芯片信息"             '芯片信息Group标题

    '全局变量
    Public swLog As StreamWriter    'LOG文件sw
    Public hWndMain As Integer      '主窗口句柄
    Public sInfoText As String             '芯片信息栏的文本

    '同时输出到LOG文件和console
    Public Sub myWriteLine(ByVal str As String)
        swLog.WriteLine(str)
        Console.WriteLine(str)
    End Sub

    Public Sub myWrite(ByVal str As String)
        swLog.Write(str)
        Console.Write(str)
    End Sub

    '设置16位UInt16的部分位
    'value：设置值（部分位）
    'start：起始位（从高位到低位）
    'n：共n位
    '返回值：设置后的新值
    Public Function SetBitN_16(ByVal value As UInt16, ByVal start As Integer, ByVal n As Integer) As UInt16
        If start > 16 Or start < 0 Then
            Throw New System.Exception("SetBitN_16 param (start) set error")
        End If

        'n不能出现以下情况：n<1 or start-n+1<0
        If n < 1 Or n > start + 1 Then
            Throw New System.Exception("SetBitN_16 param (n) set error")
        End If

        Dim temp As UInt16
        Dim mask As UInt16
        temp = value
        temp <<= start - n + 1
        mask = &HFFFF
        mask <<= n
        mask = Not mask
        mask <<= start - n + 1
        mask = Not mask
        Return (temp Or mask)

    End Function

    '取得芯片信息栏的文字，利用EnumChildWindows及EnumWindowsCallback实现
    '直接设置sInfoText变量
    Public Sub GetInfoText()
        EnumChildWindows(hWndMain, New EnumWindowsCallback(AddressOf Do_FindInfoText), 0)
    End Sub

    Private Function Do_FindInfoText(ByVal hWnd As Integer, ByVal lParam As Integer) As Boolean
        Dim sbClassName As New StringBuilder(STR_BUFFER_LEN)
        GetClassName(hWnd, sbClassName, STR_BUFFER_LEN)
        If sbClassName.ToString = "TMemo" Then             '唯一一个TMemo控件即是芯片信息栏
            Dim sb As New StringBuilder(STR_BUFFER_LEN)
            SendMessageSB(hWnd, WM_GETTEXT, STR_BUFFER_LEN, sb)
            sInfoText = sb.ToString
            Return False
        Else
            Return True
        End If
    End Function

    '点击菜单项
    'hWnd菜单所在窗口句柄
    'nSubMenuPos 子菜单序号，0起始
    'nMenuItemPos 菜单项序号，0起始
    Public Function ClickMenu(ByVal hWnd As Integer, ByVal nSubMenuPos As Integer, ByVal nMenuItemPos As Integer) As Integer
        Dim hMenu As Integer = GetMenu(hWnd)
        Dim hSubMenu As Integer = GetSubMenu(hMenu, nSubMenuPos)
        Dim hMenuItem As Integer = GetMenuItemID(hSubMenu, nMenuItemPos)
        BringWindowToTop(hWnd)
        PostMessage(hWnd, WM_COMMAND, hMenuItem, 0)
    End Function

    '点击菜单方式打开Open对话框，并打开指定的WRT文件
    'sFileName WRT文件名，含路径
    '如文件不存在，则返回False；否则返回Ture
    Public Function OpenWRTFile(ByVal sFileName As String) As Boolean
        If Not My.Computer.FileSystem.FileExists(sFileName) Then
            myWriteLine("ERROR: File " & sFileName & " not found!")
            Return False
        Else
            ClickMenu(hWndMain, 0, 0)       '点击“Open”菜单
            System.Threading.Thread.Sleep(300)
            Dim hDlgOpen As Integer = FindWindow(vbNullString, "打开")
            Dim hBtnOpen As Integer = FindWindowEx(hDlgOpen, 0, vbNullString, "打开(&O)")
            Dim hTxtFileName As Integer = GetDlgItem(hDlgOpen, &H47C)   '以ID方式取得filename控件（Edit）的句柄
            SendMessageS(hTxtFileName, WM_SETTEXT, 0, sFileName)   '填写filename的值
            SendMessage(hBtnOpen, BM_CLICK, 0, 0)   '点击Open按钮
            myWriteLine("Open ... " & sFileName)
            Return True
        End If
    End Function

    '点击菜单方式打开Save对话框，并保存为指定的WRT文件
    'sFileName WRT文件名，含路径
    '如路径不存在，则返回False；否则返回Ture
    Public Function SaveWRTFile(ByVal sFileName As String) As Boolean
        Dim sPath As String = My.Computer.FileSystem.GetParentPath(sFileName)
        If Not My.Computer.FileSystem.DirectoryExists(sPath) Then
            myWriteLine("ERROR: Path " & sPath & " not found!")
            Return False
        Else
            '文件存在则先删除，保证更新
            If My.Computer.FileSystem.FileExists(sFileName) Then
                My.Computer.FileSystem.DeleteFile(sFileName)
            End If
            ClickMenu(hWndMain, 0, 1)       '点击“Save as”菜单
            System.Threading.Thread.Sleep(300)
            Dim hDlgSave As Integer = FindWindow(vbNullString, "另存为")
            Dim hBtnSave As Integer = FindWindowEx(hDlgSave, 0, vbNullString, "保存(&S)")
            Dim hTxtFileName As Integer = GetDlgItem(hDlgSave, &H47C)   '以ID方式取得filename控件（Edit）的句柄
            SendMessageS(hTxtFileName, WM_SETTEXT, 0, sFileName)   '填写filename的值
            SendMessage(hBtnSave, BM_CLICK, 0, 0)   '点击Save按钮
            myWriteLine("Save ... " & sFileName)
            Return True
        End If
    End Function
End Module
