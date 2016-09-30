'Edition history
'0.1 新建 2016/9/29

Imports System.Text
Imports System.IO

Module myMain
    Const sVersion As String = "0.1"    '本软件版本号

    Dim sInfoText As String             '芯片信息栏的文本
    Dim alChipList As ArrayList         '所有芯片型号
    Dim sEzproVersionFromFile As String         'EZPRO软件版本号，来自checklist文件
    Dim sEzproVersionFromAbout As String         'EZPRO软件版本号，来自About窗口

    Sub Main()
        '初始化
        InitMain()

        'forMC30P6060.CheckAll()

        GetVersionFromAbout()
        myWriteLine(sEzproVersionFromAbout)
        'myWriteLine(hWndMain.ToString)
        'ClickMenu(hWndMain, 4, 0)
        'System.Threading.Thread.Sleep(1000)
        'Dim hwnd As Integer = FindWindow(vbNullString, EZPRO_TITLE)
        'myWriteLine(hwnd.ToString)
        'PostMessage(hwnd, WM_KEYDOWN, VK_RETURN, 0)

        'ExtractWindow()

        ''myWriteLine(GetChipTypeHandle().ToString)

        'myWriteLine(GetInfoText())
        ''ExtractWindow()
        'GetWindowInfoHandle()
        'ClickMenu(hWndMain, 4, 0)
        'System.Threading.Thread.Sleep(1000)
        'Dim hwnd As Integer = FindWindow(hWndMain, EZPRO_TITLE)
        'PostMessage(hWndMain, WM_KEYDOWN, VK_RETURN, 0)

        swLog.Close()
    End Sub

    '获取配置选择芯片窗口的句柄
    Private Function GetChipTypeHandle() As Integer
        Try
            '逐级查找配置芯片Button的handle
            Dim hGroupTool As Integer = FindWindowEx(hWndMain, 0, vbNullString, TOOL_GROUP_TEXT)
            Dim hCmdConfig As Integer = FindWindowEx(hGroupTool, 0, vbNullString, CHIP_BUTTON_TEXT)

            '点击选择芯片Button，打开窗口
            If Not hCmdConfig = 0 Then
                PostMessage(hCmdConfig, BM_CLICK, 0, 0)
                System.Threading.Thread.Sleep(300)          '等待300ms待窗口打开，否则可能取不到句柄
                Dim hwndChip As Integer = FindWindow(vbNullString, CHIP_BUTTON_TEXT)
                Dim hTreeChipType As Integer = FindWindowEx(hGroupTool, 0, "TTreeView", vbNullString)
                Return hTreeChipType
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
            Return 0
        End Try
    End Function

    '取得芯片信息栏的文字，
    '直接设置sEzproVersionFromAbout变量
    Private Sub GetVersionFromAbout()
        ClickMenu(hWndMain, 4, 0)
        System.Threading.Thread.Sleep(3000)
        '如果弹出无法找到串口的窗口，则关闭之
        Dim hwndMessage As Integer = FindWindow(vbNullString, EZPRO_TITLE)
        If hwndMessage Then
            PostMessage(hwndMessage, WM_KEYDOWN, VK_RETURN, 0)
            System.Threading.Thread.Sleep(5000)     '等待About窗口显示的时间，特别长！
        End If
        Dim hwndAbout = FindWindow(vbNullString, ABOUT_TITLE)
        '利用EnumChildWindows及EnumWindowsCallback实现
        EnumChildWindows(hwndAbout, New EnumWindowsCallback(AddressOf Do_FindVersionText), 0)
        '关闭About窗口
        System.Threading.Thread.Sleep(1000)
        PostMessage(hwndAbout, WM_KEYDOWN, VK_RETURN, 0)
    End Sub

    Private Function Do_FindVersionText(ByVal hWnd As Integer, ByVal lParam As Integer) As Boolean
        Dim strControlText As New StringBuilder(STR_BUFFER_LEN)
        'Dim strClassName As New StringBuilder(STR_BUFFER_LEN)
        GetWindowText(hWnd, strControlText, STR_BUFFER_LEN)
        'GetClassName(hWnd, strClassName, STR_BUFFER_LEN)

        'myWriteLine("(Class)" & strClassName.ToString.PadRight(15) & "(Text)" & strControlText.ToString)
        If strControlText.ToString.Contains("Version") Then             '查找到含Version那一个TEXT
            sEzproVersionFromAbout = strControlText.ToString
            Return False
        Else
            Return True
        End If
    End Function


    '取得芯片信息栏的文字，利用EnumChildWindows及EnumWindowsCallback实现
    '直接设置sInfoText变量
    Private Sub GetInfoText()
        EnumChildWindows(hWndMain, New EnumWindowsCallback(AddressOf Do_FindInfoText), 0)
    End Sub

    Private Function Do_FindInfoText(ByVal hWnd As Integer, ByVal lParam As Integer) As Boolean
        Dim strClassName As New StringBuilder(STR_BUFFER_LEN)
        GetClassName(hWnd, strClassName, STR_BUFFER_LEN)
        If strClassName.ToString = "TMemo" Then             '唯一一个TMemo控件即是芯片信息栏
            Dim str As New StringBuilder(STR_BUFFER_LEN)
            SendMessageS(hWnd, WM_GETTEXT, STR_BUFFER_LEN, str)
            sInfoText = str.ToString
            Return False
        Else
            Return True
        End If
    End Function

    Private Sub InitMain()
        '判断EZPro100是否运行，且只有一个进程
        Dim Processes As Process() = Process.GetProcessesByName(EZPRO_TITLE)
        If Processes.Length = 0 Then
            Console.WriteLine("ERROR: {0} is not running!", EZPRO_TITLE)
            End
        ElseIf Processes.Length > 1 Then
            Console.WriteLine("ERROR: More than one {0} are running!", EZPRO_TITLE)
            End
        Else
            Console.WriteLine("Found {0} running!" & vbCrLf, EZPRO_TITLE)
        End If
        hWndMain = FindWindow(vbNullString, EZPRO_TITLE)

        '打开LOG文件
        Try
            swLog = New StreamWriter(LOG_FILE_NAME)
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try

        '打开checklist文件，读取信息
        Try
            Dim sr As StreamReader = New StreamReader(CHKLIST_FILE_NAME)
            '从checklist读EZPRO的版本
            sEzproVersionFromFile = sr.ReadLine

            '从checklist读所有芯片型号
            alChipList = New ArrayList
            Dim str As String = sr.ReadLine
            Do
                If str = "*END" Then Exit Do
                alChipList.Add(str)
                str = sr.ReadLine
            Loop
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try

        'result文件头
        myWriteLine("CheckWRT(v" & sVersion & ")   Start: " & Now.ToString)
        myWriteLine("==================================================")
        myWriteLine("")
    End Sub

    '点击菜单项
    'hWnd菜单所在窗口句柄
    'nSubMenuPos子菜单序号，0起始
    'nMenuItemPos菜单项序号，0起始
    Public Function ClickMenu(ByVal hWnd As Integer, ByVal nSubMenuPos As Integer, ByVal nMenuItemPos As Integer) As Integer
        Dim hMenu As Integer = GetMenu(hWnd)
        Dim hSubMenu As Integer = GetSubMenu(hMenu, nSubMenuPos)
        Dim hMenuItem As Integer = GetMenuItemID(hSubMenu, nMenuItemPos)
        BringWindowToTop(hWnd)
        PostMessage(hWnd, WM_COMMAND, hMenuItem, 0)
    End Function

    Private Function ExtractWindow() As Boolean
        Dim hWndMain As Integer = FindWindow(vbNullString, EZPRO_TITLE)
        'Dim hWndMain As Integer = GetWindowInfoHandle()
        myWriteLine(hWndMain)
        '遍历窗口的所有控件
        If Not hWndMain = 0 Then
            EnumChildWindows(hWndMain, New EnumWindowsCallback(AddressOf DoControl1), 0)
        End If
        Return True
    End Function

    Private Function DoControl1(ByVal hWnd As Integer, ByVal lParam As Integer) As Boolean
        Dim strControlText As New StringBuilder(STR_BUFFER_LEN)
        Dim strClassName As New StringBuilder(STR_BUFFER_LEN)

        GetWindowText(hWnd, strControlText, STR_BUFFER_LEN)
        GetClassName(hWnd, strClassName, STR_BUFFER_LEN)

        'myWrite("[" & nIndexControl.ToString("00") & "]" & "(Class)" & strClassName.ToString.PadRight(15) _
        '    & "(Text)" & strControlText.ToString.PadLeft(50))

        myWriteLine("(Class)" & strClassName.ToString.PadRight(15) & "(Text)" & strControlText.ToString)
        If strClassName.ToString = "TMemo" Then
            myWriteLine("found")
            Dim strSelect As New StringBuilder(STR_BUFFER_LEN)
            SendMessageS(hWnd, WM_GETTEXT, 256, strSelect)
            myWriteLine(strSelect.ToString)
            Return False
        End If

        Return True
    End Function

    'Private Function GetWindowInfoHandle() As Integer
    '    Dim hWndMain As Integer
    '    Dim hGroupTool As Integer
    '    'Dim hCmdConfig As Integer

    '    Try
    '        '逐级查找配置芯片信息栏的handle
    '        hWndMain = FindWindow(vbNullString, EZPRO_TITLE)
    '        hGroupTool = FindWindowEx(hWndMain, 0, vbNullString, INFO_GROUP_TEXT)
    '        'hCmdConfig = FindWindowEx(hGroupTool, 0, vbNullString, CONFIG_BUTTON_TEXT)
    '        'myWriteLine(hGroupTool)

    '        '点击配置芯片Button，打开Option窗口
    '        'If Not hCmdConfig = 0 Then
    '        '    PostMessage(hCmdConfig, BM_CLICK, 0, 0)
    '        '    System.Threading.Thread.Sleep(300)
    '        '    Return FindWindow(vbNullString, Option_Win_Title)
    '        'Else
    '        '    Return 0
    '        'End If
    '        Return hGroupTool

    '    Catch ex As Exception
    '        MsgBox(ex.Message.ToString)
    '        Return 0
    '    End Try
    'End Function
End Module
