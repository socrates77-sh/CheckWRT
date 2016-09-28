Imports System.Text
Imports System.IO

Module myMain

    Sub Main()

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

        '初始化
        Try
            swLog = New StreamWriter(LOG_FILE_NAME)
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try

        myWriteLine(Now.ToString & vbCrLf)

        'forMC30P6060.CheckAll()

        ExtractWindow()
        'GetWindowInfoHandle()

        swLog.Close()
    End Sub

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
        End If

        Return True
    End Function

    Private Function GetWindowInfoHandle() As Integer
        Dim hWndMain As Integer
        Dim hGroupTool As Integer
        'Dim hCmdConfig As Integer

        Try
            '逐级查找配置芯片信息栏的handle
            hWndMain = FindWindow(vbNullString, EZPRO_TITLE)
            hGroupTool = FindWindowEx(hWndMain, 0, vbNullString, INFO_GROUP_TEXT)
            'hCmdConfig = FindWindowEx(hGroupTool, 0, vbNullString, CONFIG_BUTTON_TEXT)
            'myWriteLine(hGroupTool)

            '点击配置芯片Button，打开Option窗口
            'If Not hCmdConfig = 0 Then
            '    PostMessage(hCmdConfig, BM_CLICK, 0, 0)
            '    System.Threading.Thread.Sleep(300)
            '    Return FindWindow(vbNullString, Option_Win_Title)
            'Else
            '    Return 0
            'End If
            Return hGroupTool

        Catch ex As Exception
            MsgBox(ex.Message.ToString)
            Return 0
        End Try
    End Function
End Module
