Imports System.Text
Imports System.IO

Module myMain

    Sub Main()

        '初始化
        Try
            swLog = New StreamWriter(LOG_FILE_NAME)
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try

        myWriteLine(Now.ToString & vbCrLf)

        forMC30P6060.CheckAll()


        'Dim hWndMain As Integer
        'Dim hWndChild As Integer
        'Dim hWndChild1 As Integer
        'Dim className As New StringBuilder(256)

        'hWndMain = FindWindow(vbNullString, "EZPro100")
        ''Win32API.GetClassName(hWndMain, className, 256)
        'hWndChild = FindWindowEx(hWndMain, 0, vbNullString, "快捷工具栏")
        'hWndChild1 = FindWindowEx(hWndChild, 0, vbNullString, "配置芯片")
        'GetClassName(hWndChild1, className, 256)
        'PostMessage(hWndChild1, BM_CLICK, 0, 0)
        'System.Threading.Thread.Sleep(100)
        'hWndMain = FindWindow(vbNullString, "配置 MC30P6060")
        ''hWndChild = FindWindowEx(hWndMain, 0, vbNullString, "加密")
        ''GetClassName(hWndChild, className, 256)

        ''EnumWindows(New EnumWindowsCallback(AddressOf FillActiveProcessList), 0)
        'EnumChildWindows(hWndMain, New EnumWindowsCallback(AddressOf FillActiveProcessList), 0)
        ''EnumChildWindows(hWndMain, New EnumWindowsCallback(AddressOf FillActiveProcessList), 0)
        ''EnumChildWindows(hWndMain, New EnumWindowsCallback(AddressOf FillActiveProcessList), 0)
        ''Console.WriteLine(hWndMain.ToString)
        ''Console.WriteLine(hWndChild.ToString)
        ''Console.WriteLine(hWndChild1.ToString)
        ''Console.WriteLine(className.ToString)
        'Console.ReadLine()
        swLog.Close()
    End Sub

    'Function FillActiveProcessList(ByVal hWnd As Integer, ByVal lParam As Integer) As Boolean
    '    Dim windowText As New StringBuilder(StringBufferLength)
    '    Dim className As New StringBuilder(StringBufferLength)

    '    GetWindowText(hWnd, windowText, StringBufferLength)
    '    GetClassName(hWnd, className, StringBufferLength)

    '    Console.WriteLine(windowText.ToString)
    '    Console.WriteLine(className.ToString)

    '    Return True
    'End Function

    'Private Function ExtractWindowOption(ByVal sProductType As String) As Boolean
    '    Select Case sProductType
    '        Case "MC30P6060"
    '        Case Else
    '            MsgBox("该型号暂不支持", MsgBoxStyle.OkOnly, "Warnning")
    '    End Select
    'End Function

End Module
