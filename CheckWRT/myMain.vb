'Edition history
'0.1 �½� 2016/9/29

Imports System.Text
Imports System.IO

Module myMain
    Const sVersion As String = "0.1"    '������汾��

    Dim sInfoText As String             'оƬ��Ϣ�����ı�
    Dim alChipList As ArrayList         '����оƬ�ͺ�
    Dim sEzproVersion As String         'EZPRO����汾��

    Sub Main()
        '��ʼ��
        InitMain()

        forMC30P6060.CheckAll()
        'myWriteLine(GetInfoText())
        ''ExtractWindow()
        ''GetWindowInfoHandle()
        'ClickMenu(hWndMain, 4, 0)
        'System.Threading.Thread.Sleep(1000)
        'Dim hwnd As Integer = FindWindow(hWndMain, EZPRO_TITLE)
        'PostMessage(hWndMain, WM_KEYDOWN, VK_RETURN, 0)


        swLog.Close()
    End Sub

    'ȡ��оƬ��Ϣ�������֣�����EnumChildWindows��EnumWindowsCallbackʵ��
    Private Function GetInfoText() As String
        EnumChildWindows(hWndMain, New EnumWindowsCallback(AddressOf Do_FindInfoText), 0)
        Return sInfoText
    End Function

    Private Function Do_FindInfoText(ByVal hWnd As Integer, ByVal lParam As Integer) As Boolean
        Dim strClassName As New StringBuilder(STR_BUFFER_LEN)
        GetClassName(hWnd, strClassName, STR_BUFFER_LEN)
        If strClassName.ToString = "TMemo" Then             'Ψһһ��TMemo�ؼ�����оƬ��Ϣ��
            Dim str As New StringBuilder(STR_BUFFER_LEN)
            SendMessageS(hWnd, WM_GETTEXT, STR_BUFFER_LEN, str)
            sInfoText = str.ToString
            Return False
        Else
            Return True
        End If
    End Function

    Private Sub InitMain()
        '�ж�EZPro100�Ƿ����У���ֻ��һ������
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

        '��LOG�ļ�
        Try
            swLog = New StreamWriter(LOG_FILE_NAME)
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try

        '��checklist�ļ�����ȡ��Ϣ
        Try
            Dim sr As StreamReader = New StreamReader(CHKLIST_FILE_NAME)
            '��checklist��EZPRO�İ汾
            sEzproVersion = sr.ReadLine

            '��checklist������оƬ�ͺ�
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

        'result�ļ�ͷ
        myWriteLine("CheckWRT(v" & sVersion & ")   Start: " & Now.ToString)
        myWriteLine("==================================================")
        myWriteLine("")
    End Sub

    '����˵���
    'hWnd�˵����ڴ��ھ��
    'nSubMenuPos�Ӳ˵���ţ�0��ʼ
    'nMenuItemPos�˵�����ţ�0��ʼ
    Public Function ClickMenu(ByVal hWnd As Integer, ByVal nSubMenuPos As Integer, ByVal nMenuItemPos As Integer) As Integer
        Dim hMenu As Integer
        Dim hSubMenu As Integer
        Dim hMenuItem As Integer

        hMenu = GetMenu(hWnd)
        hSubMenu = GetSubMenu(hMenu, nSubMenuPos)
        hMenuItem = GetMenuItemID(hSubMenu, nMenuItemPos)
        BringWindowToTop(hWnd)
        PostMessage(hWnd, WM_COMMAND, hMenuItem, 0)
    End Function

    'Private Function ExtractWindow() As Boolean
    '    Dim hWndMain As Integer = FindWindow(vbNullString, EZPRO_TITLE)
    '    'Dim hWndMain As Integer = GetWindowInfoHandle()
    '    myWriteLine(hWndMain)
    '    '�������ڵ����пؼ�
    '    If Not hWndMain = 0 Then
    '        EnumChildWindows(hWndMain, New EnumWindowsCallback(AddressOf DoControl1), 0)
    '    End If
    '    Return True
    'End Function

    'Private Function DoControl1(ByVal hWnd As Integer, ByVal lParam As Integer) As Boolean
    '    Dim strControlText As New StringBuilder(STR_BUFFER_LEN)
    '    Dim strClassName As New StringBuilder(STR_BUFFER_LEN)

    '    GetWindowText(hWnd, strControlText, STR_BUFFER_LEN)
    '    GetClassName(hWnd, strClassName, STR_BUFFER_LEN)

    '    'myWrite("[" & nIndexControl.ToString("00") & "]" & "(Class)" & strClassName.ToString.PadRight(15) _
    '    '    & "(Text)" & strControlText.ToString.PadLeft(50))

    '    myWriteLine("(Class)" & strClassName.ToString.PadRight(15) & "(Text)" & strControlText.ToString)
    '    If strClassName.ToString = "TMemo" Then
    '        myWriteLine("found")
    '        Dim strSelect As New StringBuilder(STR_BUFFER_LEN)
    '        SendMessageS(hWnd, WM_GETTEXT, 256, strSelect)
    '        myWriteLine(strSelect.ToString)
    '        Return False
    '    End If

    '    Return True
    'End Function

    'Private Function GetWindowInfoHandle() As Integer
    '    Dim hWndMain As Integer
    '    Dim hGroupTool As Integer
    '    'Dim hCmdConfig As Integer

    '    Try
    '        '�𼶲�������оƬ��Ϣ����handle
    '        hWndMain = FindWindow(vbNullString, EZPRO_TITLE)
    '        hGroupTool = FindWindowEx(hWndMain, 0, vbNullString, INFO_GROUP_TEXT)
    '        'hCmdConfig = FindWindowEx(hGroupTool, 0, vbNullString, CONFIG_BUTTON_TEXT)
    '        'myWriteLine(hGroupTool)

    '        '�������оƬButton����Option����
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
