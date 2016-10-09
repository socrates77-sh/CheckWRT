Imports System.IO
Imports System.Text

Module General

    'ȫ�ֳ��� 
    Public Const LOG_FILE_NAME As String = "result.log"             'LOG�ļ���
    Public Const CHKLIST_FILE_NAME As String = "checklist.txt"      'Checklist�ļ���
    Public Const STR_BUFFER_LEN As Integer = 256                    'StringBuffer��Ĭ�ϳ���
    Public Const EZPRO_TITLE As String = "EZPro100"                 'EZPro100���������ڱ���
    'Public Const ABOUT_TITLE As String = "About Writer"             'About���ڱ���
    Public Const TOOL_GROUP_TEXT As String = "��ݹ�����"           '��ݹ���Group����
    Public Const CONFIG_BUTTON_TEXT As String = "����оƬ"          '����оƬButton����
    Public Const CHIP_BUTTON_TEXT As String = "ѡ��оƬ"           'ѡ��оƬButton����
    Public Const INFO_GROUP_TEXT As String = "оƬ��Ϣ"             'оƬ��ϢGroup����

    'ȫ�ֱ���
    Public swLog As StreamWriter    'LOG�ļ�sw
    Public hWndMain As Integer      '�����ھ��
    Public sInfoText As String             'оƬ��Ϣ�����ı�

    'ͬʱ�����LOG�ļ���console
    Public Sub myWriteLine(ByVal str As String)
        swLog.WriteLine(str)
        Console.WriteLine(str)
    End Sub

    Public Sub myWrite(ByVal str As String)
        swLog.Write(str)
        Console.Write(str)
    End Sub

    '����16λUInt16�Ĳ���λ
    'value������ֵ������λ��
    'start����ʼλ���Ӹ�λ����λ��
    'n����nλ
    '����ֵ�����ú����ֵ
    Public Function SetBitN_16(ByVal value As UInt16, ByVal start As Integer, ByVal n As Integer) As UInt16
        If start > 16 Or start < 0 Then
            Throw New System.Exception("SetBitN_16 param (start) set error")
        End If

        'n���ܳ������������n<1 or start-n+1<0
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

    'ȡ��оƬ��Ϣ�������֣�����EnumChildWindows��EnumWindowsCallbackʵ��
    'ֱ������sInfoText����
    Public Sub GetInfoText()
        EnumChildWindows(hWndMain, New EnumWindowsCallback(AddressOf Do_FindInfoText), 0)
    End Sub

    Private Function Do_FindInfoText(ByVal hWnd As Integer, ByVal lParam As Integer) As Boolean
        Dim sbClassName As New StringBuilder(STR_BUFFER_LEN)
        GetClassName(hWnd, sbClassName, STR_BUFFER_LEN)
        If sbClassName.ToString = "TMemo" Then             'Ψһһ��TMemo�ؼ�����оƬ��Ϣ��
            Dim sb As New StringBuilder(STR_BUFFER_LEN)
            SendMessageSB(hWnd, WM_GETTEXT, STR_BUFFER_LEN, sb)
            sInfoText = sb.ToString
            Return False
        Else
            Return True
        End If
    End Function

    '����˵���
    'hWnd�˵����ڴ��ھ��
    'nSubMenuPos �Ӳ˵���ţ�0��ʼ
    'nMenuItemPos �˵�����ţ�0��ʼ
    Public Function ClickMenu(ByVal hWnd As Integer, ByVal nSubMenuPos As Integer, ByVal nMenuItemPos As Integer) As Integer
        Dim hMenu As Integer = GetMenu(hWnd)
        Dim hSubMenu As Integer = GetSubMenu(hMenu, nSubMenuPos)
        Dim hMenuItem As Integer = GetMenuItemID(hSubMenu, nMenuItemPos)
        BringWindowToTop(hWnd)
        PostMessage(hWnd, WM_COMMAND, hMenuItem, 0)
    End Function

    '����˵���ʽ��Open�Ի��򣬲���ָ����WRT�ļ�
    'sFileName WRT�ļ�������·��
    '���ļ������ڣ��򷵻�False�����򷵻�Ture
    Public Function OpenWRTFile(ByVal sFileName As String) As Boolean
        If Not My.Computer.FileSystem.FileExists(sFileName) Then
            myWriteLine("ERROR: File " & sFileName & " not found!")
            Return False
        Else
            ClickMenu(hWndMain, 0, 0)       '�����Open���˵�
            System.Threading.Thread.Sleep(300)
            Dim hDlgOpen As Integer = FindWindow(vbNullString, "��")
            Dim hBtnOpen As Integer = FindWindowEx(hDlgOpen, 0, vbNullString, "��(&O)")
            Dim hTxtFileName As Integer = GetDlgItem(hDlgOpen, &H47C)   '��ID��ʽȡ��filename�ؼ���Edit���ľ��
            SendMessageS(hTxtFileName, WM_SETTEXT, 0, sFileName)   '��дfilename��ֵ
            SendMessage(hBtnOpen, BM_CLICK, 0, 0)   '���Open��ť
            myWriteLine("Open ... " & sFileName)
            Return True
        End If
    End Function

    '����˵���ʽ��Save�Ի��򣬲�����Ϊָ����WRT�ļ�
    'sFileName WRT�ļ�������·��
    '��·�������ڣ��򷵻�False�����򷵻�Ture
    Public Function SaveWRTFile(ByVal sFileName As String) As Boolean
        Dim sPath As String = My.Computer.FileSystem.GetParentPath(sFileName)
        If Not My.Computer.FileSystem.DirectoryExists(sPath) Then
            myWriteLine("ERROR: Path " & sPath & " not found!")
            Return False
        Else
            '�ļ���������ɾ������֤����
            If My.Computer.FileSystem.FileExists(sFileName) Then
                My.Computer.FileSystem.DeleteFile(sFileName)
            End If
            ClickMenu(hWndMain, 0, 1)       '�����Save as���˵�
            System.Threading.Thread.Sleep(300)
            Dim hDlgSave As Integer = FindWindow(vbNullString, "���Ϊ")
            Dim hBtnSave As Integer = FindWindowEx(hDlgSave, 0, vbNullString, "����(&S)")
            Dim hTxtFileName As Integer = GetDlgItem(hDlgSave, &H47C)   '��ID��ʽȡ��filename�ؼ���Edit���ľ��
            SendMessageS(hTxtFileName, WM_SETTEXT, 0, sFileName)   '��дfilename��ֵ
            SendMessage(hBtnSave, BM_CLICK, 0, 0)   '���Save��ť
            myWriteLine("Save ... " & sFileName)
            Return True
        End If
    End Function
End Module
