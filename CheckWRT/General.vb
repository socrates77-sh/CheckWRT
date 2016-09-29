Imports System.IO

Module General

    'ȫ�ֳ��� 
    Public Const LOG_FILE_NAME As String = "result.log"             'LOG�ļ���
    Public Const CHKLIST_FILE_NAME As String = "checklist.txt"      'Checklist�ļ���
    Public Const STR_BUFFER_LEN As Integer = 256                    'StringBuffer��Ĭ�ϳ���
    Public Const EZPRO_TITLE As String = "EZPro100"                 'EZPro100���������ڱ���
    Public Const ABOUT_TITLE As String = "About Writer"             'About���ڱ���
    Public Const TOOL_GROUP_TEXT As String = "��ݹ�����"           '��ݹ���Group����
    Public Const CONFIG_BUTTON_TEXT As String = "����оƬ"          '����оƬButton����
    Public Const CHIP_BUTTON_TEXT As String = "ѡ��оƬ"           'ѡ��оƬButton����
    Public Const INFO_GROUP_TEXT As String = "оƬ��Ϣ"             'оƬ��ϢGroup����

    'ȫ�ֱ���
    Public swLog As StreamWriter    'LOG�ļ�sw
    Public hWndMain As Integer      '�����ھ��

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

End Module
