Imports System.IO

Module General

    '全局常量 
    Public Const LOG_FILE_NAME As String = "result.log"             'LOG文件名
    Public Const CHKLIST_FILE_NAME As String = "checklist.txt"      'Checklist文件名
    Public Const STR_BUFFER_LEN As Integer = 256                    'StringBuffer类默认长度
    Public Const EZPRO_TITLE As String = "EZPro100"                 'EZPro100程序主窗口标题
    Public Const ABOUT_TITLE As String = "About Writer"             'About窗口标题
    Public Const TOOL_GROUP_TEXT As String = "快捷工具栏"           '快捷工具Group标题
    Public Const CONFIG_BUTTON_TEXT As String = "配置芯片"          '配置芯片Button标题
    Public Const CHIP_BUTTON_TEXT As String = "选择芯片"           '选择芯片Button标题
    Public Const INFO_GROUP_TEXT As String = "芯片信息"             '芯片信息Group标题

    '全局变量
    Public swLog As StreamWriter    'LOG文件sw
    Public hWndMain As Integer      '主窗口句柄

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

End Module
