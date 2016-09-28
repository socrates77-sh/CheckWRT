Imports System.Text

#Const FOR_DEBUG = 1

Module forMC30P6060
    Private Const Product_Type_Name = "MC30P6060"                   '品名
    Private Const Option_Win_Title As String = "配置 MC30P6060"     '配置Option的窗口标题
    Dim nIndexControl As Integer                                    '遍历控件的编号
    Dim OptionFromWin(7) As UInt16                                  '存储Option窗口获取的option值

    Public Function CheckAll() As Boolean
        Dim i As Integer

        For i = 0 To 7
            OptionFromWin(i) = &HFFFF
        Next

        myWriteLine("=======================================================================================")
        myWriteLine("  " & Product_Type_Name)
        myWriteLine("=======================================================================================")
        ExtractWindowOption()

        Return True
    End Function

    '提取Option窗口的信息
    Private Function ExtractWindowOption() As Boolean
        Dim hWnd As Integer
        hWnd = GetWindowOptionHandle()

        '遍历Option窗口的所有控件
        nIndexControl = 0
        If Not hWnd = 0 Then
            EnumChildWindows(hWnd, New EnumWindowsCallback(AddressOf DoControl), 0)
        End If

#If FOR_DEBUG Then
        Dim i As Integer
        'OptionFromWin(0) = OptionFromWin(0) And SetBitN_16(3, 14, 3)
        myWriteLine("******************")
        For i = 0 To 7
            myWriteLine("OPTION[" & i.ToString & "]=" & OptionFromWin(i).ToString("X4"))
        Next
#End If

        Return True
    End Function

    'EnumWindowsCallback函数，用于处理每个控件
    Private Function DoControl(ByVal hWnd As Integer, ByVal lParam As Integer) As Boolean
        Dim strControlText As New StringBuilder(STR_BUFFER_LEN)
        Dim strClassName As New StringBuilder(STR_BUFFER_LEN)

        GetWindowText(hWnd, strControlText, STR_BUFFER_LEN)
        GetClassName(hWnd, strClassName, STR_BUFFER_LEN)

        'myWrite("[" & nIndexControl.ToString("00") & "]" & "(Class)" & strClassName.ToString.PadRight(15) _
        '    & "(Text)" & strControlText.ToString.PadLeft(50))

        myWrite("[" & nIndexControl.ToString("00") & "]" & "(Class)" & strClassName.ToString.PadRight(15) _
            & "(Text)" & strControlText.ToString)

        Dim str As String
        str = strClassName.ToString
        Select Case str
            Case "TRadioButton"
                GetValue_TRadioButton(hWnd)
            Case "TComboBox"
                GetValue_TComboBox(hWnd)
            Case Else
                'myWrite(vbCrLf)
                myWriteLine("")
        End Select

        nIndexControl += 1

        Return True
    End Function

    '获取TRadioButton的值
    Private Sub GetValue_TRadioButton(ByVal hWnd As Integer)
        Dim nResult As Integer
        nResult = SendMessage(hWnd, BM_GETCHECK, 0, 0)
        myWriteLine("  (Value)" & nResult.ToString)

        '根据nIndexControl序号，分别处理对应option位
        Select Case nIndexControl
            Case 7
                If nResult = 1 Then
                    OptionFromWin(2) = OptionFromWin(2) And SetBitN_16(1, 13, 1)
                End If
            Case 8
                If nResult = 1 Then
                    OptionFromWin(2) = OptionFromWin(2) And SetBitN_16(0, 13, 1)
                End If
            Case 10
                If nResult = 1 Then
                    OptionFromWin(2) = OptionFromWin(2) And SetBitN_16(1, 5, 1)
                End If
            Case 11
                If nResult = 1 Then
                    OptionFromWin(2) = OptionFromWin(2) And SetBitN_16(0, 5, 1)
                End If
            Case 13
                If nResult = 1 Then
                    OptionFromWin(2) = OptionFromWin(2) And SetBitN_16(1, 4, 1)
                End If
            Case 14
                If nResult = 1 Then
                    OptionFromWin(2) = OptionFromWin(2) And SetBitN_16(0, 4, 1)
                End If
            Case 16
                If nResult = 1 Then
                    OptionFromWin(2) = OptionFromWin(2) And SetBitN_16(0, 3, 1)
                End If
            Case 17
                If nResult = 1 Then
                    OptionFromWin(2) = OptionFromWin(2) And SetBitN_16(1, 3, 1)
                End If
            Case 19
                If nResult = 1 Then
                    OptionFromWin(2) = OptionFromWin(2) And SetBitN_16(1, 2, 1)
                End If
            Case 20
                If nResult = 1 Then
                    OptionFromWin(2) = OptionFromWin(2) And SetBitN_16(0, 2, 1)
                End If
            Case 22
                If nResult = 1 Then
                    OptionFromWin(0) = OptionFromWin(0) And SetBitN_16(1, 13, 1)
                End If
            Case 23
                If nResult = 1 Then
                    OptionFromWin(0) = OptionFromWin(0) And SetBitN_16(0, 13, 1)
                End If
            Case 25
                If nResult = 1 Then
                    OptionFromWin(0) = OptionFromWin(0) And SetBitN_16(1, 8, 1)
                End If
            Case 26
                If nResult = 1 Then
                    OptionFromWin(0) = OptionFromWin(0) And SetBitN_16(0, 8, 1)
                End If
            Case 30
                If nResult = 1 Then
                    OptionFromWin(0) = OptionFromWin(0) And SetBitN_16(1, 4, 1)
                End If
            Case 31
                If nResult = 1 Then
                    OptionFromWin(0) = OptionFromWin(0) And SetBitN_16(0, 4, 1)
                End If
            Case 33
                If nResult = 1 Then
                    OptionFromWin(2) = OptionFromWin(2) And SetBitN_16(0, 1, 1)
                End If
            Case 34
                If nResult = 1 Then
                    OptionFromWin(2) = OptionFromWin(2) And SetBitN_16(1, 1, 1)
                End If
            Case 40
                If nResult = 1 Then
                    OptionFromWin(2) = OptionFromWin(2) And SetBitN_16(1, 7, 1)
                End If
            Case 41
                If nResult = 1 Then
                    OptionFromWin(2) = OptionFromWin(2) And SetBitN_16(0, 7, 1)
                End If
            Case 43
                If nResult = 1 Then
                    OptionFromWin(0) = OptionFromWin(0) And SetBitN_16(1, 12, 1)
                End If
            Case 46
                If nResult = 1 Then
                    OptionFromWin(0) = OptionFromWin(0) And SetBitN_16(0, 12, 1)
                End If
        End Select

    End Sub

    '获取TComboBox的值
    Private Sub GetValue_TComboBox(ByVal hWnd As Integer)
        'Dim nCount As Integer
        Dim nSelect As Integer
        'Dim nResult As Integer
        Dim strSelect As New StringBuilder(STR_BUFFER_LEN)

        'nCount = SendMessage(hWnd, CB_GETCOUNT, 0, 0)
        nSelect = SendMessage(hWnd, CB_GETCURSEL, 0, 0)

        'nSelect=-1，表示combobox未选择条目
        If Not nSelect = -1 Then
            SendMessageS(hWnd, CB_GETLBTEXT, nSelect, strSelect)
        Else
            strSelect.Append("<no select>")
        End If

        myWrite("  (Value)" & nSelect.ToString)
        myWriteLine("  " & strSelect.ToString)

        '根据nIndexControl序号，分别处理对应option位
        Select Case nIndexControl
            Case 0
                If nSelect = 1 Then
                    OptionFromWin(2) = OptionFromWin(2) And SetBitN_16(1, 13, 1)
                End If
        End Select
    End Sub

    '获取配置Option窗口的句柄
    Private Function GetWindowOptionHandle() As Integer
        Dim hWndMain As Integer
        Dim hGroupTool As Integer
        Dim hCmdConfig As Integer

        Try
            '逐级查找配置芯片Button的handle
            hWndMain = FindWindow(vbNullString, EZPRO_TITLE)
            hGroupTool = FindWindowEx(hWndMain, 0, vbNullString, TOOL_GROUP_TEXT)
            hCmdConfig = FindWindowEx(hGroupTool, 0, vbNullString, CONFIG_BUTTON_TEXT)

            '点击配置芯片Button，打开Option窗口
            If Not hCmdConfig = 0 Then
                PostMessage(hCmdConfig, BM_CLICK, 0, 0)
                System.Threading.Thread.Sleep(300)
                Return FindWindow(vbNullString, Option_Win_Title)
            Else
                Return 0
            End If

        Catch ex As Exception
            MsgBox(ex.Message.ToString)
            Return 0
        End Try
    End Function

End Module
