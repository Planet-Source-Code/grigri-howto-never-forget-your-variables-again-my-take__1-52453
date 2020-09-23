Attribute VB_Name = "CaptainHook"
Option Explicit

Private Const WH_CBT As Long = 5
Private Const WH_KEYBOARD As Long = 2

Private Enum HCBT_ACTIONCODES
    HCBT_MOVESIZE = 0
    HCBT_MINMAX = 1
    HCBT_QS = 2
    HCBT_CREATEWND = 3
    HCBT_DESTROYWND = 4
    HCBT_ACTIVATE = 5
    HCBT_CLICKSKIPPED = 6
    HCBT_KEYSKIPPED = 7
    HCBT_SYSCOMMAND = 8
    HCBT_SETFOCUS = 9
End Enum
Private Const HC_ACTION As Long = 0
Private Const KF_UP As Long = &H8000

Private Declare Sub OutputDebugString Lib "kernel32.dll" Alias "OutputDebugStringA" (ByVal lpOutputString As String)

Private Declare Function CallNextHookEx Lib "user32.dll" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowsHookEx Lib "user32.dll" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32.dll" (ByVal hHook As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer
Private Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hwnd As Long, lpClassName As Any, ByVal nMaxCount As Long) As Long
Private Declare Function lstrcmp Lib "kernel32.dll" Alias "lstrcmpA" (lpString1 As Any, lpString2 As Any) As Long

Private vbInst As VBE
Private szVbaWindowClass() As Byte
Private hHookCBT As Long
Private hHookKybd As Long
Private hWndCodePane As Long
Private bMonitoring As Boolean

Public Property Set VBInstance(NewVBInst As VBE)
    Set vbInst = NewVBInst
    
    ' Convert the string from unicode to ansi
    ' for faster comparaison later on
    szVbaWindowClass() = StrConv("VbaWindow" & vbNullChar, vbFromUnicode)
End Property

Public Function SetHook() As Boolean
    If InIDE Then
        MsgBox "This add-in must be COMPILED to run" & vbCrLf & "It can't be run from the IDE", vbCritical
    End If
    If hHookCBT <> 0 Or hHookKybd <> 0 Then Exit Function
    
    hHookCBT = SetWindowsHookEx(WH_CBT, AddressOf CBTHookProc, 0, App.ThreadID)
    hHookKybd = SetWindowsHookEx(WH_KEYBOARD, AddressOf KybdHookProc, 0, App.ThreadID)
    
    If hHookCBT <> 0 And hHookKybd <> 0 Then SetHook = True
    
    If Not vbInst Is Nothing Then
        If Not vbInst.ActiveWindow Is Nothing Then
            If vbInst.ActiveWindow.Type = vbext_wt_CodeWindow Then
                bMonitoring = True
            End If
        End If
    End If
End Function

Public Function StopHook() As Boolean
    Dim lReturnCode As Long
    
    If hHookCBT = 0 And hHookKybd = 0 Then Exit Function
    
    lReturnCode = UnhookWindowsHookEx(hHookCBT)
    lReturnCode = UnhookWindowsHookEx(hHookKybd)
    
    hHookCBT = 0
    hHookKybd = 0
    
    If lReturnCode <> 0 Then
        StopHook = True
    End If
End Function

Public Function IsHooked() As Boolean
    If hHookCBT = 0 Or hHookKybd = 0 Then Exit Function
    IsHooked = True
End Function

Private Function CBTHookProc(ByVal eCode As HCBT_ACTIONCODES, ByVal wParam As Long, ByVal lParam As Long) As Long
    Static szBuf(260) As Byte
    Select Case eCode
    Case HCBT_SETFOCUS
        If vbInst Is Nothing Then GoTo EXIT_STAGE_LEFT
        ' Get the class of the window being activated
        GetClassName wParam, szBuf(0), 260
        ' Check if it is "VbaWindow"
        If lstrcmp(szBuf(0), szVbaWindowClass(0)) = 0 Then
            bMonitoring = True
            hWndCodePane = wParam
            ODS "VbaWindow activated - monitoring"
        Else
            bMonitoring = False
            hWndCodePane = 0
            ODS "non-VbaWindow activated - not monitoring"
        End If
    End Select
    
EXIT_STAGE_LEFT:
    CBTHookProc = CallNextHookEx(hHookCBT, eCode, wParam, lParam)
End Function

Private Function KybdHookProc(ByVal iCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If bMonitoring Then
        ' Key message?
        If iCode = HC_ACTION Then
            ' M key?
            If wParam = vbKeyM Then
                ' Being pressed?
                If lParam And KF_UP Then
                    ' Is control key down?
                    If GetAsyncKeyState(vbKeyControl) And KF_UP Then
                        ' Get to work!
                        DoPopVarList
                        ' Return a non-zero value to keep
                        ' the window from getting the message
                        KybdHookProc = 1
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    
    KybdHookProc = CallNextHookEx(hHookKybd, iCode, wParam, lParam)
End Function

Private Sub DoPopVarList()
    ODS "DoPopVarList"
    
    Dim cp As CodePane, cm As CodeModule
    Dim Line1 As Long, Line2 As Long, Col1 As Long, Col2 As Long
    Dim sLine As String
    Dim sVarName As String
    
    ' Can't be too careful...
    If vbInst Is Nothing Then Exit Sub
    Set cp = vbInst.ActiveCodePane
    If cp Is Nothing Then Exit Sub
    
    ' Get the code module
    Set cm = cp.CodeModule
    
    ' Retrieve the current selection
    cp.GetSelection Line1, Col1, Line2, Col2
    ' Make sure nothing is selected
    If Line1 <> Line2 Then Exit Sub
    If Col1 <> Col2 Then Exit Sub
    ' Extract the line
    sLine = cm.Lines(Line1, 1)
    If Right$(sLine, 2) = vbCrLf Then sLine = Left$(sLine, Len(sLine) - 2)
    
    ' Show the list and return the selection
    sVarName = VarList.ShowVariables(cm, hWndCodePane)
    If sVarName <> "" Then
        ' Insert the variable name into the code
        sLine = Left$(sLine, Col1) & sVarName & Mid$(sLine, Col1 + 1)
        cm.ReplaceLine Line1, sLine
        ' Update the caret position
        Col1 = Col1 + Len(sVarName)
        Col2 = Col1
        cp.SetSelection Line1, Col1, Line2, Col2
    End If
End Sub

Public Sub ODS(sOut As String)
    ' Very, very useful when you can only
    ' work in compiled mode
    ' It's like a Debug.Print
    ' (you need the "DbMon" application to see the output)
    OutputDebugString sOut & vbCrLf
End Sub

Private Function InIDE() As Boolean
    On Error Resume Next
    Debug.Print 1 / 0
    If Err Then InIDE = True
End Function
