Attribute VB_Name = "basProcess"
'/******************************************************************/
'/*�����G�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*��m��t�ε{�ǳB�z�������a��C                                  */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�ҲաG�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@    �@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�ѦҡG    �@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*����G    �@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�`�N�ƶ��G�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*�L�C                                                            */
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*                                           Edit By Edward Hsieh */
'/*                                      Last Edit Date 2009/05/13 */
'/******************************************************************/
Option Explicit

'/******************�B�z�{�ǥΪ�Win32API�`��******************/

'/*�t�Φ�C��*/
Public Const NIM_ADD  As Long = &H0
Public Const NIM_DELETE  As Long = &H2
Public Const NIM_MODIFY  As Long = &H1
Public Const NIF_ICON  As Long = &H2
Public Const NIF_MESSAGE  As Long = &H1
Public Const NIF_TIP  As Long = &H4
'/**/

'/*�t�ΩR�O��*/
Public Const WM_CLOSE  As Long = &H10
Public Const WM_CLEAR  As Long = &H303
Public Const WM_CHAR  As Long = &H102
'/**/


'/*��L�ƥ��*/
Public Const WM_KEYDOWN  As Long = &H100
Public Const WM_KEYUP  As Long = &H101
Public Const WH_KEYBOARD As Long = 2
'/**/

'/*�ƹ��ƥ��*/
Public Const WM_MOUSEMOVE  As Long = &H200
Public Const WM_LBUTTONDBLCLK  As Long = &H203
Public Const WM_LBUTTONDOWN  As Long = &H201
Public Const WM_LBUTTONUP  As Long = &H202
Public Const WM_MBUTTONDBLCLK  As Long = &H209
Public Const WM_MBUTTONDOWN  As Long = &H207
Public Const WM_MBUTTONUP  As Long = &H208
Public Const WM_RBUTTONDBLCLK  As Long = &H206
Public Const WM_RBUTTONDOWN  As Long = &H204
Public Const WM_RBUTTONUP  As Long = &H205
'/**/

'/*�ܴ����������A��*/
Public Const SW_HIDE As Long = 0
Public Const SW_SHOWNORMAL  As Long = 1
Public Const SW_NORMAL  As Long = 1
Public Const SW_SHOWMINIMIZED  As Long = 2
Public Const SW_SHOWMAXIMIZED  As Long = 3
Public Const SW_MAXIMIZE  As Long = 3
Public Const SW_SHOWNOACTIVATE  As Long = 4
Public Const SW_SHOW  As Long = 5
Public Const SW_MINIMIZE  As Long = 6
Public Const SW_SHOWMINNOACTIVE  As Long = 7
Public Const SW_SHOWNA  As Long = 8
Public Const SW_RESTORE  As Long = 9
Public Const SW_SHOWDEFAULT  As Long = 10
Public Const SW_FORCEMINIMIZE  As Long = 11
Public Const SW_MAX  As Long = 11
'/**/


'/*���ܵ����Ϊ���*/
Public Const LWA_COLORKEY As Long = &H1
Public Const LWA_ALPHA As Long = &H2
Public Const GWL_EXSTYLE As Long = -20
Public Const WS_EX_LAYERED As Long = &H80000
'/**/


'/*�n��{�ǤU�����O*/
Public Const PROCESS_QUERY_INFORMATION = &H400
'/**/


'/*�ˬd�{�ǥثe���A��*/
Public Const STILL_ALIVE = &H103
'/**/

'/***********************************************************/

'/******************�B�z�{�ǥΪ�Win32API���c******************/

'/*�t�ε{�Ǧb���w�α�����Ʈɷ|�Ψ쪺���c*/
Public Type EVENTMSG
        message As Long
        paramL As Long
        paramH As Long
        time As Long
        hWnd As Long
End Type
'/**/
'/*�t�Φ�C�|�Ψ쪺���c*/
Public Type NOTIFYICONDATA
        cbSize As Long
        hWnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type
'/**/

'/***********************************************************/

'/****************�B�z�{�ǥΪ�Win32API�禡*******************/
Public Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpExitCode As Long) As Long
Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long

Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Public Declare Function CallNextHookEx Lib "user32.dll" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Public Declare Function SetWindowsHookEx Lib "user32.dll" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32.dll" (ByVal hHook As Long) As Long
Public Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer

Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Public Declare Function Shell_NotifyIconA Lib "shell32.dll" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
'/**************************************************/




'/********************��t�ε{�ǳB�z�������`��*********************/
Public Const MAX_PROCESS As Long = 755739 '���@�~�t�C�̦h�i�֦����{�Ǽƶq
'/*********************�p�حק諸(2009/05/13)**********************/


'/********************��t�ε{�ǳB�z���������c*********************/
'/*�@�Өt�ε{�����ӷ|�������c*/
Public Type SYSTEM_PROCESS
    Handle As Long
    Class As String
    Name As String
End Type
'/*�p�حק諸(2009/05/13)*/
'/*********************�p�حק諸(2009/05/13)**********************/


'/********************��t�ε{�ǳB�z�������ܼ�*********************/

'/*�x�s�C�|�����Ψ쪺*/
Public All_Process(MAX_PROCESS) As SYSTEM_PROCESS '�Ψ��x�s�C�|�Ҧ��{�ǮɡA�Ȧs�Ϊ��ܼ�
Public All_Process_Flag As Long '�Ψ��x�s�C�|�Ҧ��{�ǮɡA�ϥΪ��`�@����J�X�ӿz�����
'/**/

'/*�Y�p�Φ^�_��t�ΦC�|�Ψ쪺*/
Public mlngID As Long '�Ψ��x�s�o���Y�p��t�ΦC�ɬO�ĴX��
Public mcolNID As New Collection '�Ψ��x�s�����Y�p��t�ΦC�ɪ�hWnd
'/*�p�حק諸(2009/05/13)*/

'/*�x�s���������ܼ�*/
Public Dialog_OutputName As String '�n�۰ʿ�iDialog���󪺤�r
Public DialogName As String '��������ܤ�������D�W��
Public DialogType As String '��������ܤ���������N�X
'/**/

'/*�x�s�ثe��쪺�t����L�N�X*/
Public hNowHook As Long
Public hKeyCode As Long
Public hHaveKeyEvent As Boolean
'/*�p�حק諸(2009/06/16)*/

'/*********************�p�حק諸(2009/05/13)**********************/



'/*********************�Q�ο�J�Ȱ����z�����A�h��X�Ҧ��ŦX���󪺵{�ǡA�B�ӵ{�ǥ��ݬODialog�{��*********************/
Public Function Search_All_Dialog_Process(Optional ByVal Handle As Long, Optional ByVal Class As String, Optional ByVal Name As String) As SYSTEM_PROCESS()
    Dim i As Long
    Dim j As Long
    
    Dim hWnd1 As Long
    Dim hWnd2 As Long
    
    Dim Temp_Process() As SYSTEM_PROCESS
    Dim Mark_Process() As Boolean
    
    Dim Dialog_Count As Long
    Dim DIALOG_PROCESS() As SYSTEM_PROCESS
    
    
    
    Temp_Process = Search_All_Windows_By_Conditions(Handle, Class, Name)



    ReDim Mark_Process(UBound(Temp_Process))
    For i = 0 To UBound(Temp_Process) - 1
        hWnd1 = FindWindowEx(Temp_Process(i).Handle, 0, "ComboBoxEx32", vbNullString)
        
        If hWnd1 > 0 Then
            hWnd2 = FindWindowEx(hWnd1, 0, "ComboBox", vbNullString)
            
            If hWnd2 > 0 Then
                hWnd1 = FindWindowEx(hWnd2, 0, "Edit", vbNullString)
                
                Mark_Process(i) = True
                Dialog_Count = Dialog_Count + 1
            Else
                Mark_Process(i) = False
            End If
        Else
            Mark_Process(i) = False
        End If
    Next
    
    
    
    j = 0
    ReDim DIALOG_PROCESS(Dialog_Count)
    For i = 0 To UBound(Temp_Process) - 1
        If Mark_Process(i) Then
            DIALOG_PROCESS(j) = Temp_Process(i)
            j = j + 1
        End If
    Next
    Search_All_Dialog_Process = DIALOG_PROCESS
End Function
'/*********************�p�حק諸(2009/05/13)**********************/




'/*********************�Q�ο�J�Ȱ����z�����A�h��X�Ĥ@�ӲŦX���󪺵{��*********************/
Function Search_First_Windows_By_Conditions(Optional ByVal Handle As Long, Optional ByVal Class As String, Optional ByVal Name As String) As SYSTEM_PROCESS
    Dim i As Long
    Dim Result_Flag As Integer
    Dim Result_Need_Flag As Integer
    Dim Result_Process As SYSTEM_PROCESS
    
    
    
    All_Process_Flag = 0
    
    If EnumAllWindows() > 0 Then
        Result_Need_Flag = 0
        If Handle <> 0 Then
            Result_Need_Flag = Result_Need_Flag + 1
        End If
        If Class <> "" Then
            Result_Need_Flag = Result_Need_Flag + 1
        End If
        If Name <> "" Then
            Result_Need_Flag = Result_Need_Flag + 1
        End If
        
        
        
        If Result_Need_Flag > 0 Then
            For i = 0 To All_Process_Flag - 1
                Result_Flag = 0
                
                If Handle <> 0 And All_Process(i).Handle = Handle Then
                    Result_Flag = Result_Flag + 1
                End If
                If Class <> "" And All_Process(i).Class = Class Then
                    Result_Flag = Result_Flag + 1
                End If
                If Name <> "" And All_Process(i).Name = Name Then
                    Result_Flag = Result_Flag + 1
                End If
                
                If Result_Flag >= Result_Need_Flag Then
                    Result_Process = All_Process(i)
                    Exit For
                End If
            Next
        End If
    End If
    
    Search_First_Windows_By_Conditions = Result_Process
End Function
'/*********************�p�حק諸(2009/05/13)**********************/



'/*********************�Q�ο�J�Ȱ����z�����A�h��X�Ҧ��ŦX���󪺵{��*********************/
Function Search_All_Windows_By_Conditions(Optional ByVal Handle As Long, Optional ByVal Class As String, Optional ByVal Name As String) As SYSTEM_PROCESS()
    Dim i As Long
    Dim j As Long
    
    Dim Result_Flag As Integer
    Dim Result_Need_Flag As Integer
    
    Dim Mark_Process() As Boolean
    
    Dim Result_Process() As SYSTEM_PROCESS
    Dim Result_Process_Count As Long
    
    
    
    All_Process_Flag = 0
    Result_Process_Count = 0
    
    If EnumAllWindows() > 0 Then
        ReDim Mark_Process(All_Process_Flag)
    
    
        Result_Need_Flag = 0
        If Handle <> 0 Then
            Result_Need_Flag = Result_Need_Flag + 1
        End If
        If Class <> "" Then
            Result_Need_Flag = Result_Need_Flag + 1
        End If
        If Name <> "" Then
            Result_Need_Flag = Result_Need_Flag + 1
        End If
        
        
        If Result_Need_Flag > 0 Then
            For i = 0 To All_Process_Flag - 1
                Result_Flag = 0
                If Handle <> 0 And All_Process(i).Handle = Handle Then
                    Result_Flag = Result_Flag + 1
                End If
                If Class <> "" And All_Process(i).Class = Class & Chr(0) Then
                    Result_Flag = Result_Flag + 1
                End If
                If Name <> "" And Trim(All_Process(i).Name) = Name Then
                    Result_Flag = Result_Flag + 1
                End If
                
                
                
                If Result_Flag >= Result_Need_Flag Then
                    Mark_Process(i) = True
                    Result_Process_Count = Result_Process_Count + 1
                Else
                    Mark_Process(i) = False
                End If
            Next
        End If
    End If
    
    
    
    j = 0
    ReDim Result_Process(Result_Process_Count)
    For i = 0 To All_Process_Flag - 1
        If Mark_Process(i) Then
            Result_Process(j) = All_Process(i)
            j = j + 1
        End If
    Next
    Search_All_Windows_By_Conditions = Result_Process
End Function
'/*********************�p�حק諸(2009/05/13)**********************/



'/*********************�a�|�Ҧ����{��*********************/
Function EnumAllWindows() As Long
    On Error GoTo errout
    

    Call EnumChildWindows(GetDesktopWindow, AddressOf EnumChildProcess, ByVal 0&)
    
    EnumAllWindows = All_Process_Flag
    
    If False Then
errout:
            EnumAllWindows = 0
    End If
End Function
'/*********************�p�حק諸(2009/05/13)**********************/



'/*********************�a�|�Ҧ����l�{��*********************/
Public Function EnumChildProcess(ByVal hWnd As Long, ByVal lParam As Long) As Boolean
    Dim ProcessName As String
    Dim ProcessClass As String
    
    
    ProcessName = Space(GetWindowTextLength(hWnd) + 1)
    Call GetWindowText(hWnd, ProcessName, Len(ProcessName))
    ProcessName = Left(ProcessName, Len(ProcessName) - 1)
    
    
    ProcessClass = Space(256)
    Call GetClassName(hWnd, ProcessClass, 256)
    
    
    All_Process(All_Process_Flag).Handle = hWnd
    All_Process(All_Process_Flag).Class = Trim(ProcessClass)
    All_Process(All_Process_Flag).Name = ProcessName
    All_Process_Flag = All_Process_Flag + 1
    

    EnumChildProcess = True
End Function
'/*********************�p�حק諸(2009/05/13)**********************/




'/*********************����w�{�ǥ[�J��t�Φ�C*********************/
Public Function AddToSystemTray(ByVal hWnd As Long, ByVal vlngCallbackMessage As Long, ByVal vipdIcon As IPictureDisp, ByVal vstrTip As String) As Long
    Dim nidTemp As NOTIFYICONDATA
  
  
    mlngID = mlngID + 1
  
    nidTemp.cbSize = Len(nidTemp)
    nidTemp.hWnd = hWnd
    nidTemp.uID = mlngID
    nidTemp.uFlags = NIF_MESSAGE + NIF_ICON + NIF_TIP
    nidTemp.uCallbackMessage = vlngCallbackMessage
    nidTemp.hIcon = CLng(vipdIcon)
    nidTemp.szTip = vstrTip & vbNullChar
    
    Call mcolNID.Add(hWnd, CStr(mlngID))

    Call Shell_NotifyIconA(NIM_ADD, nidTemp)
  
    AddToSystemTray = mlngID
End Function
'/*********************�p�حק諸(2009/05/20)**********************/



'/*********************��ʫ��w�{�Ǧb�t�Φ�C����T*********************/
Public Sub ModifySystemTrayMessage(ByVal vlngID As Long, ByVal vlngCallbackMessage As Long)
    Dim nidTemp As NOTIFYICONDATA
  

    nidTemp.cbSize = Len(nidTemp)
    nidTemp.hWnd = mcolNID(CStr(vlngID))
    nidTemp.uID = vlngID
    nidTemp.uFlags = NIF_MESSAGE
    nidTemp.uCallbackMessage = vlngCallbackMessage
    nidTemp.hIcon = 0
    nidTemp.szTip = vbNullChar

    Call Shell_NotifyIconA(NIM_MODIFY, nidTemp)
End Sub
'/*********************�p�حק諸(2009/05/20)**********************/



'/*********************��ʫ��w�{�Ǧb�t�Φ�C���ϥ�*********************/
Public Sub ModifySystemTrayIcon(ByVal vlngID As Long, ByVal vipdIcon As IPictureDisp)
    Dim nidTemp As NOTIFYICONDATA
  

    nidTemp.cbSize = Len(nidTemp)
    nidTemp.hWnd = mcolNID(CStr(vlngID))
    nidTemp.uID = vlngID
    nidTemp.uFlags = NIF_ICON
    nidTemp.uCallbackMessage = 0
    nidTemp.hIcon = CLng(vipdIcon)
    nidTemp.szTip = vbNullChar

    Call Shell_NotifyIconA(NIM_MODIFY, nidTemp)
End Sub
'/*********************�p�حק諸(2009/05/20)**********************/




'/*********************��ʫ��w�{�Ǧb�t�Φ�C�����ܰT��*********************/
Public Sub ModifySystemTrayTip(ByVal vlngID As Long, ByVal vstrTip As String)
    Dim nidTemp As NOTIFYICONDATA
  

    nidTemp.cbSize = Len(nidTemp)
    nidTemp.hWnd = mcolNID(CStr(vlngID))
    nidTemp.uID = vlngID
    nidTemp.uFlags = NIF_TIP
    nidTemp.uCallbackMessage = 0
    nidTemp.hIcon = 0
    nidTemp.szTip = vstrTip & vbNullChar

    Call Shell_NotifyIconA(NIM_MODIFY, nidTemp)
End Sub
'/*********************�p�حק諸(2009/05/20)**********************/




'/*********************����w�{�Ǳq�t�Φ�C����*********************/
Public Sub DeleteFromSystemTray(ByVal vlngID As Long)
    Dim nidTemp As NOTIFYICONDATA
  

    nidTemp.cbSize = Len(nidTemp)
    nidTemp.hWnd = mcolNID(CStr(vlngID))
    nidTemp.uID = vlngID
    nidTemp.uFlags = NIF_MESSAGE + NIF_ICON + NIF_TIP

    Call Shell_NotifyIconA(NIM_DELETE, nidTemp)
End Sub
'/*********************�p�حק諸(2009/05/20)**********************/




'/*********************������L��w�A�H������w�S�w������ƥ�A�~��b�S���o��J�I�����p�U����A��o��**********************/
Public Sub DisableGetKeyBoardEvent()
    If hNowHook <> 0 Then
        Call UnhookWindowsHookEx(hNowHook)
        hNowHook = 0
    End If
End Sub
'/*********************�p�حק諸(2009/06/16)**********************/


'/***********�}����L��w�A�H��w�S�w������ƥ�A�~��b�S���o��J�I�����p�U�]��o��**********************/
Public Function EnableGetKeyBoardEvent(ByVal KeyCode As Long) As Long
    hKeyCode = KeyCode

    If hNowHook <> 0 Then
        EnableGetKeyBoardEvent = 0
    Else
        hNowHook = SetWindowsHookEx(WH_KEYBOARD, AddressOf EnumGetKeyBoardEvent, App.hInstance, App.ThreadID)
    
        If hNowHook <> 0 Then
            EnableGetKeyBoardEvent = hNowHook
        Else
            EnableGetKeyBoardEvent = 0
        End If
    End If
End Function
'/*********************�p�حק諸(2009/06/16)**********************/


'/*********************���_���P�_�O�_�����S�w������**********************/
Public Function EnumGetKeyBoardEvent(ByVal iCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If iCode < 0 Then
        EnumGetKeyBoardEvent = CallNextHookEx(hNowHook, iCode, wParam, lParam)
    Else
        If wParam = hKeyCode And Not hHaveKeyEvent Then
            hHaveKeyEvent = True
            EnumGetKeyBoardEvent = 1
        Else
            hHaveKeyEvent = False
            EnumGetKeyBoardEvent = 0
        End If
    End If
End Function
'/*********************�p�حק諸(2009/06/16)**********************/

