Attribute VB_Name = "basProcess"
'/******************************************************************/
'/*說明：　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*放置跟系統程序處理相關的地方。                                  */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*模組：　　　　　　　　　　　　　　　    　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*參考：    　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*元件：    　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*注意事項：　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*無。                                                            */
'/*　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*                                           Edit By Edward Hsieh */
'/*                                      Last Edit Date 2009/05/13 */
'/******************************************************************/
Option Explicit

'/******************處理程序用的Win32API常數******************/

'/*系統佇列用*/
Public Const NIM_ADD  As Long = &H0
Public Const NIM_DELETE  As Long = &H2
Public Const NIM_MODIFY  As Long = &H1
Public Const NIF_ICON  As Long = &H2
Public Const NIF_MESSAGE  As Long = &H1
Public Const NIF_TIP  As Long = &H4
'/**/

'/*系統命令用*/
Public Const WM_CLOSE  As Long = &H10
Public Const WM_CLEAR  As Long = &H303
Public Const WM_CHAR  As Long = &H102
'/**/


'/*鍵盤事件用*/
Public Const WM_KEYDOWN  As Long = &H100
Public Const WM_KEYUP  As Long = &H101
Public Const WH_KEYBOARD As Long = 2
'/**/

'/*滑鼠事件用*/
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

'/*變換視窗的狀態用*/
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


'/*改變視窗形狀用*/
Public Const LWA_COLORKEY As Long = &H1
Public Const LWA_ALPHA As Long = &H2
Public Const GWL_EXSTYLE As Long = -20
Public Const WS_EX_LAYERED As Long = &H80000
'/**/


'/*要對程序下的指令*/
Public Const PROCESS_QUERY_INFORMATION = &H400
'/**/


'/*檢查程序目前狀態用*/
Public Const STILL_ALIVE = &H103
'/**/

'/***********************************************************/

'/******************處理程序用的Win32API結構******************/

'/*系統程序在指定或接收資料時會用到的結構*/
Public Type EVENTMSG
        message As Long
        paramL As Long
        paramH As Long
        time As Long
        hWnd As Long
End Type
'/**/
'/*系統佇列會用到的結構*/
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

'/****************處理程序用的Win32API函式*******************/
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




'/********************跟系統程序處理相關的常數*********************/
Public Const MAX_PROCESS As Long = 755739 '本作業系列最多可擁有的程序數量
'/*********************小華修改的(2009/05/13)**********************/


'/********************跟系統程序處理相關的結構*********************/
'/*一個系統程序應該會有的結構*/
Public Type SYSTEM_PROCESS
    Handle As Long
    Class As String
    Name As String
End Type
'/*小華修改的(2009/05/13)*/
'/*********************小華修改的(2009/05/13)**********************/


'/********************跟系統程序處理相關的變數*********************/

'/*儲存列舉視窗用到的*/
Public All_Process(MAX_PROCESS) As SYSTEM_PROCESS '用來儲存列舉所有程序時，暫存用的變數
Public All_Process_Flag As Long '用來儲存列舉所有程序時，使用者總共有輸入幾個篩選條件
'/**/

'/*縮小及回復到系統列會用到的*/
Public mlngID As Long '用來儲存這次縮小到系統列時是第幾次
Public mcolNID As New Collection '用來儲存歷次縮小到系統列時的hWnd
'/*小華修改的(2009/05/13)*/

'/*儲存視窗盒的變數*/
Public Dialog_OutputName As String '要自動輸進Dialog物件的文字
Public DialogName As String '視窗盒對話方塊的標題名稱
Public DialogType As String '視窗盒對話方塊的類型代碼
'/**/

'/*儲存目前抓到的系統鍵盤代碼*/
Public hNowHook As Long
Public hKeyCode As Long
Public hHaveKeyEvent As Boolean
'/*小華修改的(2009/06/16)*/

'/*********************小華修改的(2009/05/13)**********************/



'/*********************利用輸入值做為篩選條件，去找出所有符合條件的程序，且該程序必需是Dialog程序*********************/
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
'/*********************小華修改的(2009/05/13)**********************/




'/*********************利用輸入值做為篩選條件，去找出第一個符合條件的程序*********************/
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
'/*********************小華修改的(2009/05/13)**********************/



'/*********************利用輸入值做為篩選條件，去找出所有符合條件的程序*********************/
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
'/*********************小華修改的(2009/05/13)**********************/



'/*********************窮舉所有的程序*********************/
Function EnumAllWindows() As Long
    On Error GoTo errout
    

    Call EnumChildWindows(GetDesktopWindow, AddressOf EnumChildProcess, ByVal 0&)
    
    EnumAllWindows = All_Process_Flag
    
    If False Then
errout:
            EnumAllWindows = 0
    End If
End Function
'/*********************小華修改的(2009/05/13)**********************/



'/*********************窮舉所有的子程序*********************/
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
'/*********************小華修改的(2009/05/13)**********************/




'/*********************把指定程序加入到系統佇列*********************/
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
'/*********************小華修改的(2009/05/20)**********************/



'/*********************更動指定程序在系統佇列的資訊*********************/
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
'/*********************小華修改的(2009/05/20)**********************/



'/*********************更動指定程序在系統佇列的圖示*********************/
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
'/*********************小華修改的(2009/05/20)**********************/




'/*********************更動指定程序在系統佇列的提示訊息*********************/
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
'/*********************小華修改的(2009/05/20)**********************/




'/*********************把指定程序從系統佇列移除*********************/
Public Sub DeleteFromSystemTray(ByVal vlngID As Long)
    Dim nidTemp As NOTIFYICONDATA
  

    nidTemp.cbSize = Len(nidTemp)
    nidTemp.hWnd = mcolNID(CStr(vlngID))
    nidTemp.uID = vlngID
    nidTemp.uFlags = NIF_MESSAGE + NIF_ICON + NIF_TIP

    Call Shell_NotifyIconA(NIM_DELETE, nidTemp)
End Sub
'/*********************小華修改的(2009/05/20)**********************/




'/*********************關閉鍵盤鎖定，以停止鎖定特定的按鍵事件，才能在沒有得到焦點的情況下停止再抓得到**********************/
Public Sub DisableGetKeyBoardEvent()
    If hNowHook <> 0 Then
        Call UnhookWindowsHookEx(hNowHook)
        hNowHook = 0
    End If
End Sub
'/*********************小華修改的(2009/06/16)**********************/


'/***********開啟鍵盤鎖定，以鎖定特定的按鍵事件，才能在沒有得到焦點的情況下也抓得到**********************/
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
'/*********************小華修改的(2009/06/16)**********************/


'/*********************不斷的判斷是否有抓到特定的按鍵**********************/
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
'/*********************小華修改的(2009/06/16)**********************/

