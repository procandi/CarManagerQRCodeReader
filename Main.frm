VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "QRCodeReader"
   ClientHeight    =   6375
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   12210
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   12210
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.Timer TimerDeleteList 
      Interval        =   5000
      Left            =   8520
      Top             =   1080
   End
   Begin VB.ListBox ListNoAccess 
      Height          =   5820
      Left            =   9120
      TabIndex        =   5
      Top             =   240
      Width           =   2892
   End
   Begin VB.Timer TimerAutoCapture 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8520
      Top             =   600
   End
   Begin VB.TextBox StatusText1 
      Height          =   372
      Left            =   600
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   6000
      Width           =   5892
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   8520
      Top             =   120
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "�}�l���y"
      Height          =   372
      Left            =   7320
      TabIndex        =   2
      Top             =   5520
      Width           =   1692
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   264
      Left            =   7560
      TabIndex        =   1
      Text            =   "100"
      Top             =   6000
      Width           =   852
   End
   Begin VB.PictureBox Picture1 
      Height          =   5172
      Left            =   480
      ScaleHeight     =   5115
      ScaleWidth      =   7875
      TabIndex        =   0
      Top             =   120
      Width           =   7932
   End
   Begin VB.Timer TimerSyncCapture 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   120
      Top             =   600
   End
   Begin VB.Timer TimerConvert 
      Interval        =   120
      Left            =   120
      Top             =   120
   End
   Begin VB.Label Label1 
      Height          =   372
      Left            =   480
      TabIndex        =   3
      Top             =   5400
      Width           =   2652
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ConvertEXE As String
Dim QRCodePng As String
Dim QRCodeTXTOutPut As String

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal nID As Long) As Long
Private Const GET_FRAME As Long = 1084
Private Const COPY As Long = 1054
Private Const CONNECT As Long = 1034
Private Const DISCONNECT As Long = 1035
Private CapHwnd As Long

Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Dim cc As Integer
Private Sub cmdStart_Click()
    TimerAutoCapture.Enabled = True
End Sub

Private Sub Form_Load()

    ConvertEXE = InputINI("Setting", "EXEPath", App.Path & "\QRCodeReader.ini")
    QRCodePng = InputINI("Setting", "QRCodePNGPath", App.Path & "\QRCodeReader.ini")
    QRCodeTXTOutPut = InputINI("Setting", "QRCodeOutPut", App.Path & "\QRCodeReader.ini")
    
    CapHwnd = capCreateCaptureWindow("WebcamCapture", 0, 0, 0, 640, 480, Me.hWnd, 0)
    SendMessage CapHwnd, CONNECT, 0, 0
    TimerSyncCapture.Enabled = True
End Sub

Public Sub GetQRCode(ByVal n As String)
    Dim fso As New FileSystemObject
    Open App.Path & "\Exec.bat" For Output As #1
        Print #1, ConvertEXE & " " & n & " >> " & QRCodeTXTOutPut & vbCrLf & ""
    Close #1
    pId = Shell(App.Path & "\Exec.bat", vbHide)
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, 0, pId)
    Do
        Call GetExitCodeProcess(hProcess, ExitCode)
        DoEvents
    Loop While ExitCode = STILL_ALIVE
    Call CloseHandle(hProcess)
    
    '/�p�G�����temp�o���ɥN����
    Dim ConvertContent As String
    Dim temp As String
    Dim LinePos As Integer
    Dim QRDecode As String
    If fso.FileExists(QRCodeTXTOutPut) Then
        Open QRCodeTXTOutPut For Input As #2
        Do While Not EOF(2)
            Line Input #2, temp
            ConvertContent = ConvertContent & temp
        Loop
        Close #2
        Debug.Print "===================================="
        QRDecode = GetDeCodeTXT(ConvertContent)
    
        Dim CheckFlag As Boolean
        CheckFlag = True
        If QRDecode <> "" Then
            bee
            Timer3.Enabled = True
            Label1.Caption = "���oQRCODE =>" & QRDecode
            '�[�Jlistbox
            For i = 0 To ListNoAccess.ListCount
                If InStr(1, ListNoAccess.List(i), QRDecode) > 0 Then
                    Debug.Print "��Ʃ|������"
                    CheckFlag = False
                End If
            Next
            If CheckFlag = True Then
                ListNoAccess.AddItem (QRDecode & cc)
                '/�T�{�O�_�n�g�J��Ʈw
                Access (QRDecode)
            End If

        End If
        fso.DeleteFile (QRCodeTXTOutPut)
        fso.DeleteFile (n)
    End If
    
    'Timer1.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SendMessage CapHwnd, DISCONNECT, 0, 0
   TimerConvert.Enabled = False
End Sub

Public Function GetDeCodeTXT(ByVal s As String) As String
    Dim Pos1 As Integer
    Dim Pos2 As Integer
    Pos1 = InStr(1, s, "Raw result:")
    Pos2 = InStr(1, s, "Parsed result:")
    If Pos1 = 0 And Pos2 = 0 Then Exit Function
    
    GetDeCodeTXT = Mid(s, Pos1 + Len("Raw result:") + 1, Pos2 - 2 - (Pos1 + Len("Raw result:")))
    GetDeCodeTXT = Replace(GetDeCodeTXT, vbCrLf, "")
End Function

Private Sub Timer2_Timer()

End Sub

Private Sub Timer3_Timer()
    Label1.Caption = ""
    Timer3.Enabled = False
End Sub


Public Sub Access(cid)
    Dim uID As String
    Dim status As String
    Dim n As Integer
    Dim dollar As Integer
    Dim newdollar As Integer
    Dim conn As New LiteConnection
    Dim record As New LiteStatement
    
    '�s�u��Ʈw�θ�ƪ�
    Call conn.Open(App.Path & "\CarManager")
    record.ActiveConnection = conn
    
    '�ˬd�O�_�����d���s�b

    sqlstring = "select * from userlist where cid='" & cid & "'"
    Call record.Prepare(sqlstring)
    n = record.RowCount
    If n > 0 Then   '�p�G�����
        Call record.Step(1)
        dollar = record.ColumnValue("dollar")
        uID = record.ColumnValue("uid")
        
        
        '�ˬd���ϥΪ̬O�n�����٬O�i��
        Call record.Close
        sqlstring = "select * from checkinout where cid='" & cid & "' order by date desc, time desc"
        Call record.Prepare(sqlstring)
        n = record.RowCount
        If n > 0 Then   '�p�G����ơA�h��X��
            Call record.Step(1)
            status = record.ColumnValue("status")
            If status = "�i�J" Then
                status = "���}"
                newdollar = dollar
            Else
                status = "�i�J"
                newdollar = dollar - 40
            End If
        Else    '�S����ƪ��ܡA�ܤ֨����i��b�̭��A�@�w�O�n�i��
            status = "�i�J"
            newdollar = dollar - 40
        End If
        
        If newdollar < 0 And status = "�i�J" Then   '�p�G�����O�i�J�����p�U�A���l�B�S�����A���Ӧ^��
            Debug.Print "�l�B�Ⱦl" & dollar & "�A�����\�q�L"
            StatusText1.Text = "�l�B�Ⱦl" & dollar & "�A�����\�q�L"
        Else
            '�}�l��s��Ʈw�����
            Call record.Close
                        
            '�W�[�i�X�O��
            sqlstring = "insert into CheckInOut (date,time,status,cid,uid) values ("
            sqlstring = sqlstring & ("'" & Format(Now, "YYYY/MM/DD") & "',")
            sqlstring = sqlstring & ("'" & Format(Now, "hh:mm:ss") & "',")
            sqlstring = sqlstring & ("'" & status & "',")
            sqlstring = sqlstring & ("'" & cid & "',")
            sqlstring = sqlstring & ("'" & uID & "'")
            sqlstring = sqlstring & ")"
            Call conn.Execute(sqlstring)
            
            If status = "�i�J" Then '�p�G�����O�i�J�����p�U�A�~�ݭn�����ڡA��W�C����O��
                '��s�l�B
                sqlstring = "update userlist set dollar='" & newdollar & "' where cid='" & cid & "'"
                Call conn.Execute(sqlstring)
                
                '�W�[����O��
                sqlstring = "insert into Dollar (date,time,uid,status,dollar) values ("
                sqlstring = sqlstring & ("'" & Format(Now, "YYYY/MM/DD") & "',")
                sqlstring = sqlstring & ("'" & Format(Now, "hh:mm:ss") & "',")
                sqlstring = sqlstring & ("'" & uID & "',")
                sqlstring = sqlstring & ("'����',")
                sqlstring = sqlstring & ("'40'")
                sqlstring = sqlstring & ")"
                Call conn.Execute(sqlstring)
            End If
            
            '�@�����`�����p
            Debug.Print "�w" & status & "�A�l�B�|��:" & newdollar & "��"
            StatusText1.Text = "�w" & status & "�A�l�B�|��:" & newdollar & "��"
        End If
    Else    '�S��ƪ����p
        Debug.Print "���d�����s�b�A�����\�q�L"
        StatusText1.Text = "���d�����s�b�A�����\�q�L"
    End If
    
    
    '������Ʈw
    Call record.Close
    Call conn.Close
    
    '��s�e���W�����A
    'RxList.AddItem Format(Now, "hh:mm:ss") & StatusText1.Text, 0
End Sub

Private Sub TimerAutoCapture_Timer()
    SaveJPG Picture1, QRCodePng & Timer & ".jpg", 80
End Sub

Private Sub TimerConvert_Timer()
    Dim fso As New FileSystemObject
    Dim fi As File
    For Each fi In fso.GetFolder(QRCodePng).Files
        GetQRCode (fi.Path)
    Next
End Sub

Private Sub TimerDeleteList_Timer()
    If ListNoAccess.ListCount >= 1 Then
        ListNoAccess.RemoveItem (0)
    End If
End Sub

Private Sub TimerSyncCapture_Timer()
   On Error Resume Next
   TimerSyncCapture.Interval = Val(Text1.Text)
   SendMessage CapHwnd, GET_FRAME, 0, 0
   SendMessage CapHwnd, COPY, 0, 0
   Picture1.Picture = Clipboard.GetData
   Clipboard.Clear
End Sub

Public Sub bee()
    Beep 400, 200
End Sub

