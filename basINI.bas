Attribute VB_Name = "basINI"
'/******************************************************************/
'/*�����G�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@*/
'/*��m��iniŪ���B�g�J��������ƪ��a��C                           */
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
'/*                                      Last Edit Date 2009/02/26 */
'/******************************************************************/
Option Explicit



'/*********************************ReadINI Example***********************************/
'ReadINI_Return = ReadIni("Tag", "Name", "", ReadINI_String, Len(ReadINI_String), "INI")
'Result = Trim(Left(ReadINI_String, ReadINI_Return))
'/**************************�p�حק諸(2009/04/09)***********************************/

'/*��ini�B�z������Win32API�`��*/
Public Declare Function ReadINI Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long 'Ū��ini�ɩһݭn�Ψ쪺WindowsAPI
'/**/

'/**************************��ini�B�z�������ܼ�***********************************/
Public ReadINI_Return As Integer '�ΰ��^��ini���e�ɡA�O���^�Ǧr�����
Public ReadINI_String As String * 256 '�Ω�^��ini���e�ɡA�ȥΪ��w�Ħr��
'/**************************�p�حק諸(2009/02/25)***********************************/



'/************�Ω�����K���B�zŪini�ɪ����D�A���禡�i�H�۰ʥh�ťո�B�zŪ������ini�����D********/
Public Function InputINI(ByVal ClassName As String, ByVal TitleName As String, ByVal FileName As String) As String
    Dim InputINI_Return As Integer
    Dim InputINI_String As String * 256
    Dim Result_String As String
        
    InputINI_Return = ReadINI(ClassName, TitleName, "", InputINI_String, Len(InputINI_String), FileName)
    Result_String = Trim(Left(InputINI_String, InputINI_Return))
    
    If Len(Result_String) = 0 Then
        InputINI = Result_String
    Else
        Do While Len(Result_String) > 0 And (Asc(Right(Result_String, 1)) = 32 Or Asc(Right(Result_String, 1)) = 0 Or Asc(Right(Result_String, 1)) = 77 Or Asc(Right(Result_String, 1)) = 121)
            Result_String = Left(Result_String, Len(Result_String) - 1)
        Loop
        
        InputINI = Result_String
    End If
End Function
'/**************************�p�حק諸(2009/04/16)***********************************/


