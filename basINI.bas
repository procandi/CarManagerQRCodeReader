Attribute VB_Name = "basINI"
'/******************************************************************/
'/*說明：　　　　　　　　　　　　　　　　　　　　　　　　　　　　　*/
'/*放置跟ini讀取、寫入等相關資料的地方。                           */
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
'/*                                      Last Edit Date 2009/02/26 */
'/******************************************************************/
Option Explicit



'/*********************************ReadINI Example***********************************/
'ReadINI_Return = ReadIni("Tag", "Name", "", ReadINI_String, Len(ReadINI_String), "INI")
'Result = Trim(Left(ReadINI_String, ReadINI_Return))
'/**************************小華修改的(2009/04/09)***********************************/

'/*跟ini處理有關的Win32API常數*/
Public Declare Function ReadINI Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long '讀取ini檔所需要用到的WindowsAPI
'/**/

'/**************************跟ini處理有關的變數***********************************/
Public ReadINI_Return As Integer '用除回傳ini內容時，記錄回傳字串長度
Public ReadINI_String As String * 256 '用於回傳ini內容時，暫用的緩衝字串
'/**************************小華修改的(2009/02/25)***********************************/



'/************用於比較方便的處理讀ini檔的問題，此函式可以自動去空白跟處理讀取中文ini的問題********/
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
'/**************************小華修改的(2009/04/16)***********************************/


