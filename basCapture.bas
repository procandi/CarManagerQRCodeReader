Attribute VB_Name = "basCapture"
Option Explicit
Private Type GUID
    Data1  As Long
    Data2  As Integer
    Data3  As Integer
    Data4(0 To 7)    As Byte
End Type
Private Type GdiplusStartupInput
    GdiplusVersion  As Long
    DebugEventCallback  As Long
    SuppressBackgroundThread  As Long
    SuppressExternalCodecs  As Long
End Type
Private Type EncoderParameter
    GUID  As GUID
    NumberOfvalues  As Long
type  As Long
    value  As Long
End Type
Private Type EncoderParameters
    Count  As Long
    Parameter  As EncoderParameter
End Type
Private Declare Function GdiplusStartup Lib "GDIPlus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdiplusShutdown Lib "GDIPlus" (ByVal token As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" (ByVal hbm As Long, ByVal hpal As Long, Bitmap As Long) As Long
Private Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal Image As Long) As Long
Private Declare Function GdipSaveImageToFile Lib "GDIPlus" (ByVal Image As Long, ByVal filename As Long, clsidEncoder As GUID, encoderParams As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal str As Long, id As GUID) As Long

'  ----====  SaveJPG  ====----

Public Sub SaveJPG(ByVal pict As StdPicture, ByVal filename As String, Optional ByVal quality As Byte = 80)
    Dim tSI   As GdiplusStartupInput
    Dim lRes   As Long
    Dim lGDIP   As Long
    Dim lBitmap   As Long
    '  Initialize  GDI+
    tSI.GdiplusVersion = 1
    lRes = GdiplusStartup(lGDIP, tSI)
    If lRes = 0 Then
        '  Create  the  GDI+  bitmap
        '  from  the  image  handle
        lRes = GdipCreateBitmapFromHBITMAP(pict.Handle, 0, lBitmap)
        If lRes = 0 Then
            Dim tJpgEncoder   As GUID
            Dim tParams   As EncoderParameters
            '  Initialize  the  encoder  GUID
            CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
            '  Initialize  the  encoder  parameters
            tParams.Count = 1
            With tParams.Parameter   '  Quality
                '  Set  the  Quality  GUID
                CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB3505E7EB}"), .GUID
                .NumberOfvalues = 1
                .type = 1
                .value = VarPtr(quality)
            End With
            '  Save  the  image
            lRes = GdipSaveImageToFile(lBitmap, StrPtr(filename), tJpgEncoder, tParams)
            '  Destroy  the  bitmap
            GdipDisposeImage lBitmap
        End If
        '  Shutdown  GDI+
        GdiplusShutdown lGDIP
    End If
    If lRes Then
        Err.Raise 5, , "Cannot  save  the  image.  GDI+  Error:" & lRes
    End If
End Sub
