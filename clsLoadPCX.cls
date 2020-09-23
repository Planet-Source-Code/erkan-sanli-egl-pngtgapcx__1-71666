VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLoadPCX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Original Code : CodeId=56537 (PNG Class By Alfred Koppold)
' Revision By Erkan Þanlý 2009
Option Explicit

Private Type HEADER '30 Byte
    Manufacturer            As Byte      '1
    Version                 As Byte      '1
    Encoding                As Byte      '1
    Bpp                     As Byte      '1
    Xmin                    As Integer   '2
    Ymin                    As Integer   '2
    Xmax                    As Integer   '2
    Ymax                    As Integer   '2
    HDpi                    As Integer   '2
    VDpi                    As Integer   '2
    ColorPalette(15)        As RGBTRIPLE '3
    Reserved1               As Byte      '1
    Planes                  As Byte      '1
    BytesPerLine            As Integer   '2
    PaletteInfo             As Integer   '2
    HScreenSize             As Integer   '2
    VScreenSize             As Integer   '2
    Reserved2(53)           As Byte      '1
End Type

Private mHeader             As HEADER
Private mWidth              As Long
Private mHeight             As Long
Private mBitmapData()       As Byte

Private mNumFile            As Integer
Private mLineSize           As Long

Public Sub LoadPCX(ByVal Filename As String)
    
    mNumFile = FreeFile
    Open Filename For Binary Lock Write As #mNumFile
        With mHeader
            Get #mNumFile, , mHeader
            ReDim mBitmapData(LOF(mNumFile) - Len(mHeader))
            Get #mNumFile, , mBitmapData()
            If .Encoding = 1 Then Call DecompRLE(mBitmapData)
            mWidth = .Xmax - .Xmin + 1
            mHeight = .Ymax - .Ymin + 1
            mLineSize = .Planes * .BytesPerLine
            Select Case .Bpp
                Case 1: If .Planes = 1 Then Call Read_bd01
                Case 4: If .Planes = 1 Then Call Read_bd04
                Case 8: If .Planes = 1 Then Call Read_bd08 Else Call Read_bd24
            End Select
        End With
    Close #mNumFile
    Call DrawBitmap(mWidth, mHeight)
   
    
End Sub

Private Sub Read_bd01()
    
    Call CreateTable(ct_GRAY, bd_01)
    Call MakeBitmap(mLineSize, mHeight, mBitmapData, True)
    Call CreateBitmap(mWidth, mHeight, mBitmapData, bd_01)

End Sub

Private Sub Read_bd04()
        
    Call QuickColorTable_04(mHeader.ColorPalette)
    Call MakeBitmap(mLineSize, mHeight, mBitmapData, True)
    Call CreateBitmap(mWidth, mHeight, mBitmapData, bd_04)

End Sub

Private Sub Read_bd08()

    Dim ColorPalette(255) As RGBTRIPLE
    Dim PalByte As Byte
    Dim idx As Long

    Seek #mNumFile, LOF(mNumFile) - 768
    Get #mNumFile, , PalByte
    If PalByte = 12 Then
        Seek #mNumFile, LOF(mNumFile) - 767
        Get #mNumFile, , ColorPalette()
    Else
        For idx = 0 To 255
            ColorPalette(idx).Blue = idx
            ColorPalette(idx).Green = idx
            ColorPalette(idx).Red = idx
        Next idx
    End If

    Call QuickColorTable_08(ColorPalette)
    Call MakeBitmap(mLineSize, mHeight, mBitmapData, True)
    Call CreateBitmap(mWidth, mHeight, mBitmapData, bd_08)

End Sub

Private Function Read_bd24()
    
    Dim quad()  As RGBQUAD
    Dim X       As Long
    Dim Y       As Long
    Dim idx     As Long
    
    ReDim quad(mWidth * mHeight)
    mHeight = mHeight - 1
    For Y = 0 To mHeight
        For X = 0 To mWidth - 1
            idx = Y * mHeader.BytesPerLine * 3 + X
            With quad((X + (mHeight - Y) * mWidth))
                .rgbRed = mBitmapData(idx)
                .rgbGreen = mBitmapData(idx + mHeader.BytesPerLine)
                .rgbBlue = mBitmapData(idx + mHeader.BytesPerLine * 2)
            End With
        Next X
    Next Y
    ReDim mBitmapData(UBound(quad) * 4 + 4)
    CopyMemory mBitmapData(0), quad(0), UBound(mBitmapData) + 1
    mHeight = mHeight + 1
    Call CreateBitmap(mWidth, mHeight, mBitmapData, bd_32)
    
End Function

Private Sub DecompRLE(Buffer() As Byte)
    
    Dim TempBuffer() As Byte
    Dim State        As Long
    Dim idx          As Long
    Dim IdxRepeat    As Long
    Dim NumRepeat    As Long
         
    For idx = 0 To UBound(Buffer) - 1
        NumRepeat = Buffer(idx)
        If NumRepeat >= 192 Then
            idx = idx + 1
            For IdxRepeat = 1 To NumRepeat - 192
                ReDim Preserve TempBuffer(State)
                TempBuffer(State) = Buffer(idx)
                State = State + 1
            Next IdxRepeat
        Else
            ReDim Preserve TempBuffer(State)
            TempBuffer(State) = NumRepeat
            State = State + 1
        End If
    Next idx
    Buffer = TempBuffer

End Sub
