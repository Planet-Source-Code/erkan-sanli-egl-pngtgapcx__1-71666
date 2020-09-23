VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLoadTGA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Original Code : CodeId=56537 (PNG Class By Alfred Koppold)
' Revision By Erkan Þanlý 2009
Option Explicit

Private Type HEADER '18 Byte
    IDSize          As Byte      '1
    ColorMapType    As Byte      '1
    ImageType       As Byte      '1     0=none,1=indexed,2=rgb,3=grey,>8=rle
    ColorMapStart   As Integer   '2
    ColorMapLenght  As Integer   '2
    ColorMapBits    As Byte      '1
    XStart          As Integer   '2
    YStart          As Integer   '2
    Width           As Integer   '2
    Height          As Integer   '2
    Bits            As Byte      '1
    Descriptor      As Byte      '1
End Type

Private Type PALLET2
    Byte1           As Byte
    Byte2           As Byte
End Type

Private mHeader         As HEADER
Private mWidth          As Long
Private mHeight         As Long
Private mBitmapData()   As Byte

Private mNumFile        As Integer
Private mNumPixel       As Long

Public Sub LoadTGA(ByVal Filename As String)
    
    mNumFile = FreeFile
    Open Filename For Binary As #mNumFile
        Get #mNumFile, , mHeader
        With mHeader
            mWidth = .Width - .XStart
            mHeight = .Height - .YStart
            mNumPixel = mWidth * mHeight
            Select Case .Bits
                Case 8:  Read_bd08
                Case 16: Read_bd16
                Case 24: Read_bd24
                Case 32: Read_bd32
            End Select
        End With
    Close #mNumFile
    Call DrawBitmap(mWidth, mHeight)

End Sub

Private Sub Read_bd08()
    
    Dim Pal(255)    As RGBTRIPLE
    Dim Pal16()     As PALLET2
    Dim Pal24()     As RGBTRIPLE
    Dim Pal32()     As RGBQUAD
    Dim idx         As Integer
    
    Select Case mHeader.ColorMapBits
        Case 16
            ReDim Pal16(mHeader.ColorMapLenght - 1)
            Get #mNumFile, , Pal16
            For idx = 0 To UBound(Pal16)
                If GetByte(Pal16(idx).Byte1, 1) Then Pal(idx).Blue = 16
                If GetByte(Pal16(idx).Byte1, 2) Then Pal(idx).Blue = Pal(idx).Blue + 8
                If GetByte(Pal16(idx).Byte1, 3) Then Pal(idx).Blue = Pal(idx).Blue + 4
                If GetByte(Pal16(idx).Byte1, 4) Then Pal(idx).Blue = Pal(idx).Blue + 2
                If GetByte(Pal16(idx).Byte1, 5) Then Pal(idx).Blue = Pal(idx).Blue + 1
                If GetByte(Pal16(idx).Byte1, 6) Then Pal(idx).Green = 16
                If GetByte(Pal16(idx).Byte1, 7) Then Pal(idx).Green = Pal(idx).Green + 8
                If GetByte(Pal16(idx).Byte1, 8) Then Pal(idx).Green = Pal(idx).Green + 4
                If GetByte(Pal16(idx).Byte2, 1) Then Pal(idx).Green = Pal(idx).Green + 2
                If GetByte(Pal16(idx).Byte2, 2) Then Pal(idx).Green = Pal(idx).Green + 1
                If GetByte(Pal16(idx).Byte2, 3) Then Pal(idx).Red = 16
                If GetByte(Pal16(idx).Byte2, 4) Then Pal(idx).Red = Pal(idx).Red + 8
                If GetByte(Pal16(idx).Byte2, 5) Then Pal(idx).Red = Pal(idx).Red + 4
                If GetByte(Pal16(idx).Byte2, 6) Then Pal(idx).Red = Pal(idx).Red + 2
                If GetByte(Pal16(idx).Byte2, 7) Then Pal(idx).Red = Pal(idx).Red + 1
                Pal(idx).Blue = Pal(idx).Blue * 4
                Pal(idx).Green = Pal(idx).Green * 4
                Pal(idx).Red = Pal(idx).Red * 4
            Next idx
        Case 24
            ReDim Pal24(mHeader.ColorMapLenght - 1)
            Get #mNumFile, , Pal24
            For idx = 0 To UBound(Pal24)
                Pal(idx).Red = Pal24(idx).Blue 'Attention !! Red = Blue (rgb>>bgra)
                Pal(idx).Green = Pal24(idx).Green
                Pal(idx).Blue = Pal24(idx).Red 'Attention !! Blue = Red (rgb>>bgra)
            Next idx
        Case 32
            ReDim Pal32(mHeader.ColorMapLenght - 1)
            Get #mNumFile, , Pal32
            For idx = 0 To UBound(Pal32)
                Pal(idx).Red = Pal32(idx).rgbRed
                Pal(idx).Green = Pal32(idx).rgbGreen
                Pal(idx).Blue = Pal32(idx).rgbBlue
            Next idx
        Case Else
            ReDim mBitmapData(LOF(mNumFile) - Len(mHeader))
            For idx = 0 To 255
                Pal(idx).Blue = idx
                Pal(idx).Green = idx
                Pal(idx).Red = idx
            Next idx
    End Select
    Call QuickColorTable_08(Pal)
    ReDim mBitmapData(LOF(mNumFile) - Len(mHeader) - 765) '765=((mHeader.ColorMapLenght - 1) * 3))
    Get #mNumFile, , mBitmapData()
    If mHeader.ImageType > 8 Then
        MsgBox "Compression under consruction"
    Else
        Call MakeBitmap(mWidth, mHeight, mBitmapData)
        Call CreateBitmap(mWidth, mHeight, mBitmapData, bd_08)
    End If
    
End Sub

Private Sub Read_bd16()
    
    ReDim mBitmapData(LOF(mNumFile) - Len(mHeader))
    Get #mNumFile, , mBitmapData()
    If mHeader.ImageType > 8 Then
        MsgBox "Compression under consruction"
    Else
        Call MakeBitmap(mWidth * 2, mHeight, mBitmapData)
        Call CreateBitmap(mWidth, mHeight, mBitmapData, bd_16)
    End If
    
End Sub

Private Sub Read_bd24()
    
    Dim quad()      As RGBQUAD
    Dim idx         As Long
    Dim idx3        As Long
    
    ReDim mBitmapData(LOF(mNumFile) - Len(mHeader))
    Get #mNumFile, , mBitmapData()
    If mHeader.ImageType > 8 Then
        MsgBox "Compression under consruction"
    Else
        ReDim quad(UBound(mBitmapData) / 3)
        For idx = 0 To UBound(mBitmapData) / 3 - 1
            With quad(idx)
                idx3 = idx * 3
                .rgbBlue = mBitmapData(idx3)
                .rgbGreen = mBitmapData(idx3 + 1)
                .rgbRed = mBitmapData(idx3 + 2)
            End With
        Next idx
        ReDim mBitmapData(UBound(quad) * 4 + 4)
        CopyMemory mBitmapData(0), quad(0), UBound(mBitmapData)
        Call CreateBitmap(mWidth, mHeight, mBitmapData, bd_32)
    End If
    
End Sub

Private Sub Read_bd32()
    
    Dim quad()      As RGBQUAD
    Dim idx         As Long
    Dim idx4        As Long
    
    ReDim mBitmapData(LOF(mNumFile) - Len(mHeader))
    Get #mNumFile, , mBitmapData()
    If mHeader.ImageType > 8 Then
        MsgBox "Compression under consruction"
    Else
        If mAlpha = True Then
            Call MakeAlpha(mWidth, mHeight, mBitmapData)
            Call CreateBitmap(mWidth, mHeight, mBitmapData, bd_24)
        Else
            ReDim quad(UBound(mBitmapData) / 4)
            For idx = 0 To UBound(mBitmapData) / 4 - 1
                With quad(idx)
                    idx4 = idx * 4
                    .rgbBlue = mBitmapData(idx4)
                    .rgbGreen = mBitmapData(idx4 + 1)
                    .rgbRed = mBitmapData(idx4 + 2)
                End With
            Next idx
            ReDim mBitmapData(UBound(quad) * 4 + 4)
            CopyMemory mBitmapData(0), quad(0), UBound(mBitmapData)
            Call CreateBitmap(mWidth, mHeight, mBitmapData, bd_32)
        End If
    End If
    
End Sub

Private Function GetByte(Bytes As Byte, Position As Byte) As Byte
    
    GetByte = 0
    Select Case Position
        Case 1: If Bytes And 128 Then GetByte = 1   ' 128= 10000000
        Case 2: If Bytes And 64 Then GetByte = 1    ' 64 = 01000000
        Case 3: If Bytes And 32 Then GetByte = 1    ' 32 = 00100000
        Case 4: If Bytes And 16 Then GetByte = 1    ' 16 = 00010000
        Case 5: If Bytes And 8 Then GetByte = 1     ' 8  = 00001000
        Case 6: If Bytes And 4 Then GetByte = 1     ' 4  = 00000100
        Case 7: If Bytes And 2 Then GetByte = 1     ' 2  = 00000010
        Case 8: If Bytes And 1 Then GetByte = 1     ' 1  = 00000001
    End Select

End Function
