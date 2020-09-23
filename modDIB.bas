Attribute VB_Name = "modDIB"
Option Explicit

Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0
Private Const CBM_INIT = &H4

Public Enum BIT_DEPTH
    bd_01 = 1       '2 Color (b-w)          - 1/8 Byte per pixel
    bd_02 = 2       '4 Color                - 1/4 Byte per pixel
    bd_04 = 4       '16 Color               - 1/2 Byte per pixel
    bd_08 = 8       '256 Color              - 1   Byte per pixel
    bd_16 = 16      '65536 Color            - 2   Bytes per pixel
    bd_24 = 24      '16777216 Color         - 3   Bytes per pixel
    bd_32 = 32      '16777216 Color + Alpha - 4   Bytes per pixel
End Enum

Public Enum COLOR_TYPE 'for png
    ct_GRAY = 0     'Grayscale
    ct_RGB = 2      'RGB
    ct_PAL = 3      'Color Palettes
    ct_GRAYA = 4    'Grayscale + Alpha
    ct_RGBA = 6     'RGB + Alpha
End Enum

Public Enum FILTER_TYPE 'for png
    ft_NONE
    ft_SUB
    ft_UP
    ft_AVERAGE
    ft_PAETH
End Enum

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Type RGBTRIPLE
    Red                 As Byte
    Green               As Byte
    Blue                As Byte
End Type

Public Type RGBQUAD
    rgbBlue             As Byte
    rgbGreen            As Byte
    rgbRed              As Byte
    rgbReserved         As Byte
End Type

Public Type BITMAPINFOHEADER
    biSize              As Long
    biWidth             As Long
    biHeight            As Long
    biPlanes            As Integer
    biBitCount          As Integer
    biCompression       As Long
    biSizeImage         As Long
    biXPelsPerMeter     As Long
    biYPelsPerMeter     As Long
    biClrUsed           As Long
    biClrImportant      As Long
End Type

Private Type BITMAPINFO_01
    bmiHeader           As BITMAPINFOHEADER
    bmiColors(1)        As RGBQUAD
End Type

Private Type BITMAPINFO_02
    bmiHeader           As BITMAPINFOHEADER
    bmiColors(3)        As RGBQUAD
End Type

Private Type BITMAPINFO_04
    bmiHeader           As BITMAPINFOHEADER
    bmiColors(15)       As RGBQUAD
End Type

Private Type BITMAPINFO_08
    bmiHeader           As BITMAPINFOHEADER
    bmiColors(255)      As RGBQUAD
End Type

Private Type BITMAPINFO_16
    bmiHeader           As BITMAPINFOHEADER
    bmiColors           As RGBQUAD
End Type

Private Type BITMAPINFO_24
    bmiHeader           As BITMAPINFOHEADER
    bmiColors           As RGBTRIPLE
End Type

Private Type BITMAPINFO_32
    bmiHeader           As BITMAPINFOHEADER
    bmiColors           As RGBQUAD
End Type

Private Declare Function CreateDIBitmap_01 Lib "gdi32" Alias "CreateDIBitmap" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO_01, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBitmap_02 Lib "gdi32" Alias "CreateDIBitmap" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO_02, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBitmap_04 Lib "gdi32" Alias "CreateDIBitmap" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO_04, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBitmap_08 Lib "gdi32" Alias "CreateDIBitmap" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO_08, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBitmap_16 Lib "gdi32" Alias "CreateDIBitmap" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO_16, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBitmap_24 Lib "gdi32" Alias "CreateDIBitmap" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO_24, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBitmap_32 Lib "gdi32" Alias "CreateDIBitmap" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO_32, ByVal wUsage As Long) As Long

Private bmi_01          As BITMAPINFO_01
Private bmi_02          As BITMAPINFO_02
Private bmi_04          As BITMAPINFO_04
Private bmi_08          As BITMAPINFO_08
Private bmi_16          As BITMAPINFO_16
Private bmi_24          As BITMAPINFO_24
Private bmi_32          As BITMAPINFO_32
Private hBmp            As Long
Public mColorPalette()  As Byte

Public mAlpha           As Boolean
Public mPicBox          As PictureBox

Public Sub CreateBitmap(ByVal Width As Long, ByVal Height As Long, Buffer() As Byte, BDepth As BIT_DEPTH, Optional Orientation As Boolean = False)
    
    Dim hdc     As Long
    Dim bmih    As BITMAPINFOHEADER
    Dim Bits()  As RGBTRIPLE
    Dim BitsA() As RGBQUAD
    
    If Orientation Then Height = -Height
    With bmih
        .biSize = Len(bmih)
        .biWidth = Width
        .biHeight = Height
        .biPlanes = 1
        .biBitCount = BDepth
        .biCompression = BI_RGB
        .biSizeImage = 0
        .biXPelsPerMeter = 0
        .biYPelsPerMeter = 0
        .biClrUsed = 0
        .biClrImportant = 0
    End With
    hdc = GetDC(0)
    
    Select Case BDepth
        Case bd_01
            CopyMemory bmi_01.bmiHeader, bmih, bmih.biSize
            hBmp = CreateDIBitmap_01(hdc, bmi_01.bmiHeader, CBM_INIT, Buffer(0), bmi_01, DIB_RGB_COLORS)
        Case bd_02
            CopyMemory bmi_02.bmiHeader, bmih, bmih.biSize
            hBmp = CreateDIBitmap_02(hdc, bmi_02.bmiHeader, CBM_INIT, Buffer(0), bmi_02, DIB_RGB_COLORS)
         Case bd_04
            CopyMemory bmi_04.bmiHeader, bmih, bmih.biSize
            hBmp = CreateDIBitmap_04(hdc, bmi_04.bmiHeader, CBM_INIT, Buffer(0), bmi_04, DIB_RGB_COLORS)
        Case bd_08
            CopyMemory bmi_08.bmiHeader, bmih, bmih.biSize
            hBmp = CreateDIBitmap_08(hdc, bmi_08.bmiHeader, CBM_INIT, Buffer(0), bmi_08, DIB_RGB_COLORS)
        Case bd_16
            CopyMemory bmi_16.bmiHeader, bmih, bmih.biSize
            hBmp = CreateDIBitmap_16(hdc, bmi_16.bmiHeader, CBM_INIT, Buffer(0), bmi_16, DIB_RGB_COLORS)
        Case bd_24
            ReDim Bits((UBound(Buffer) / 3) - 1)
            CopyMemory Bits(0), Buffer(0), UBound(Buffer)
            CopyMemory bmi_24.bmiHeader, bmih, bmih.biSize
            hBmp = CreateDIBitmap_24(hdc, bmi_24.bmiHeader, CBM_INIT, Bits(0), bmi_24, DIB_RGB_COLORS)
        Case bd_32
            ReDim BitsA((UBound(Buffer) / 4) - 1)
            CopyMemory BitsA(0), Buffer(0), UBound(Buffer)
            CopyMemory bmi_32.bmiHeader, bmih, bmih.biSize
            bmi_32.bmiHeader.biBitCount = bd_32
            hBmp = CreateDIBitmap_32(hdc, bmi_32.bmiHeader, CBM_INIT, BitsA(0), bmi_32, DIB_RGB_COLORS)
    End Select

End Sub

Public Sub DrawBitmap(Width As Long, Height As Long)
    
    Dim cDC As Long
    
    mPicBox.Cls
    If hBmp Then
        cDC = CreateCompatibleDC(mPicBox.hdc)
        SelectObject cDC, hBmp
        Call StretchBlt(mPicBox.hdc, 0, 0, Width, Height, cDC, 0, 0, Width, Height, vbSrcCopy)
        DeleteDC cDC
        DeleteObject hBmp
        hBmp = 0
    End If

End Sub

Public Sub MakeBitmap(LineSize As Long, Height As Long, BitmapData() As Byte, Optional Orientation As Boolean = False)
    
    Dim Übergabe() As Byte
    Dim Zugabe As Integer
    Dim Standort As Long
    Dim Width32 As Long
    Dim idx As Long

    If (LineSize) Mod 4 = 0 Then
        Width32 = LineSize - 1
    Else
        Width32 = (LineSize \ 4) * 4 + 3
    End If
    If Width32 + 1 <> LineSize Then Zugabe = Width32 - LineSize + 1
    ReDim Übergabe(UBound(BitmapData))
    CopyMemory Übergabe(0), BitmapData(0), UBound(BitmapData) + 1
    ReDim BitmapData(Height * (Width32 + 1) - 1)
    
    If Orientation Then
        For idx = 0 To LineSize * Height - LineSize Step LineSize
            CopyMemory BitmapData(Standort), Übergabe((LineSize * Height) - idx - LineSize), LineSize
            Standort = Standort + Width32 + 1
        Next idx
    Else
        For idx = 0 To LineSize * Height - LineSize Step LineSize
            CopyMemory BitmapData(Standort), Übergabe(idx), LineSize
            Standort = Standort + Width32 + 1
        Next idx
    End If
    
End Sub

Public Sub MakeAlpha(mWidth As Long, mHeight As Long, Buffer() As Byte, Optional Orientation As Boolean = False, Optional X As Long = 0, Optional Y As Long = 0)

    Dim Myx As Long, Myy As Long, DatOff As Long
    Dim R   As Long, G   As Long, b      As Long, A  As Long
    Dim sR  As Long, sG  As Long, sB     As Long
    Dim dR  As Long, dG  As Long, dB     As Long
    Dim DestData() As Byte, bytesperrow  As Long
    Dim DestOff As Long, DestHdr As BITMAPINFOHEADER
    Dim MemDC As Long, hBmp As Long, hOldBmp As Long
    Dim SrcData() As Byte
    Dim hdc As Long
    Dim Height As Long
    Dim Width As Long
    
    On Error Resume Next

    Height = mHeight
    Width = mWidth
    hdc = mPicBox.hdc
    
    If mPicBox.Width < Width * Screen.TwipsPerPixelX Then
        mPicBox.Width = Screen.TwipsPerPixelX * Width + 100
    End If
    If mPicBox.Height < Height * Screen.TwipsPerPixelY Then
        mPicBox.Height = Screen.TwipsPerPixelY * Height + 100
    End If
    hdc = mPicBox.hdc
    bytesperrow = LineBytes(Width, 24)
    ReDim DestData(bytesperrow * Height - 1)
    ReDim SrcData(UBound(Buffer))
    DestHdr.biBitCount = 24
    If Orientation Then
        DestHdr.biHeight = -Height
    Else
        DestHdr.biHeight = Height
    End If
    DestHdr.biWidth = Width
    DestHdr.biPlanes = 1
    DestHdr.biSize = 40
    MemDC = CreateCompatibleDC(hdc)
    hBmp = CreateCompatibleBitmap(hdc, Width, Height)
    hOldBmp = SelectObject(MemDC, hBmp)
    BitBlt MemDC, 0, 0, Width, Height, hdc, X, Y, vbSrcCopy
    GetDIBits MemDC, hBmp, 0, Height, SrcData(0), DestHdr, 0
    SelectObject hOldBmp, MemDC
    DeleteObject hBmp
    DeleteDC MemDC
    For Myy = 0 To Height - 1
        For Myx = 0 To Width - 1
            DestOff = Myy * bytesperrow + Myx * 3
            sR = SrcData(DestOff + 2)
            sG = SrcData(DestOff + 1)
            sB = SrcData(DestOff)
            b = Buffer(DatOff)
            G = Buffer(DatOff + 1)
            R = Buffer(DatOff + 2)
            A = Buffer(DatOff + 3)
            If A = 255 Then
                DestData(DestOff + 2) = R
                DestData(DestOff + 1) = G
                DestData(DestOff) = b
            ElseIf A = 0 Then
                DestData(DestOff + 2) = sR
                DestData(DestOff + 1) = sG
                DestData(DestOff) = sB
            Else
                dR = R * A + (255 - A) * sR + 255
                dG = G * A + (255 - A) * sG + 255
                dB = b * A + (255 - A) * sB + 255
                CopyMemory DestData(DestOff + 2), ByVal VarPtr(dR) + 1, 1
                CopyMemory DestData(DestOff + 1), ByVal VarPtr(dG) + 1, 1
                CopyMemory DestData(DestOff), ByVal VarPtr(dB) + 1, 1
            End If
            DatOff = DatOff + 4
        Next Myx
    Next Myy
 Buffer = DestData
End Sub

Public Sub CreateTable(CType As COLOR_TYPE, BDepth As BIT_DEPTH, Optional To8Bit As Boolean = False)
    
    Dim TableRGBA() As RGBQUAD
    Dim TableRGB()  As RGBTRIPLE
    Dim LevelDiff   As Byte
    Dim CurLevel    As Integer
    Dim idx         As Long
    
    ReDim TableRGBA(2 ^ BDepth - 1)
    Select Case CType
        Case ct_GRAY
            LevelDiff = 255 \ UBound(TableRGBA)
            For idx = 0 To UBound(TableRGBA)
                TableRGBA(idx).rgbRed = CurLevel
                TableRGBA(idx).rgbGreen = CurLevel
                TableRGBA(idx).rgbBlue = CurLevel
                CurLevel = CurLevel + LevelDiff
            Next idx
        Case ct_PAL
            ReDim TableRGB(UBound(TableRGBA))
            CopyMemory TableRGB(0), mColorPalette(0), UBound(mColorPalette) + 1
            For idx = 0 To UBound(TableRGBA)
                TableRGBA(idx).rgbRed = TableRGB(idx).Red
                TableRGBA(idx).rgbGreen = TableRGB(idx).Green
                TableRGBA(idx).rgbBlue = TableRGB(idx).Blue
            Next idx
    End Select
        
    Select Case BDepth
        Case bd_01
            If To8Bit Then
                CopyMemory ByVal VarPtr(bmi_08.bmiColors(0).rgbBlue), ByVal VarPtr(TableRGBA(0).rgbBlue), 8
            Else
                CopyMemory ByVal VarPtr(bmi_01.bmiColors(0).rgbBlue), ByVal VarPtr(TableRGBA(0).rgbBlue), 8
            End If
        Case bd_02
'            If To8Bit Then
                CopyMemory ByVal VarPtr(bmi_08.bmiColors(0).rgbBlue), ByVal VarPtr(TableRGBA(0).rgbBlue), 16
'            Else
'                CopyMemory ByVal VarPtr(bmi_02.bmiColors(0).rgbBlue), ByVal VarPtr(TableRGBA(0).rgbBlue), 16
'            End If
        Case bd_04
            If To8Bit Then
                CopyMemory ByVal VarPtr(bmi_08.bmiColors(0).rgbBlue), ByVal VarPtr(TableRGBA(0).rgbBlue), 64
            Else
                CopyMemory ByVal VarPtr(bmi_04.bmiColors(0).rgbBlue), ByVal VarPtr(TableRGBA(0).rgbBlue), 64
            End If
        Case bd_08
            CopyMemory ByVal VarPtr(bmi_08.bmiColors(0).rgbBlue), ByVal VarPtr(TableRGBA(0).rgbBlue), 1024
    End Select

End Sub

Public Sub QuickColorTable_04(Palette() As RGBTRIPLE)

    Dim idx As Integer

    For idx = 0 To UBound(bmi_04.bmiColors)
        bmi_04.bmiColors(idx).rgbRed = Palette(idx).Red
        bmi_04.bmiColors(idx).rgbGreen = Palette(idx).Green
        bmi_04.bmiColors(idx).rgbBlue = Palette(idx).Blue
    Next idx

End Sub

Public Sub QuickColorTable_08(Palette() As RGBTRIPLE)
    
    Dim idx As Integer
    
    For idx = 0 To UBound(bmi_08.bmiColors)
        bmi_08.bmiColors(idx).rgbRed = Palette(idx).Red
        bmi_08.bmiColors(idx).rgbGreen = Palette(idx).Green
        bmi_08.bmiColors(idx).rgbBlue = Palette(idx).Blue
    Next idx
    
End Sub

Private Function LineBytes(Width As Long, BitCount As Integer) As Long
    
    LineBytes = ((Width * BitCount + 31) \ 32) * 4

End Function

