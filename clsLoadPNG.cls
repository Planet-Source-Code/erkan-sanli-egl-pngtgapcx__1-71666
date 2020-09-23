VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLoadPNG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Original Code : CodeId=56537 (PNG Class By Alfred Koppold)
' Revision By Erkan Þanlý 2009
Option Explicit

Private Type HEADER '13 Byte
    Width               As Long     '4
    Height              As Long     '4
    Bitdepht            As Byte     '1
    ColorType           As Byte     '1
    Compression         As Byte     '1
    Filter              As Byte     '1
    Interlacing         As Byte     '1
End Type

Private mHeader         As HEADER
Private mWidth          As Long
Private mHeight         As Long
Private mBitmapData()   As Byte
  
Private mTrans          As Boolean
Private mTransData()    As Byte
Private mBPP            As Byte

Public Sub LoadPNG(Filename As String)
    
    Dim Filenumber      As Long
    Dim LenIDAT         As Long
    Dim ChunkSize       As Long
    Dim ChunkType       As String * 4
    Dim Seeker          As Long
    Dim Signature(1)    As Long
    Dim IDATData()      As Byte
    Dim Buffer()        As Byte
    Dim CRC32           As Long
    Dim DataSize        As Long

'Reset data
    LenIDAT = 0
    ReDim IDATData(LenIDAT)
'Open file
    Filenumber = FreeFile
    Open Filename For Binary As Filenumber
'Check signature
        Get Filenumber, , Signature
        If Signature(0) <> &H474E5089 Or Signature(1) <> &HA1A0A0D Then Exit Sub
        Do
'Read chunk
            Get Filenumber, , ChunkSize
            Get Filenumber, , ChunkType
            Call Swap4Bytes(ChunkSize)
            If ChunkSize > 0 Then ReDim Buffer(ChunkSize - 1)
'Check EOF
            Seeker = Seek(Filenumber)
            If Seeker + ChunkSize > LOF(Filenumber) Then Exit Sub
'Read Buffer
            Get Filenumber, , Buffer
            Get Filenumber, , CRC32
            Select Case ChunkType
                Case "IHDR"
                    With mHeader
                        CopyMemory .Width, Buffer(0), 4:  Call Swap4Bytes(.Width)
                        CopyMemory .Height, Buffer(4), 4: Call Swap4Bytes(.Height)
                        .Bitdepht = Buffer(8)
                        .ColorType = Buffer(9)
                        .Compression = Buffer(10)
                        .Filter = Buffer(11)
                        .Interlacing = Buffer(12)
                    End With
                Case "PLTE"
                    mColorPalette = Buffer
                Case "IDAT"
                    ReDim Preserve IDATData(LenIDAT + UBound(Buffer))
                    CopyMemory IDATData(LenIDAT), Buffer(0), UBound(Buffer) + 1
                    LenIDAT = UBound(IDATData) + 1
                Case "IEND"
                    Exit Do
                Case "tRNS"
                    mTrans = True
                    mTransData = Buffer
            End Select
        Loop
        If LenIDAT = 0 Then Exit Sub
    Close Filenumber
    Erase Buffer
'Copy IDAT to mBitmapData and set mWidth and mHeight
    mWidth = mHeader.Width
    mHeight = mHeader.Height
    ReDim mBitmapData(UBound(IDATData) - 2)
    CopyMemory mBitmapData(0), IDATData(2), UBound(IDATData) - 1
    Erase IDATData
'Decompress
    If mHeader.Compression = 0 Then
        DataSize = mHeight * DataPerRowBytes
        If mHeader.Interlacing Then DataSize = DataSize + mHeight
        Call DecompLZ77(DataSize, mBitmapData)
    End If
'Defilter
    If mHeader.Interlacing Then
        Call DeFilterInterlaced
    Else
        Call DeFilter
    End If
'Process
    Call DataProcess
    Call DrawBitmap(mWidth, mHeight)

End Sub

Private Sub DataProcess()

    With mHeader
'Create Bitmap
        Select Case .ColorType
            
            Case ct_GRAY
                Select Case .Bitdepht
                    
                    Case bd_01
                        If .Interlacing Then
                            Call CreateTable(ct_GRAY, bd_01, True)
                            Call CreateBitmap(mWidth, mHeight, mBitmapData, bd_08, True)
                        Else
                            Call CreateTable(ct_GRAY, bd_01)
                            Call CreateBitmap(mWidth, mHeight, mBitmapData, bd_01, True)
                        End If
                    
                    Case bd_02
                        If .Interlacing = 0 Then _
                        Call BitsToBytes(2, (UBound(mBitmapData) + 1) * 4, mBitmapData) ' 2 >> 8
                        Call CreateTable(ct_GRAY, bd_02)
                        Call CreateBitmap(mWidth, mHeight, mBitmapData, bd_08, True)
                    
                    Case bd_04
                        If .Interlacing Then
                            Call CreateTable(ct_GRAY, bd_04, True)
                            Call CreateBitmap(mWidth, mHeight, mBitmapData, bd_08, True)
                        Else
                            Call Align32(bd_04)
                            Call CreateTable(ct_GRAY, bd_04)
                            Call CreateBitmap(mWidth, mHeight, mBitmapData, bd_04, True)
                        End If
                    
                    Case bd_08
                        Call Align32(bd_08)
                        Call CreateTable(ct_GRAY, bd_08)
                        Call CreateBitmap(mWidth, mHeight, mBitmapData, bd_08, True)
                    
                    Case bd_16
                        Call Conv16To8 '16 >> 8
                        Call CreateTable(ct_GRAY, bd_08)
                        Call CreateBitmap(mWidth, mHeight, mBitmapData, bd_08, True)
                End Select
                
            Case ct_RGB
                If .Bitdepht = bd_16 Then Call Conv16To8
                Call SwapRGB
                mBPP = bd_08
                Call Align32(bd_24)
                Call CreateBitmap(mWidth, mHeight, mBitmapData, bd_24, True)
            
            Case ct_PAL
                Select Case .Bitdepht
                    
                    Case bd_01
                        If .Interlacing Then
                            Call CreateTable(ct_PAL, bd_08)
                            Call CreateBitmap(mWidth, mHeight, mBitmapData, bd_08, True)
                        Else
                            Call CreateTable(ct_PAL, bd_01)
                            Call CreateBitmap(mWidth, mHeight, mBitmapData, bd_01, True)
                        End If
                    
                    Case bd_02
                        If .Interlacing Then
                            mBPP = bd_08
                            Call Align32(bd_08)
                        Else
                            Call BitsToBytes(2, (UBound(mBitmapData) + 1) * 4, mBitmapData) ' 2 >> 8
                        End If
                        Call CreateTable(ct_PAL, bd_08)
                        Call CreateBitmap(mWidth, mHeight, mBitmapData, bd_08, True)
                    
                    Case bd_04
                        If .Interlacing Then
                            mBPP = bd_08
                            Call Align32(bd_08)
                            Call CreateTable(ct_PAL, bd_08)
                            Call CreateBitmap(mWidth, mHeight, mBitmapData, bd_08, True)
                        Else
                            Call Align32(bd_04)
                            If mTrans Then
                                Call PalToRGBA(bd_04)
                                Call MakeAlpha(mWidth, mHeight, mBitmapData, True)
                                Call CreateBitmap(mWidth, mHeight, mBitmapData, bd_24, True)
                            Else
                                Call CreateTable(ct_PAL, bd_04)
                                Call CreateBitmap(mWidth, mHeight, mBitmapData, bd_04, True)
                            End If
                        End If
                    
                    Case bd_08
                            Call Align32(bd_08)
                            If mTrans Then
                                Call PalToRGBA(bd_08)
                                Call MakeAlpha(mWidth, mHeight, mBitmapData, True)
                                Call CreateBitmap(mWidth, mHeight, mBitmapData, bd_24, True)
                            Else
                                Call CreateTable(ct_PAL, bd_08)
                                Call CreateBitmap(mWidth, mHeight, mBitmapData, bd_08, True)
                            End If
                End Select
            
            Case ct_GRAYA
                If .Bitdepht = bd_16 Then Call Conv16To8
                Call GrayAToRGBA
                Call MakeAlpha(mWidth, mHeight, mBitmapData, True)
                Call CreateBitmap(mWidth, mHeight, mBitmapData, bd_24, True)
            
            Case ct_RGBA
                If .Bitdepht = bd_16 Then Call Conv16To8
                Call SwapRGBA
                Call MakeAlpha(mWidth, mHeight, mBitmapData, True)
                Call CreateBitmap(mWidth, mHeight, mBitmapData, bd_24, True)
                
        End Select
    End With
    
End Sub

Private Function DataPerRowBytes() As Long
    
    DataPerRowBytes = (mWidth * BitsPerPixel + 7) \ 8 + 1

End Function

Private Function RowBytes(Depth As BIT_DEPTH) As Long
    
    RowBytes = ((mWidth * Depth + 31) \ 32) * 4

End Function

Private Function BitsPerPixel() As Long
    
    Dim Bpp As Long

    If mBPP <> mHeader.Bitdepht And mBPP <> 0 Then
        Bpp = mBPP
    Else
        Bpp = mHeader.Bitdepht
    End If
    
    Select Case mHeader.ColorType
        Case 0: BitsPerPixel = Bpp      'Grayscale
        Case 2: BitsPerPixel = Bpp * 3  'RGB
        Case 3: BitsPerPixel = Bpp      'Palette
        Case 4: BitsPerPixel = Bpp * 2  'Grayscale + Alpha
        Case 6: BitsPerPixel = Bpp * 4  'RGB + Alpha
    End Select
    
End Function

Private Sub Conv16To8()
    
    Dim DstOff      As Long
    Dim DstBuffer() As Byte
    Dim idx         As Long
    
    ReDim DstBuffer(UBound(mBitmapData) \ 2)
    For idx = 0 To UBound(mBitmapData) Step 2
        DstBuffer(DstOff) = mBitmapData(idx)
        DstOff = DstOff + 1
    Next idx
    mBitmapData = DstBuffer

End Sub

Private Sub Align32(Depth As BIT_DEPTH)

    Dim DstRowBytes As Long
    Dim SrcRowBytes As Long
    Dim DstOff      As Long
    Dim SrcOff      As Long
    Dim DstBuffer() As Byte
    Dim idx         As Long
       
    DstRowBytes = RowBytes(Depth)
    SrcRowBytes = DataPerRowBytes - 1
    
    If mHeader.ColorType = ct_GRAYA Then SrcRowBytes = SrcRowBytes / 2
    If DstRowBytes <> SrcRowBytes Then
        ReDim DstBuffer(DstRowBytes * mHeight - 1)
        For idx = 0 To mHeight - 1
            DstOff = DstRowBytes * idx
            SrcOff = SrcRowBytes * idx
            Call CopyMemory(DstBuffer(DstOff), mBitmapData(SrcOff), SrcRowBytes)
        Next idx
        mBitmapData = DstBuffer
    End If
    
End Sub

Private Sub GrayAToRGBA()

    Dim DstOff      As Long
    Dim DstBuffer() As Byte
    Dim idx         As Long

    ReDim DstBuffer((UBound(mBitmapData) + 1) * 2 - 1)
    For idx = 0 To UBound(mBitmapData) Step 2
        DstBuffer(DstOff) = mBitmapData(idx)
        DstBuffer(DstOff + 1) = mBitmapData(idx)
        DstBuffer(DstOff + 2) = mBitmapData(idx)
        DstBuffer(DstOff + 3) = mBitmapData(idx + 1)
        DstOff = DstOff + 4
    Next idx
    mBitmapData = DstBuffer

End Sub

Private Sub PalToRGBA(Depth As BIT_DEPTH)
 
     Dim DestBuffer()   As Byte
     Dim n              As Long
     Dim PalEntry       As Byte
     Dim DestOff        As Long
     Dim TrnsBnd        As Long
     Dim X              As Long
     Dim Y              As Long
     Dim WidthBytes     As Long
     Dim Pal()          As RGBTRIPLE
     
     With mHeader
        ReDim DestBuffer(4 * .Width * .Height - 1)
        TrnsBnd = UBound(mTransData)
        WidthBytes = RowBytes(Depth)
        ReDim Pal((UBound(mColorPalette) \ 3))
        CopyMemory Pal(0), mColorPalette(0), UBound(mColorPalette) + 1
        Select Case Depth
            Case 8
                For Y = 0 To .Height - 1
                    For X = 0 To .Width - 1
                        n = Y * WidthBytes + X
                        PalEntry = mBitmapData(n)
                        With Pal(PalEntry)
                            DestBuffer(DestOff) = .Blue
                            DestBuffer(DestOff + 1) = .Green
                            DestBuffer(DestOff + 2) = .Red
                        End With
                        If PalEntry <= TrnsBnd Then
                            DestBuffer(DestOff + 3) = mTransData(PalEntry)
                        Else
                            DestBuffer(DestOff + 3) = 255
                        End If
                        DestOff = DestOff + 4
                    Next X
                Next Y
            Case 4
                For Y = 0 To .Height - 1
                    For X = 0 To .Width - 1
                        n = Y * WidthBytes + X \ 2
                        If (X Mod 2) = 1 Then
                            PalEntry = mBitmapData(n) And 15
                        Else
                            PalEntry = (mBitmapData(n) \ 16) And 15
                        End If
                        With Pal(PalEntry)
                            DestBuffer(DestOff) = .Blue
                            DestBuffer(DestOff + 1) = .Green
                            DestBuffer(DestOff + 2) = .Red
                        End With
                        If PalEntry <= TrnsBnd Then
                            DestBuffer(DestOff + 3) = mTransData(PalEntry)
                        Else
                            DestBuffer(DestOff + 3) = 255
                        End If
                        DestOff = DestOff + 4
                    Next X
                Next Y
            Case 1
                For Y = 0 To .Height - 1
                    For X = 0 To .Width - 1
                        n = Y * WidthBytes + X \ 8
                        If (X Mod 8) <> 7 Then
                            PalEntry = (mBitmapData(n) \ 2 ^ (7 - X Mod 8)) And 1
                        Else
                            PalEntry = mBitmapData(n) And 1
                        End If
                        With Pal(PalEntry)
                            DestBuffer(DestOff) = .Blue
                            DestBuffer(DestOff + 1) = .Green
                            DestBuffer(DestOff + 2) = .Red
                        End With
                        If PalEntry <= TrnsBnd Then
                            DestBuffer(DestOff + 3) = mTransData(PalEntry)
                        Else
                            DestBuffer(DestOff + 3) = 255
                        End If
                        DestOff = DestOff + 4
                    Next X
                Next Y
        End Select
    End With
    mBitmapData = DestBuffer
    
End Sub

Private Sub DeFilter()
    
    Dim SrcOff          As Long
    Dim DstOff          As Long
    Dim DstBuffer()     As Byte
    Dim idx             As Long
    Dim Interval        As Long
    Dim CurRowBytes()   As Byte 'Current
    Dim PrvRowBytes()   As Byte 'Previous
    Dim DataPRow        As Long
    
    DataPRow = DataPerRowBytes
    ReDim DstBuffer(UBound(mBitmapData) - mHeight)
    ReDim PrvRowBytes(DataPRow - 2)
    ReDim CurRowBytes(DataPRow - 2)
    
    Interval = Abs(BitsPerPixel / 8)
    If Interval = 0 Then Interval = 1
    
    For idx = 0 To mHeight - 1
        SrcOff = DataPRow * idx
        DstOff = SrcOff - idx
        Call CopyMemory(CurRowBytes(0), mBitmapData(SrcOff + 1), DataPRow - 1)
        Call Filter(CurRowBytes, PrvRowBytes, mBitmapData(SrcOff), Interval)
        Call CopyMemory(DstBuffer(DstOff), CurRowBytes(0), DataPRow - 1)
    Next idx
    mBitmapData = DstBuffer

End Sub

Private Sub DeFilterInterlaced()
    
    Dim Bpp             As Long
    Dim DstBuffer()     As Byte
    Dim CurRowBytes()   As Byte
    Dim PrvRowBytes()   As Byte
    Dim StdBuffer       As Long
    Dim Interval        As Long
    Dim Rest8           As Long
    Dim Height8         As Long
    Dim State           As String
    Dim ZL              As Long
    Dim MengeRow        As Long
    Dim idx             As Long
    Dim Nr              As Long
    Dim ZZ              As Long
    Dim ZLBytes         As Long
    
    Bpp = BitsPerPixel
    Interval = IIf(Bpp >= 8, Abs(Bpp / 8), 1)

    With mHeader
        ReDim DstBuffer((.Width * .Height * Interval) - 1)
        Rest8 = .Height Mod 8
        Height8 = (.Height - Rest8) / 8
        
        State = "1"
        ZL = BerechneRowLen(Bpp, State)
        If ZL > 0 Then
            ReDim PrvRowBytes(ZL - 1)
            MengeRow = Height8
            If Rest8 > 0 Then MengeRow = MengeRow + 1
            For idx = 1 To MengeRow
                ReDim CurRowBytes(ZL - 1)
                CopyMemory CurRowBytes(0), mBitmapData(StdBuffer + 1), ZL
                Call Filter(CurRowBytes, PrvRowBytes, mBitmapData(StdBuffer), Interval)
                StdBuffer = StdBuffer + ZL + 1
                If Bpp < 8 Then
                    ZLBytes = BerechneRowLen(8, State)
                Else
                    ZLBytes = 0
                End If
                PutBuffer DstBuffer, CurRowBytes, 1, 1, idx, ZLBytes
            Next idx
        End If
        
        State = "5"
        ZL = BerechneRowLen(Bpp, State)
        If ZL > 0 Then
            ReDim PrvRowBytes(ZL - 1)
            MengeRow = Height8
            If Rest8 > 0 Then MengeRow = MengeRow + 1
            For idx = 1 To MengeRow
                ReDim CurRowBytes(ZL - 1)
                CopyMemory CurRowBytes(0), mBitmapData(StdBuffer + 1), ZL
                Call Filter(CurRowBytes, PrvRowBytes, mBitmapData(StdBuffer), Interval)
                StdBuffer = StdBuffer + ZL + 1
                If Bpp < 8 Then
                    ZLBytes = BerechneRowLen(8, State)
                Else
                    ZLBytes = 0
                End If
                PutBuffer DstBuffer, CurRowBytes, 2, 1, idx, ZLBytes
            Next idx
        End If

        State = "15"
        ZL = BerechneRowLen(Bpp, State)
        If ZL > 0 Then
            ReDim PrvRowBytes(ZL - 1)
            MengeRow = Height8
            If Rest8 > 4 Then MengeRow = MengeRow + 1
            For idx = 1 To MengeRow
                ReDim CurRowBytes(ZL - 1)
                CopyMemory CurRowBytes(0), mBitmapData(StdBuffer + 1), ZL
                Call Filter(CurRowBytes, PrvRowBytes, mBitmapData(StdBuffer), Interval)
                StdBuffer = StdBuffer + ZL + 1
                If Bpp < 8 Then
                    ZLBytes = BerechneRowLen(8, State)
                Else
                    ZLBytes = 0
                End If
                PutBuffer DstBuffer, CurRowBytes, 3, 5, idx, ZLBytes
            Next idx
        End If
    
        State = "37"
        ZZ = 1
        ZL = BerechneRowLen(Bpp, State)
        If ZL > 0 Then
            ReDim PrvRowBytes(ZL - 1)
            MengeRow = Height8 * 2
            If Rest8 > 0 Then MengeRow = MengeRow + 1
            If Rest8 > 4 Then MengeRow = MengeRow + 1
            Nr = 1
            For idx = 1 To MengeRow
                ReDim CurRowBytes(ZL - 1)
                CopyMemory CurRowBytes(0), mBitmapData(StdBuffer + 1), ZL
                Call Filter(CurRowBytes, PrvRowBytes, mBitmapData(StdBuffer), Interval)
                StdBuffer = StdBuffer + ZL + 1
                If Bpp < 8 Then
                    ZLBytes = BerechneRowLen(8, State)
                Else
                    ZLBytes = 0
                End If
                PutBuffer DstBuffer, CurRowBytes, 4, Nr, ZZ, ZLBytes
                If Nr = 1 Then
                    Nr = 5
                Else
                    Nr = 1: ZZ = ZZ + 1
                End If
            Next idx
        End If
        
        State = "1357"
        ZL = BerechneRowLen(Bpp, State)
        If ZL > 0 Then
            ReDim PrvRowBytes(ZL - 1)
            MengeRow = Height8 * 2
            If Rest8 > 2 Then MengeRow = MengeRow + 1
            If Rest8 > 6 Then MengeRow = MengeRow + 1
            ZZ = 1: Nr = 3
            For idx = 1 To MengeRow
                ReDim CurRowBytes(ZL - 1)
                CopyMemory CurRowBytes(0), mBitmapData(StdBuffer + 1), ZL
                Call Filter(CurRowBytes, PrvRowBytes, mBitmapData(StdBuffer), Interval)
                StdBuffer = StdBuffer + ZL + 1
                If Bpp < 8 Then
                    ZLBytes = BerechneRowLen(8, State)
                Else
                    ZLBytes = 0
                End If
                PutBuffer DstBuffer, CurRowBytes, 5, Nr, ZZ, ZLBytes
                Select Case Nr
                    Case 3: Nr = 7
                    Case 7: Nr = 3: ZZ = ZZ + 1
                End Select
            Next idx
        End If
        
        State = "2468"
        ZL = BerechneRowLen(Bpp, State)
        If ZL > 0 Then
            ReDim PrvRowBytes(ZL - 1)
            ZZ = 1: Nr = 1
            MengeRow = Height8 * 4
            If Rest8 > 0 Then MengeRow = MengeRow + 1
            If Rest8 > 2 Then MengeRow = MengeRow + 1
            If Rest8 > 4 Then MengeRow = MengeRow + 1
            If Rest8 > 6 Then MengeRow = MengeRow + 1
            For idx = 1 To MengeRow
                ReDim CurRowBytes(ZL - 1)
                CopyMemory CurRowBytes(0), mBitmapData(StdBuffer + 1), ZL
                Call Filter(CurRowBytes, PrvRowBytes, mBitmapData(StdBuffer), Interval)
                StdBuffer = StdBuffer + ZL + 1
                If Bpp < 8 Then
                    ZLBytes = BerechneRowLen(8, State)
                Else
                    ZLBytes = 0
                End If
                PutBuffer DstBuffer, CurRowBytes, 6, Nr, ZZ, ZLBytes
                Select Case Nr
                    Case 1: Nr = 3
                    Case 3: Nr = 5
                    Case 5: Nr = 7
                    Case 7: Nr = 1: ZZ = ZZ + 1
                End Select
            Next idx
        End If

        State = "12345678"
        ZL = BerechneRowLen(Bpp, State)
        If ZL > 0 Then
            ReDim PrvRowBytes(ZL - 1)
            ZZ = 1: Nr = 2
            MengeRow = Height8 * 4
            If Rest8 > 1 Then MengeRow = MengeRow + 1
            If Rest8 > 3 Then MengeRow = MengeRow + 1
            If Rest8 > 5 Then MengeRow = MengeRow + 1
            If Rest8 > 7 Then MengeRow = MengeRow + 1
            For idx = 1 To MengeRow
                ReDim CurRowBytes(ZL - 1)
                CopyMemory CurRowBytes(0), mBitmapData(StdBuffer + 1), ZL
                Call Filter(CurRowBytes, PrvRowBytes, mBitmapData(StdBuffer), Interval)
                StdBuffer = StdBuffer + ZL + 1
                If Bpp < 8 Then
                    ZLBytes = BerechneRowLen(8, State)
                Else
                    ZLBytes = 0
                End If
                PutBuffer DstBuffer, CurRowBytes, 7, Nr, ZZ, ZLBytes
                Select Case Nr
                    Case 2: Nr = 4
                    Case 4: Nr = 6
                    Case 6: Nr = 8
                    Case 8: Nr = 2: ZZ = ZZ + 1
                End Select
            Next idx
        End If
    End With
    mBitmapData = DstBuffer
    
End Sub

Private Function BerechneRowLen(Bpp As Long, State As String) As Long
    
    Dim LenState        As Long
    Dim Remain          As Long
    Dim RemainLen       As Long
    Dim RemainBytes     As Long
    Dim NumBits         As Long
    Dim NumBytes        As Long
    Dim idx             As Long
    Dim Stack           As Long
    
    Dim NBytes          As Long
    Dim AnzRB           As Long
    
    LenState = Len(State)
    RemainLen = mWidth Mod 8
    
    For idx = 1 To LenState
        If CLng(Mid(State, idx, 1)) <= RemainLen Then
            Stack = Stack + 1
        Else
            Exit For
        End If
    Next idx
    If Bpp < 8 Then
        Remain = IIf(RemainLen > 0, Bpp * Stack, 0)
    Else
        Remain = Stack * (Bpp / 8)
    End If
    NumBytes = (mWidth - RemainLen) / 8
    NumBits = NumBytes * Bpp * LenState
    RemainBytes = NumBits Mod 8
    NBytes = (NumBits - RemainBytes) / 8
    Select Case Bpp
        Case Is < 8
            Remain = Remain + RemainBytes
            AnzRB = (Remain - Remain Mod 8) / 8
            If Remain Mod 8 Then AnzRB = AnzRB + 1
            BerechneRowLen = NBytes + AnzRB
        Case Else
            BerechneRowLen = NBytes + Remain
    End Select

End Function

Private Sub Filter(CurRowBytes() As Byte, PrvRowBytes() As Byte, FType As Byte, Interval As Long)
    
    Dim PrvOff          As Long
    Dim PrvVal          As Byte
    Dim BPRow           As Long
    Dim idx             As Long
    Dim X               As Integer
    Dim LeftPixOff      As Long
    Dim LeftPix         As Byte
    Dim UpperLeftPix    As Byte
    Dim PaethPredictor  As Byte
    Dim P(3)            As Integer
    
    BPRow = UBound(CurRowBytes) + 1
    Select Case FType 'FilterType
        'Case ft_NONE
        Case ft_SUB
            For idx = 0 To BPRow - 1
                PrvOff = idx - Interval
                If PrvOff >= 0 Then PrvVal = CurRowBytes(PrvOff)
                X = CInt(CurRowBytes(idx)) + CInt(PrvVal)
                Call CopyMemory(CurRowBytes(idx), X, 1)
            Next idx
        Case ft_UP
            For idx = 0 To BPRow - 1
                PrvVal = PrvRowBytes(idx)
                X = CInt(CurRowBytes(idx)) + CInt(PrvVal)
                Call CopyMemory(CurRowBytes(idx), X, 1)
            Next idx
        Case ft_AVERAGE
            For idx = 0 To BPRow - 1
                PrvOff = idx - Interval
                If PrvOff >= 0 Then PrvVal = CurRowBytes(PrvOff)
                X = CurRowBytes(idx) + (CInt(PrvRowBytes(idx)) + CInt(PrvVal)) \ 2
                Call CopyMemory(CurRowBytes(idx), X, 1)
            Next idx
        Case ft_PAETH
            For idx = 0 To BPRow - 1
                LeftPixOff = idx - Interval
                If LeftPixOff >= 0 Then
                    LeftPix = CurRowBytes(LeftPixOff)
                    UpperLeftPix = PrvRowBytes(LeftPixOff)
                End If
                P(0) = CInt(LeftPix) + CInt(PrvRowBytes(idx)) - CInt(UpperLeftPix)
                P(1) = Abs(P(0) - LeftPix)
                P(2) = Abs(P(0) - PrvRowBytes(idx))
                P(3) = Abs(P(0) - UpperLeftPix)
                If (P(1) <= P(2)) And (P(1) <= P(3)) Then
                    PaethPredictor = LeftPix
                ElseIf P(2) <= P(3) Then
                    PaethPredictor = PrvRowBytes(idx)
                Else
                    PaethPredictor = UpperLeftPix
                End If
                X = CInt(CurRowBytes(idx)) + CInt(PaethPredictor)
                Call CopyMemory(CurRowBytes(idx), X, 1)
            Next idx
    End Select
    PrvRowBytes = CurRowBytes

End Sub

Private Sub PutBuffer(Buffer() As Byte, Rowbuffer() As Byte, RowType As Byte, RowNumber As Long, Rowzähler As Long, RowLength As Long)
    
    Dim Anfang As Long
    Dim Achtschritt As Long
    Dim Zeile As Long
    Dim Rowanfang As Long
    Dim i As Long
    Dim StdBuffer As Long
    Dim Rowstand As Long
    Dim Größe As Long
    Dim BytesPerPixel As Long
    Dim Bpp As Long

    With mHeader
        Bpp = BitsPerPixel
        If Bpp >= 8 Then
            BytesPerPixel = Abs(Bpp / 8)
        Else
            BytesPerPixel = 1
            Call BitsToBytes(mHeader.Bitdepht, RowLength, Rowbuffer)
        End If
        Größe = UBound(Rowbuffer) + 1
        Rowanfang = .Width * (RowNumber - 1) * BytesPerPixel
        Achtschritt = .Width * 8 * BytesPerPixel
        Anfang = (Achtschritt * (Rowzähler - 1)) + Rowanfang
        'RowType: 1 = 1; 2 = 5; 3 = 1+5; 4 = 3+7; 5 = 1+3+5+7; 6 = 2+4+6+8; 7 = 1-8;
        StdBuffer = Anfang
        Select Case RowType
            Case 1
                Do While Rowstand < Größe
                    CopyMemory Buffer(StdBuffer), Rowbuffer(Rowstand), BytesPerPixel
                    StdBuffer = StdBuffer + (8 * BytesPerPixel)
                    Rowstand = Rowstand + BytesPerPixel
                Loop
            Case 2
                StdBuffer = StdBuffer + (4 * BytesPerPixel)
                Do While Rowstand < Größe
                    CopyMemory Buffer(StdBuffer), Rowbuffer(Rowstand), BytesPerPixel
                    StdBuffer = StdBuffer + (8 * BytesPerPixel)
                    Rowstand = Rowstand + BytesPerPixel
                Loop
            Case 3
                Do While Rowstand < Größe
                    CopyMemory Buffer(StdBuffer), Rowbuffer(Rowstand), BytesPerPixel
                    If Rowstand + BytesPerPixel < Größe Then CopyMemory Buffer(StdBuffer + (4 * BytesPerPixel)), Rowbuffer(Rowstand + BytesPerPixel), BytesPerPixel
                    StdBuffer = StdBuffer + (8 * BytesPerPixel)
                    Rowstand = Rowstand + (2 * BytesPerPixel)
                Loop
            Case 4
                StdBuffer = StdBuffer + (2 * BytesPerPixel)
                Do While Rowstand < Größe
                    CopyMemory Buffer(StdBuffer), Rowbuffer(Rowstand), BytesPerPixel
                    If Rowstand + BytesPerPixel < Größe Then CopyMemory Buffer(StdBuffer + (4 * BytesPerPixel)), Rowbuffer(Rowstand + BytesPerPixel), BytesPerPixel
                    StdBuffer = StdBuffer + (8 * BytesPerPixel)
                    Rowstand = Rowstand + (2 * BytesPerPixel)
                Loop
            Case 5
                Do While Rowstand < Größe
                    CopyMemory Buffer(StdBuffer), Rowbuffer(Rowstand), BytesPerPixel
                    If Rowstand + BytesPerPixel < Größe Then CopyMemory Buffer(StdBuffer + (2 * BytesPerPixel)), Rowbuffer(Rowstand + BytesPerPixel), BytesPerPixel
                    If Rowstand + (2 * BytesPerPixel) < Größe Then CopyMemory Buffer(StdBuffer + (4 * BytesPerPixel)), Rowbuffer(Rowstand + (2 * BytesPerPixel)), BytesPerPixel
                    If Rowstand + (3 * BytesPerPixel) < Größe Then CopyMemory Buffer(StdBuffer + (6 * BytesPerPixel)), Rowbuffer(Rowstand + (3 * BytesPerPixel)), BytesPerPixel
                    StdBuffer = StdBuffer + (8 * BytesPerPixel)
                    Rowstand = Rowstand + (4 * BytesPerPixel)
                Loop
            Case 6
                StdBuffer = StdBuffer + BytesPerPixel
                Do While Rowstand < Größe
                    CopyMemory Buffer(StdBuffer), Rowbuffer(Rowstand), BytesPerPixel
                    If Rowstand + BytesPerPixel < Größe Then CopyMemory Buffer(StdBuffer + (2 * BytesPerPixel)), Rowbuffer(Rowstand + BytesPerPixel), BytesPerPixel
                    If Rowstand + (2 * BytesPerPixel) < Größe Then CopyMemory Buffer(StdBuffer + (4 * BytesPerPixel)), Rowbuffer(Rowstand + (2 * BytesPerPixel)), BytesPerPixel
                    If Rowstand + (3 * BytesPerPixel) < Größe Then CopyMemory Buffer(StdBuffer + (6 * BytesPerPixel)), Rowbuffer(Rowstand + (3 * BytesPerPixel)), BytesPerPixel
                    StdBuffer = StdBuffer + (8 * BytesPerPixel)
                    Rowstand = Rowstand + (4 * BytesPerPixel)
                Loop
            Case 7
                CopyMemory Buffer(StdBuffer), Rowbuffer(0), UBound(Rowbuffer) + 1
        End Select
    End With

End Sub

Private Sub Swap4Bytes(Bytes As Long)
'1234 >> 4321
    Dim DstBytes As Long
    
    CopyMemory ByVal VarPtr(DstBytes), ByVal VarPtr(Bytes) + 3, 1
    CopyMemory ByVal VarPtr(DstBytes) + 1, ByVal VarPtr(Bytes) + 2, 1
    CopyMemory ByVal VarPtr(DstBytes) + 2, ByVal VarPtr(Bytes) + 1, 1
    CopyMemory ByVal VarPtr(DstBytes) + 3, ByVal VarPtr(Bytes), 1
    Bytes = DstBytes
    
End Sub

Private Sub SwapRGB()
'RGB >> BGR
    Dim idx  As Long
    Dim Temp As Byte

    For idx = 0 To UBound(mBitmapData) Step 3
        Temp = mBitmapData(idx)
        mBitmapData(idx) = mBitmapData(idx + 2)
        mBitmapData(idx + 2) = Temp
    Next idx

End Sub

Private Sub SwapRGBA()
'RGBA >> BGRA
    Dim idx As Long
    Dim Temp As Byte
    
    For idx = 0 To UBound(mBitmapData) Step 4
        Temp = mBitmapData(idx)
        If idx + 2 > UBound(mBitmapData) Then Exit For
        mBitmapData(idx) = mBitmapData(idx + 2)
        mBitmapData(idx + 2) = Temp
    Next idx

End Sub
