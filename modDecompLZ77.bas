Attribute VB_Name = "modDecompLZ77"
' Original Code : CodeId=56537 (PNG Class By Alfred Koppold)
' Revision By Erkan Þanlý 2009
Option Explicit

Private Type CODESTYPE
    Lenght()            As Long
    Code()              As Long
End Type

Private Inpos           As Long
Private OutPos          As Long
Private InStream()      As Byte
Private OutStream()     As Byte
Private BitNum          As Long
Private MinLLenght      As Long
Private MaxLLenght      As Long
Private ByteBuff        As Long
Private MinDLenght      As Long
Private MaxDLenght      As Long
Private LitLen          As CODESTYPE
Private LC              As CODESTYPE
Private Dist            As CODESTYPE
Private dc              As CODESTYPE
Private TempLit         As CODESTYPE
Private TempDist        As CODESTYPE

Private LenOrder(18)    As Long
Private BitMask(16)     As Long
Private Pow2(16)        As Long
Private IsStaticBuild   As Boolean

Public Function DecompLZ77(UncompressedSize As Long, Buffer() As Byte, Optional ZIP64 As Boolean = False) As Long

    Dim Temp()      As Variant
    Dim idx         As Long
    Dim IsLastBlock As Boolean
    Dim CompType    As Long
    Dim Char        As Long
    Dim NumBits     As Long
    Dim L1          As Long
    Dim L2          As Long
        
'Init decompression

'ReSet data
    UncompressedSize = UncompressedSize + 100
    ReDim OutStream(UncompressedSize)
    ReDim LC.Code(31)
    ReDim LC.Lenght(31)
    ReDim dc.Code(31)
    ReDim dc.Lenght(31)
    Erase LitLen.Code
    Erase LitLen.Lenght
    Erase Dist.Code
    Erase Dist.Lenght
    Inpos = 0
    OutPos = 0
    ByteBuff = 0
    BitNum = 0

'Len order
    Temp() = Array(16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15)
    For idx = 0 To UBound(Temp): LenOrder(idx) = Temp(idx): Next
'LC code
    Temp() = Array(3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258)
    For idx = 0 To UBound(Temp): LC.Code(idx) = Temp(idx): Next
'LC lenght
    Temp() = Array(0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0)
    For idx = 0 To UBound(Temp): LC.Lenght(idx) = Temp(idx): Next
'dc code
    Temp() = Array(1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577, 32769, 49153)
    For idx = 0 To UBound(Temp): dc.Code(idx) = Temp(idx): Next
'dc lenght
    Temp() = Array(0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13, 14, 14)
    For idx = 0 To UBound(Temp): dc.Lenght(idx) = Temp(idx): Next
'bitmask and pow2
    For idx = 0 To 16: Pow2(idx) = 2 ^ idx: BitMask(idx) = Pow2(idx) - 1: Next

    InStream = Buffer
    Do
        IsLastBlock = GetBits(1)
        CompType = GetBits(2)
        If CompType = 0 Then
            If Inpos + 4 > UBound(InStream) Then
                DecompLZ77 = -1
                Exit Do
            End If
            Do While BitNum >= 8
                Inpos = Inpos - 1
                BitNum = BitNum - 8
            Loop
            CopyMemory L1, InStream(Inpos), 2&
            CopyMemory L2, InStream(Inpos + 2), 2&
            Inpos = Inpos + 4
            If L1 - (Not (L2) And &HFFFF&) Then DecompLZ77 = -2
            If Inpos + L1 - 1 > UBound(InStream) Then
                DecompLZ77 = -1
                Exit Do
            End If
            If OutPos + L1 - 1 > UBound(OutStream) Then
                DecompLZ77 = -1
                Exit Do
            End If
            CopyMemory OutStream(OutPos), InStream(Inpos), L1
            OutPos = OutPos + L1
            Inpos = Inpos + L1
            ByteBuff = 0
            BitNum = 0
        ElseIf CompType = 3 Then
            DecompLZ77 = -1
            Exit Do
        Else
            If CompType = 1 Then
                If CreateStaticTree <> 0 Then
                    MsgBox "Error in tree creation (Static)"
                    Exit Function
                End If
            Else
                If CreateDynamicTree <> 0 Then
                    MsgBox "Error in tree creation (Dynamic)"
                    Exit Function
                End If
            End If
            Do
                NeedBits MaxLLenght
                NumBits = MinLLenght
                Do While LitLen.Lenght(ByteBuff And BitMask(NumBits)) <> NumBits
                    NumBits = NumBits + 1
                Loop
                Char = LitLen.Code(ByteBuff And BitMask(NumBits))
                DropBits NumBits
                If Char < 256 Then
                    OutStream(OutPos) = Char
                    OutPos = OutPos + 1
                ElseIf Char > 256 Then
                    Char = Char - 257
                    L1 = LC.Code(Char) + GetBits(LC.Lenght(Char))
                    If (L1 = 258) And ZIP64 Then L1 = GetBits(16) + 3
                    NeedBits MaxDLenght
                    NumBits = MinDLenght
                    Do While Dist.Lenght(ByteBuff And BitMask(NumBits)) <> NumBits
                        NumBits = NumBits + 1
                    Loop
                    Char = Dist.Code(ByteBuff And BitMask(NumBits))
                    DropBits NumBits
                    L2 = dc.Code(Char) + GetBits(dc.Lenght(Char))
                    For idx = 1 To L1
                        If OutPos > UncompressedSize Then
                            OutPos = UncompressedSize
                            GoTo Stop_Decompression
                        End If
                        OutStream(OutPos) = OutStream(OutPos - L2)
                        OutPos = OutPos + 1
                    Next idx
                End If
            Loop While Char <> 256 'EOB
        End If
    Loop While Not IsLastBlock

Stop_Decompression:

    If OutPos > 0 Then
        ReDim Preserve OutStream(OutPos - 1)
    Else
        Erase OutStream
    End If
    Erase InStream
    Erase BitMask
    Erase Pow2
    Erase LC.Code
    Erase LC.Lenght
    Erase dc.Code
    Erase dc.Lenght
    Erase LitLen.Code
    Erase LitLen.Lenght
    Erase Dist.Code
    Erase Dist.Lenght
    Erase LenOrder
    Buffer = OutStream
    
End Function

Private Sub NeedBits(NumBits As Long)
    
    While BitNum < NumBits
        If Inpos > UBound(InStream) Then Exit Sub
        ByteBuff = ByteBuff + (InStream(Inpos) * Pow2(BitNum))
        BitNum = BitNum + 8
        Inpos = Inpos + 1
    Wend
    
End Sub

Private Sub DropBits(NumBits As Long)
    
    ByteBuff = ByteBuff \ Pow2(NumBits)
    BitNum = BitNum - NumBits

End Sub

Private Function GetBits(NumBits As Long) As Long
    
    While BitNum < NumBits
        ByteBuff = ByteBuff + (InStream(Inpos) * Pow2(BitNum))
        BitNum = BitNum + 8
        Inpos = Inpos + 1
    Wend
    GetBits = ByteBuff And BitMask(NumBits)
    Call DropBits(NumBits)

End Function

Private Function CreateStaticTree()
    
    Dim X           As Long
    Dim Lenght(287) As Long

    If IsStaticBuild = False Then
        For X = 0 To 143: Lenght(X) = 8: Next
        For X = 144 To 255: Lenght(X) = 9: Next
        For X = 256 To 279: Lenght(X) = 7: Next
        For X = 280 To 287: Lenght(X) = 8: Next
        If CreateCodes(TempLit, Lenght, 287, MaxLLenght, MinLLenght) <> 0 Then
            CreateStaticTree = -1
            Exit Function
        End If
        For X = 0 To 31: Lenght(X) = 5: Next
        CreateStaticTree = CreateCodes(TempDist, Lenght, 31, MaxDLenght, MinDLenght)
        IsStaticBuild = True
    Else
        MinLLenght = 7
        MaxLLenght = 9
        MinDLenght = 5
        MaxDLenght = 5
    End If
    LitLen = TempLit
    Dist = TempDist
    
End Function

Private Function CreateDynamicTree() As Long
    
    Dim Lenght()    As Long
    Dim BlTree      As CODESTYPE
    Dim MinBL       As Long
    Dim MaxBL       As Long
    Dim NumLen      As Long
    Dim Numdis      As Long
    Dim NumCod      As Long
    Dim Char        As Long
    Dim NumBits     As Long
    Dim LN          As Long
    Dim Pos         As Long
    Dim X           As Long
    
    ReDim Lenght(18)
    NumLen = GetBits(5) + 257
    Numdis = GetBits(5) + 1
    NumCod = GetBits(4) + 4
    For X = 0 To NumCod - 1: Lenght(LenOrder(X)) = GetBits(3): Next
    For X = NumCod To 18: Lenght(LenOrder(X)) = 0: Next
    If CreateCodes(BlTree, Lenght, 18, MaxBL, MinBL) <> 0 Then
        CreateDynamicTree = -1
        Exit Function
    End If
    ReDim Lenght(NumLen + Numdis)
    Pos = 0
    Do While Pos < NumLen + Numdis
        NeedBits MaxBL
        NumBits = MinBL
        Do While BlTree.Lenght(ByteBuff And BitMask(NumBits)) <> NumBits
            NumBits = NumBits + 1
        Loop
        Char = BlTree.Code(ByteBuff And BitMask(NumBits))
        DropBits NumBits
        If Char < 16 Then
            Lenght(Pos) = Char
            Pos = Pos + 1
        Else
            If Char = 16 Then
                If Pos = 0 Then
                    CreateDynamicTree = -5
                    Exit Function
                End If
                LN = Lenght(Pos - 1)
                Char = 3 + GetBits(2)
            ElseIf Char = 17 Then
                Char = 3 + GetBits(3)
                LN = 0
            Else
                Char = 11 + GetBits(7)
                LN = 0
            End If
            If Pos + Char > NumLen + Numdis Then
                CreateDynamicTree = -6
                Exit Function
            End If
            Do While Char > 0
                Char = Char - 1
                Lenght(Pos) = LN
                Pos = Pos + 1
            Loop
        End If
    Loop
    If CreateCodes(LitLen, Lenght, NumLen - 1, MaxLLenght, MinLLenght) <> 0 Then
        CreateDynamicTree = -1
        Exit Function
    End If
    For X = 0 To Numdis: Lenght(X) = Lenght(X + NumLen): Next
    CreateDynamicTree = CreateCodes(Dist, Lenght, Numdis - 1, MaxDLenght, MinDLenght)

End Function

Private Function CreateCodes(Tree As CODESTYPE, Lenghts() As Long, NumCodes As Long, MaxBits As Long, Minbits As Long) As Long

    Dim Bits(16)        As Long
    Dim NextCode(16)    As Long
    Dim Code            As Long
    Dim LN              As Long
    Dim X               As Long
    
    Minbits = 16
    For X = 0 To NumCodes
        Bits(Lenghts(X)) = Bits(Lenghts(X)) + 1
        If Lenghts(X) > MaxBits Then MaxBits = Lenghts(X)
        If Lenghts(X) < Minbits And Lenghts(X) > 0 Then Minbits = Lenghts(X)
    Next
    LN = 1
    For X = 1 To MaxBits
        LN = LN + LN
        LN = LN - Bits(X)
        If LN < 0 Then CreateCodes = LN: Exit Function
    Next
    CreateCodes = LN
    ReDim Tree.Code(2 ^ MaxBits - 1)
    ReDim Tree.Lenght(2 ^ MaxBits - 1)
    Code = 0
    Bits(0) = 0
    For X = 1 To MaxBits
        Code = (Code + Bits(X - 1)) * 2
        NextCode(X) = Code
    Next
    For X = 0 To NumCodes
        LN = Lenghts(X)
        If LN <> 0 Then
            Code = BitReverse(LN, NextCode(LN))
            Tree.Lenght(Code) = LN
            Tree.Code(Code) = X
            NextCode(LN) = NextCode(LN) + 1
        End If
    Next
    
End Function

Private Function BitReverse(ByVal NumBits As Long, ByVal Value As Long) As Long

    Do While NumBits > 0
        BitReverse = BitReverse * 2 + (Value And 1)
        NumBits = NumBits - 1
        Value = Value \ 2
    Loop

End Function

Public Sub BitsToBytes(NumBits As Byte, Size As Long, SrcBytes() As Byte)

    Dim DstBytes()  As Byte
    Dim Bytes()     As Byte
    Dim MaxLen      As Long
    Dim idx         As Long
    Dim Off         As Long
    
    MaxLen = UBound(SrcBytes) + 1
    Select Case NumBits
        Case 1
            ReDim DstBytes((MaxLen * 8) - 1)
            For idx = 0 To MaxLen - 1
                ByteTo1Bit SrcBytes(idx), Bytes
                CopyMemory DstBytes(Off), Bytes(0), 8
                Off = Off + 8
            Next idx
        Case 2
            ReDim DstBytes((MaxLen * 4) - 1)
            For idx = 0 To MaxLen - 1
                ByteTo2Bit SrcBytes(idx), Bytes
                CopyMemory DstBytes(Off), Bytes(0), 4
                Off = Off + 4
            Next idx
        Case 4
            ReDim DstBytes((MaxLen * 2) - 1)
            For idx = 0 To MaxLen - 1
                ByteTo4Bit SrcBytes(idx), Bytes
                CopyMemory DstBytes(Off), Bytes(0), 2
                Off = Off + 2
            Next idx
    End Select
    ReDim Preserve DstBytes(Size - 1)
    SrcBytes = DstBytes

End Sub

Private Sub ByteTo1Bit(SrcByte As Byte, Bytes() As Byte)
    
    Dim Temp As Byte
    
    ReDim Bytes(7)
    Bytes(7) = SrcByte And 1                        ' 1  = 00000001
    Temp = SrcByte And 2:   Bytes(6) = Temp / 2     ' 2  = 00000010
    Temp = SrcByte And 4:   Bytes(5) = Temp / 4     ' 4  = 00000100
    Temp = SrcByte And 8:   Bytes(4) = Temp / 8     ' 8  = 00001000
    Temp = SrcByte And 16:  Bytes(3) = Temp / 16    ' 16 = 00010000
    Temp = SrcByte And 32:  Bytes(2) = Temp / 32    ' 32 = 00100000
    Temp = SrcByte And 64:  Bytes(1) = Temp / 64    ' 64 = 01000000
    Temp = SrcByte And 128: Bytes(0) = Temp / 128   ' 128= 10000000

End Sub

Private Sub ByteTo2Bit(SrcByte As Byte, Bytes() As Byte)
    
    Dim Temp As Byte
    
    ReDim Bytes(3)
    Bytes(3) = SrcByte And 3                        ' 3  = 00000011
    Temp = SrcByte And 12:  Bytes(2) = Temp / 4     ' 12 = 00001100
    Temp = SrcByte And 48:  Bytes(1) = Temp / 16    ' 48 = 00110000
    Temp = SrcByte And 192: Bytes(0) = Temp / 64    ' 192= 11000000

End Sub

Private Sub ByteTo4Bit(SrcByte As Byte, Bytes() As Byte)
    
    Dim Temp As Byte
    
    ReDim Bytes(1)
    Bytes(1) = SrcByte And 15                       ' 15 = 00001111
    Temp = SrcByte And 240: Bytes(0) = Temp / 16    ' 240= 11110000

End Sub

