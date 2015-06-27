Attribute VB_Name = "modCompresor"
Option Explicit
Private Declare Function qlz_compress Lib "quick32.dll" (ByRef Source As Byte, ByRef Destination As Byte, ByVal length As Long) As Long
Private Declare Function qlz_decompress Lib "quick32.dll" (ByRef Source As Byte, ByRef Destination As Byte) As Long
Private Declare Function qlz_size_decompressed Lib "quick32.dll" (ByRef Source As Byte) As Long
Private Declare Function qlz_size_source Lib "quick32.dll" (ByRef Source As Byte) As Long

Private Count As Byte

' If the Visual Basic IDE cannot find quick32.dll even though it's in the system32 directory, try adding a path
' to the quick32.dll file name in the declarations. This should never be neccessary though.

Function Compress(Source() As Byte) As Byte()
    Dim dst() As Byte
    Dim r As Long
    ReDim dst(0 To UBound(Source) * 1.2 + 36000)
    r = qlz_compress(Source(0), dst(0), UBound(Source) + 1)
    ReDim Preserve dst(0 To r - 1)
    Compress = dst
End Function

Public Function GetSize(Source() As Byte) As Long
    GetSize = qlz_size_decompressed(Source(0))
End Function

Public Function Decompress(Source() As Byte, Optional ByRef blFail As Boolean) As Byte()
    Dim dst() As Byte
    Dim r As Long
    Dim Size As Long
    Size = GetSize(Source)
    If Size = 0 Then blFail = True: Exit Function
    If Size < 20 * 1000000 Then ' Visual Basic can crash if you allocate too long strings
        ReDim dst(0 To Size - 1)
        r = qlz_decompress(Source(0), dst(0))
        ReDim Preserve dst(0 To r - 1)
        Decompress = dst
    End If
End Function




Sub ClearBit(ByRef b As Byte, ByVal bit As Byte)
    ' Create a bitmask with the 2 to the nth power bit set:
    If bit > 7 Then Exit Sub
    Dim Mask As Byte
    Mask = 2 ^ bit
    ' Clear the nth Bit:
    b = b And Not Mask
End Sub

' The ExamineBit function will return True or False depending on
' the value of the nth bit (bit) of an integer (b).
Function ExamineBit(ByVal b As Byte, ByVal bit As Byte) As Boolean
    If bit > 7 Then Exit Function
    ' Create a bitmask with the 2 to the nth power bit set:
    Dim Mask As Byte
    Mask = 2 ^ bit
    ' Return the truth state of the 2 to the nth power bit:
    ExamineBit = ((b And Mask) > 0)
End Function

' The SetBit Sub will set the nth bit (bit) of an integer (b).
Sub SetBit(ByRef b As Byte, ByVal bit As Byte)
    If bit > 7 Then Exit Sub
    ' Create a bitmask with the 2 to the nth power bit set:
    Dim Mask As Byte
    Mask = 2 ^ bit
    ' Set the nth Bit:
    b = b Or Mask
End Sub

' The ToggleBit Sub will change the state of the nth bit (bit)
' of an integer (b).
Sub ToggleBit(ByRef b As Byte, ByVal bit As Byte)
    If bit > 7 Then Exit Sub
    ' Create a bitmask with the 2 to the nth power bit set:
    Dim Mask As Byte
    Mask = 2 ^ bit
    ' Toggle the nth Bit:
    b = b Xor Mask
End Sub



Sub AddByteInfo(ByRef Data() As Byte, ByRef length As Long, ByVal b As Byte, ByVal Count As Byte)
    length = length + 2
    ReDim Preserve Data(length - 1)
    
    Data(length - 2) = b
    Data(length - 1) = Count
End Sub

Sub AnalizeByte(ByRef Data() As Byte, ByRef length As Long, ByVal b As Byte, ByRef PrevB As Byte, ByRef prevL As Long, ByRef first As Boolean)
    Dim i As Byte
    If first Then
        first = False
        PrevB = b
        GoTo ExamineFirst
    End If
    
    If b = PrevB And prevL < 255 Then 'not changed
ExamineFirst:   prevL = prevL + 1
    Else
        AddByteInfo Data, length, PrevB, prevL
        PrevB = b
        prevL = 1
    End If
End Sub

Sub AddByte(ByRef Data() As Byte, ByVal b As Byte, ByRef length As Long, ByVal Count As Byte)

If b = 0 Then
    length = length + 2
    ReDim Preserve Data(length - 1)
    
    Data(length - 2) = b
    Data(length - 1) = Count
Else
    length = length + 1
    ReDim Preserve Data(length - 1)
    
    Data(length - 1) = b
End If

End Sub
Sub AnalizeByte2(ByRef Data() As Byte, ByRef length As Long, ByVal b As Byte, ByRef PrevB As Boolean, ByRef prevL As Long, ByRef first As Boolean)
    Dim i As Byte
    If first Then
        first = False
        If b = 0 Then
            PrevB = True
            prevL = prevL + 1
        Else
            AddByte Data, b, length, 1
            PrevB = False
        End If
    Else
        If PrevB And b = 0 And prevL < 255 Then  'not changed
            prevL = prevL + 1
        ElseIf PrevB And b = 0 And prevL >= 255 Then
            AddByte Data, 0, length, prevL
            PrevB = True
            prevL = 1
        ElseIf PrevB And b <> 0 Then
            AddByte Data, 0, length, prevL
            AddByte Data, b, length, 1
            PrevB = False
            prevL = 0
        ElseIf Not PrevB And b = 0 Then
            PrevB = True
            prevL = 1
        ElseIf Not PrevB And b <> 0 Then
            AddByte Data, b, length, 1
            prevL = 0
        End If
    End If
End Sub
Function CompressData(ByRef Data() As Byte, ByVal method As Byte) As Byte()

Dim length As Long, prevL As Long
Dim PrevB As Byte, first As Boolean
Dim i As Long

If method = 1 Then
    
    
    length = 0
    prevL = 0
    first = True
    PrevB = False
    
    For i = LBound(Data) To UBound(Data)
        Call AnalizeByte(CompressData, length, Data(i), PrevB, prevL, first)
    Next

ElseIf method = 2 Then 'only compress zeros
    
    length = 0
    prevL = 0
    first = True
    Dim PrevB2 As Boolean
    PrevB2 = False
    For i = LBound(Data) To UBound(Data)
        Call AnalizeByte2(CompressData, length, Data(i), PrevB2, prevL, first)
    Next
    
    If prevL > 0 Then
        AddByte CompressData, 0, length, prevL
    End If

ElseIf method = 3 Then
    CompressData = Compress(Data)
End If




Exit Function
    
End Function
Sub UnCompressByte(ByRef UnCompressedData() As Byte, ByRef Count As Long, ByVal Byte1 As Byte, ByVal Byte2 As Byte)
    Dim i As Byte
    i = 0
    While i < Byte2
        UnCompressedData(Count) = Byte1
        Count = Count + 1
        i = i + 1
    Wend

    
End Sub
Function UnCompressByte2(ByRef UnCompressedData() As Byte, ByRef Count As Long, ByVal Byte1 As Byte, ByVal Byte2 As Byte) As Byte
    If Byte1 = 0 Then
        Dim i As Byte
        i = 0
        ReDim Preserve UnCompressedData(Count + Byte2 - 1)
        While i < Byte2
            UnCompressedData(Count) = Byte1
            Count = Count + 1
            i = i + 1
            UnCompressByte2 = 2
        Wend
    Else
        ReDim Preserve UnCompressedData(Count)
        UnCompressedData(Count) = Byte1
        Count = Count + 1
        UnCompressByte2 = 1
    End If

    
End Function

Function UnCompressData(ByRef Data() As Byte, ByVal method As Byte) As Byte()

Dim Count As Long
If method = 1 Then
    
    
    Count = 0
    Dim i As Long
    For i = LBound(Data) To UBound(Data) Step 2
        Call UnCompressByte(UnCompressData, Count, Data(i), Data(i + 1))
    Next
    
ElseIf method = 2 Then
    Count = 0
    Dim j As Long
    j = LBound(Data)
    While j <= UBound(Data)
        Dim step As Byte
        If j = UBound(Data) Then
            step = UnCompressByte2(UnCompressData, Count, Data(j), 1)
        Else
            step = UnCompressByte2(UnCompressData, Count, Data(j), Data(j + 1))
        End If
        j = j + step
    Wend
ElseIf method = 3 Then
    UnCompressData = Decompress(Data)
End If
End Function



