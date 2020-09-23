VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Convert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) 2005 Kelly Ethridge
'
'    This file is part of VBCorLib.
'
'    VBCorLib is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    VBCorLib is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: Convert
'

''
' Provides methods used to encode and decode byte arrays to and from base-64 encoded characters.
'
' @remarks
' To access the methods of this class, simply call the method using the <b>Convert.*</b> syntax.
' <pre>
'''This example takes a byte array containing 10 values
'''from 1 to 10 and encodes it into a Base-64 string encoding.
'''The encoded string is then decoded back into the original
'''byte array and displayed for view.
'
''Private Sub Main()
''    Dim original() As Byte
''    Dim decoded() As Byte
''    Dim s As String
'
''    ' Create a byte array containing values 1 to 10.
''    original = cArray.NewArray(ciByte, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10)
'
''    ' Display the original byte array values
''    Console.WriteLine "Original Bytes"
''    DisplayBytes original
'
''    ' Encode the byte array into a Base-64 Encoded string.
''    s = Convert.ToBase64String(original)
'
''    ' Display Base-64 Encoded string.
''    Console.WriteLine
''    Console.WriteLine "Base-64 Encoded String"
''    Console.WriteLine " " & s
'
''    ' Decode Base-64 string into byte array.
''    decoded = Convert.FromBase64String(s)
'
''    ' Display decoded byte array.
''    Console.WriteLine "Bytes After Decoding"
''    DisplayBytes decoded
'
''    ' Wait for user to press enter
''    Console.ReadLine
''End Sub
'
'''Displays the contents of a byte array.
''Private Sub DisplayBytes(ByRef bytes() As Byte)
''    Dim i As Long
'
''    For i = LBound(bytes) To UBound(bytes)
''        Console.WriteValue " " & bytes(i)
''    Next i
''End Sub
'
'''This code produces the following output.
'
'''original bytes
''' 1 2 3 4 5 6 7 8 9 10
'''Base-64 Encoded String
''' AQIDBAUGBwgJCg==
'''Bytes After Decoding
''' 1 2 3 4 5 6 7 8 9 10
' </pre>
'
Option Explicit

Private mChars As WordBuffer


''
' Converts the value of a subset of a byte array to an equivalent
' subset of a Unicode character string consisting of base-64 characters. Parameters specify
' the subsets as offsets of the input array and the number of elements in the input array to process.
'
' @param Bytes The bytes to be converted to a base-64 character array.
' @param Offset Offset within <i>Bytes</i> to begin encoding to base-64 characters.
' @param Length The number of bytes to encode.
' @param InsertLineBreaks Indicates if a vbCrLf should be inserted every 76 characters to
' conform with MIME line formatting.
' @return A string containg the byte array encoded as base-64 characters.
'
Public Function ToBase64String(ByRef bytes() As Byte, Optional ByRef Offset As Variant, Optional ByRef Length As Variant, Optional ByVal InsertLineBreaks As Boolean = False) As String
    Dim o As Long
    Dim l As Long
    Dim Size As Long
    
    AssignByteOffsetAndLength bytes, Offset, o, Length, l
    Size = l / 3 * 4
    If (Size Mod 4) <> 0 Then Size = Size + (4 - (Size Mod 4))
    If InsertLineBreaks Then
        Size = Size + ((Size - 1) \ 76) * 2
    End If
    ToBase64String = SysAllocStringLen(0, Size)
    mChars.SA.pvData = StrPtr(ToBase64String)
    mChars.SA.cElements = Len(ToBase64String)
    ToBase64CharArray bytes, o, l, mChars.Data, 0, InsertLineBreaks
End Function

''
' Converts the value of a subset of a byte array to an equivalent
' subset of a Unicode character array consisting of base 64 digits. Parameters specify
' the subsets as offsets of the input and output arrays and the number of elements in the input array.
'
' @param InArray The bytes to be converted to a base-64 character array.
' @param OffsetIn Offset within <i>InArray</i> to begin encoding to base-64 characters.
' @param Length The number of bytes to encode.
' @param OutArray The output character array containing the base-64 characters.
' @param OffsetOut The starting index in <i>OutArray</i> to place base-64 characters.
' @param InsertLineBreaks Indicates if a vbCrLf should be inserted every 76 characters to
' conform with MIME line formatting.
' @return The number of base-64 characters created.
'
Public Function ToBase64CharArray(ByRef InArray() As Byte, ByVal OffsetIn As Long, ByVal Length As Long, ByRef OutArray() As Integer, ByVal OffsetOut As Long, Optional ByVal InsertLineBreaks As Boolean = False) As Long
    Dim i As Long
    Dim Size As Long
    Dim CharCapacity As Long
    Dim CharCount As Long
    
    If cArray.IsNull(InArray) Then _
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Array), "InArray")
    If OffsetIn < LBound(InArray) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "OffsetIn", OffsetIn)
    If Length < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "Length", Length)
    If OffsetIn + Length - 1 > UBound(InArray) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "Length")
    If cArray.IsNull(OutArray) Then _
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Array), "OutArray")
    If OffsetOut < LBound(OutArray) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "OffsetOut")
    
    Size = OffsetIn + (Length - (Length Mod 3))
    CharCapacity = Length / 3 * 4
    If (CharCapacity Mod 4) <> 0 Then CharCapacity = CharCapacity + (4 - (CharCapacity Mod 4))
    If InsertLineBreaks Then
        CharCapacity = CharCapacity + ((CharCapacity - 1) \ 76) * 2
    End If
    
    If OffsetOut + CharCapacity - 1 > UBound(OutArray) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_SmallConversionBuffer), "OutArray")
    
    For i = OffsetIn To Size - 1 Step 3
        If InsertLineBreaks And (CharCount = 76) Then
            OutArray(OffsetOut) = 13
            OutArray(OffsetOut + 1) = 10
            OffsetOut = OffsetOut + 2
            CharCount = 0
        End If
        
        OutArray(OffsetOut) = Base64Bytes((InArray(i) And &HFC) \ 4)
        OutArray(OffsetOut + 1) = Base64Bytes(((InArray(i) And &H3) * &H10) Or ((InArray(i + 1) And &HF0) \ &H10))
        OutArray(OffsetOut + 2) = Base64Bytes(((InArray(i + 1) And &HF) * 4) Or ((InArray(i + 2) And &HC0) \ &H40))
        OutArray(OffsetOut + 3) = Base64Bytes(InArray(i + 2) And &H3F)
        OffsetOut = OffsetOut + 4
        If InsertLineBreaks Then CharCount = CharCount + 4
    Next i
    
    If InsertLineBreaks Then
        Select Case Length Mod 3
            Case 1, 2
                If CharCount = 76 Then
                    OutArray(OffsetOut) = 13
                    OutArray(OffsetOut + 1) = 10
                    OffsetOut = OffsetOut + 2
                End If
        End Select
    End If
    
    Select Case Length Mod 3
        Case 1
            OutArray(OffsetOut) = Base64Bytes((InArray(Size) And &HFC) \ 4)
            OutArray(OffsetOut + 1) = Base64Bytes((InArray(Size) And &H3) * &H10)
            OutArray(OffsetOut + 2) = CHAR_EQUAL
            OutArray(OffsetOut + 3) = CHAR_EQUAL
            OffsetOut = OffsetOut + 4
        Case 2
            OutArray(OffsetOut) = Base64Bytes((InArray(Size) And &HFC) \ 4)
            OutArray(OffsetOut + 1) = Base64Bytes(((InArray(Size) And &H3) * &H10) Or ((InArray(Size + 1) And &HF0) \ &H10))
            OutArray(OffsetOut + 2) = Base64Bytes((InArray(Size + 1) And &HF) * 4)
            OutArray(OffsetOut + 3) = CHAR_EQUAL
            OffsetOut = OffsetOut + 4
    End Select
    ToBase64CharArray = OffsetOut
End Function

''
' Decodes a base-64 character array to a byte array. The Offset and length specify a subset of
' the character array to be decoded.
'
' @param InArray A base-64 character array to be decoded to a byte array.
' @param Offset The character within the character array to begin decoding.
' @param Length The number of characters to be decoded into the byte array.
' @return A byte array that has been decoded from a base-64 character array.
'
Public Function FromBase64CharArray(ByRef InArray() As Integer, ByVal Offset As Long, ByVal Length As Long) As Byte()
    Dim i As Long
    Dim j As Long
    Dim ret() As Byte
    Dim bits As Long
    Dim bitCount As Long
    Dim ch As Integer
    Dim CharCount As Long
    Dim EqualCount As Long
    
    If cArray.IsNull(InArray) Then _
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Array), "InArray")
    If Offset < LBound(InArray) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "Offset", Offset)
    If Length < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "Length", Length)
    If Offset + Length - 1 > UBound(InArray) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "Length")
    
    If cArray.GetLength(InArray) = 0 Then
        ret = cArray.CreateInstance(ciByte)
    Else
        ReDim ret(0 To Length - 1)
        For i = Offset To Offset + Length - 1
            ch = InArray(i)
            Select Case ch
                Case CHAR_EQUAL
                    CharCount = CharCount + 1
                    EqualCount = EqualCount + 1
                    
                Case 32, 13, 10, 9
                    If EqualCount > 0 Then _
                        Throw Cor.NewFormatException("An invalid character was found in the Base-64 string.")
                        
                Case LOWER_A_CHAR To LOWER_Z_CHAR, UPPER_A_CHAR To UPPER_Z_CHAR, CHAR_0 To CHAR_9, CHAR_PLUS_SIGN, CHAR_FORSLASH
                    If EqualCount > 0 Then _
                        Throw Cor.NewFormatException("An invalid character was foundin the Base-64 string.")
                    
                    bits = (bits * &H40) Or Base64CharToBits(InArray(i))
                    bitCount = bitCount + 6
                    If bitCount >= 8 Then
                        ret(j) = Helper.ShiftRight(bits, bitCount - 8) And &HFF
                        j = j + 1
                        bitCount = bitCount - 8
                        bits = bits And (Powers(bitCount) - 1)
                    End If
                    CharCount = CharCount + 1
                
                Case Else
                    Throw Cor.NewFormatException("An invalid character was foundin the Base-64 string.")
            End Select
            If EqualCount > 2 Then _
                Throw Cor.NewFormatException("An invalid character was foundin the Base-64 string.")
        Next i
        
        If (CharCount Mod 4) <> 0 Then _
            Throw Cor.NewFormatException("An invalid number of characters were found in the base-64 string.")
            
        Do While bitCount >= 8
            ret(j) = Helper.ShiftRight(bits, bitCount - 8) And &HFF
            j = j + 1
            bitCount = bitCount - 8
            bits = bits And (Powers(bitCount) - 1)
        Loop
        
        If j - 1 <> UBound(ret) Then ReDim Preserve ret(0 To j - 1)
    End If
    
    FromBase64CharArray = ret
End Function

''
' A base-64 string containing characters to be decoded to a byte array.
'
' @param s The base-64 string to be decoded.
' @return A byte array containing the decoded base-64 characters.
'
Public Function FromBase64String(ByRef s As String) As Byte()
    mChars.SA.pvData = StrPtr(s)
    mChars.SA.cElements = Len(s)
    FromBase64String = FromBase64CharArray(mChars.Data, 0, Len(s))
End Function

''
' Converts a datatype value to a string representation using any
' supplied formatting or provider arguments.
'
' @param Value The value to convert to a string.
' @param Format Formatting information for converting the value.
' @param Provider A formatting provider to help custom formatting.
' @return A string representation of the value.
'
Public Function ToString(ByRef Value As Variant, Optional ByVal Format As String, Optional ByVal provider As IFormatProvider) As String
    Dim nf As NumberFormatInfo
    Dim df As DateTimeFormatInfo
    
    Select Case VarType(Value)
        Case vbLong, vbInteger, vbByte, vbDouble, vbSingle
            If Not provider Is Nothing Then Set nf = provider.GetFormat("numberformatinfo")
            If nf Is Nothing Then Set nf = NumberFormatInfo.CurrentInfo
            ToString = nf.Format(Value, Format)
        Case vbDate
            If Not provider Is Nothing Then Set df = provider.GetFormat("datetimeformatinfo")
            If df Is Nothing Then Set df = DateTimeFormatInfo.CurrentInfo
            ToString = df.Format(Value, Format)
        Case vbObject
            If TypeOf Value Is IFormattable Then
                Dim f As IFormattable
                Set f = Value
                ToString = f.ToString(Format, provider)
            ElseIf TypeOf Value Is cObject Then
                Dim o As cObject
                Set o = Value
                ToString = o.ToString
            Else
                Throw Cor.NewArgumentException("Object must implement either IFormattable or cObject interface.", "value")
            End If
        Case Else
            ToString = Value
    End Select
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AssignByteOffsetAndLength(ByRef bytes() As Byte, ByRef vOffset As Variant, ByRef Offset As Long, ByRef vLength As Variant, ByRef Length As Long)
    If cArray.IsNull(bytes) Then _
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Array), "Bytes")
    
    If IsMissing(vOffset) Then
        Offset = LBound(bytes)
    Else
        Offset = vOffset
    End If
    If IsMissing(vLength) Then
        Length = cArray.GetLength(bytes)
    Else
        Length = vLength
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    InitWordBuffer mChars, 0, &H7FFFFFFF
End Sub
