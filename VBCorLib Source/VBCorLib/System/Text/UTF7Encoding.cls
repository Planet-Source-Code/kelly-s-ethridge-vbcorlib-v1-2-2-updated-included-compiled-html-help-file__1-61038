VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UTF7Encoding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) 2004 Kelly Ethridge
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
'    Module: UTF7Encoding
'

''
' A set of functions used to convert unicode characters to and from byte arrays.
'
' @see Constructors
' @see Encoding
'
Option Explicit
Implements cObject
Implements Encoding

Private Const CODE_PAGE As Long = 65000

Private mChars() As Integer
Private mCharsSA As SafeArray1d
Private mDirectlyEncodable() As Boolean
Private mAllowOptionals As Boolean



''
' Returns the encoding name to be used in with the mail agent body tags.
'
' @return The body tag compatible encoding name.
'
Public Property Get BodyName() As String
    BodyName = "utf-7"
End Property

''
' Returns the encoding name registered with the Internet Assigned Numbers Authority.
'
' @return The registered encoding name.
'
Public Property Get WebName() As String
    WebName = "utf-7"
End Property

''
' Returns the encoding name to be used in with the mail agent header tags.
'
' @return The header tag compatible encoding name.
'
Public Property Get HeaderName() As String
    HeaderName = "utf-7"
End Property

''
' Returns the code page identifier for this encoding.
'
' @return Code page identifier.
'
Public Property Get CodePage() As Long
    CodePage = CODE_PAGE
End Property

''
' Returns the Windows Operating System's code page for this encoding.
'
' @return The Windows code page for this encoding.
'
Public Property Get WindowsCodePage() As Long
    WindowsCodePage = 1200
End Property

''
' The name of this encoding.
'
' @return The encoding's name.
'
Public Property Get EncodingName() As String
    EncodingName = "Unicode (UTF-7)"
End Property

''
' Indicates if this encoding can be used by browsers to display text.
'
' @return The indication for browser display capable.
'
Public Property Get IsBrowserDisplay() As Boolean
    IsBrowserDisplay = False
End Property

''
' Indicates if this encoding can be used to save data with this encoding.
'
' @return The indication for the browser saving capable.
'
Public Property Get IsBrowserSave() As Boolean
    IsBrowserSave = False
End Property

''
' Indicates if this encoding can be used to display mail and news by
' mail and news clients.
'
' @return Indication for mail and news client capable.
'
Public Property Get IsMailNewsDisplay() As Boolean
    IsMailNewsDisplay = True
End Property

''
' Indicates if this encoding can be used to save date by mail and news clients.
'
' @return Indication for the mail and news clients to use this encoding for saving.
'
Public Property Get IsMailNewsSave() As Boolean
    IsMailNewsSave = True
End Property

''
' Returns the number of bytes that would be produced from the set of characters using this encoding.
'
' @param Chars A set of characters used to calculated the number of bytes once encoded with this encoding.
' @param vIndex The starting index of the first character to be used.
' @param vCount The number of characters to include in the calculation.
' @return The number of bytes that would be produced by encoding the set of characters.
' @remarks The <i>Chars</i> parameter will accept either a String or an Integer array. The vIndex
' is zero based for both the String and Integer array, not 1 based like Mid$.
'
Public Function GetByteCount(ByRef Chars As Variant, Optional ByRef vIndex As Variant, Optional ByRef vCount As Variant) As Long
    Dim Index As Long
    Dim Count As Long
    AttachChars Chars, vIndex, Index, vCount, Count
    GetByteCount = InternalGetByteCount(mChars, Index, Count, Nothing)
End Function

''
' Encodes a set of characters into an array of bytes.
'
' @param Chars The set of characters to be encoded into an array of bytes. This can
' be either an Integer array or a String.
' @param vIndex The index of the first character to begin encoding from.
' @param vCount The number of characters to be encoded.
' @return A byte array containing the encoded characters as bytes.
' @remarks The <i>Chars</i> parameter will accept either a String or an Integer array. The vIndex
' is zero based for both the String and Integer array, not 1 based like Mid$.
'
Public Function GetBytes(ByRef Chars As Variant, Optional ByRef vIndex As Variant, Optional ByRef vCount As Variant) As Byte()
    Dim Index As Long
    Dim Count As Long
    Dim ret() As Byte
    Dim size As Long
    
    AttachChars Chars, vIndex, Index, vCount, Count
    size = InternalGetByteCount(mChars, Index, Count, Nothing)
    ReDim ret(0 To size - 1)
    InternalGetBytes mChars, Index, Count, ret, 0, Nothing
    GetBytes = ret
End Function

''
' Encodes a set of characters into an array of bytes, returning the number of bytes produced.
'
' @param Chars the set of characters to be encoded into bytes. This can
' be either an Integer array or a String.
' @param CharIndex The index of the first character to begin encoding from.
' @param CharCount The number of characters to be encoded.
' @param Bytes The destination for the encoded characters.
' @param ByteIndex The index of the first byte stored in the destination array.
' @return The number of bytes produce by the encoding.
' @remarks The <i>Bytes</i> array cannot be null and must be large enough to hold the encoded
' set of characters. To obtain the size required, use GetByteCount to find the necessary size.
' Or, use GetMaxByteCount to get the largest size that could be produced. The largest size is not
' the perfect size, only the maximum number of bytes that could ever be produced by a number of
' characters. Resizing of the resultant <i>Bytes</i> array may be necessary.
'
Public Function GetBytesEx(ByRef Chars As Variant, ByVal CharIndex As Long, ByVal CharCount As Long, ByRef Bytes() As Byte, ByVal ByteIndex As Long) As Long
    AttachChars Chars, CharIndex, 0, CharCount, 0
    GetBytesEx = InternalGetBytes(mChars, CharIndex, CharCount, Bytes, ByteIndex, Nothing)
End Function

''
' Decodes a set of bytes into a String.
'
' @param Bytes The set of bytes to be decoded into a string.
' @param vIndex The index of the first byte to be decoded.
' @param vCount The number of bytes to be used in the decoding.
' @return A string containing the decoded set of bytes.
'
Public Function GetString(ByRef Bytes() As Byte, Optional ByRef vIndex As Variant, Optional ByRef vCount As Variant) As String
    Dim ret As String
    Dim size As Long
    Dim Index As Long
    Dim Count As Long
    
    AssignByteIndexAndCount Bytes, vIndex, Index, vCount, Count
    size = GetMaxCharCount(Count)
    ret = SysAllocStringLen(0, size)
    With mCharsSA
        .pvData = StrPtr(ret)
        .lLbound = 0
        .cElements = size
    End With
    SAPtr(mChars) = VarPtr(mCharsSA)
    size = GetCharsEx(Bytes, Index, Count, mChars, 0)
    GetString = left$(ret, size)
End Function

''
' Returns the number of characters that would be produced by decoding a byte array.
'
' @param Bytes The bytes to use when calculating the number of characters.
' @param vIndex The starting index in <i>Bytes</i> to being calculating from.
' @param vCount The number of bytes to be used in the calculation.
' @return The number of characters that would be decoded from the set of bytes.
'
Public Function GetCharCount(ByRef Bytes() As Byte, Optional ByRef vIndex As Variant, Optional ByRef vCount As Variant) As Long
    Dim Index As Long
    Dim Count As Long
    
    AssignByteIndexAndCount Bytes, vIndex, Index, vCount, Count
    GetCharCount = InternalGetCharCount(Bytes, Index, Count)
End Function

''
' Decodes a set of bytes into a set of characters.
'
' @param Bytes The set of bytes to be decoded into characters.
' @param vIndex The index of the first byte to be used in decoding.
' @param vCount The number of bytes to be used in decoding.
' @return An array of characters decoded from the set of bytes.
'
Public Function GetChars(ByRef Bytes() As Byte, Optional ByRef vIndex As Variant, Optional ByRef vCount As Variant) As Integer()
    Dim Index As Long
    Dim Count As Long
    Dim ret() As Integer
    Dim size As Long
    
    AssignByteIndexAndCount Bytes, vIndex, Index, vCount, Count
    
    ReDim ret(0 To GetMaxCharCount(Count))
    size = InternalGetChars(Bytes, Index, Count, ret, 0, Nothing)
    ReDim Preserve ret(0 To size - 1)
    GetChars = ret
End Function

''
' Decodes a set of bytes into the supplied Integer array.
'
' @param Bytes The set of bytes to be decoded into characters.
' @param ByteIndex The index of the first byte to begin decoding from.
' @param ByteCount The number of bytes to be used in decoding.
' @param Chars The destination character array for the decoded bytes.
' @param CharIndex The first index to begin storing decoded bytes in <i>Chars</i>.
' @return The number of characters decoded from the array of bytes.
' @remarks The <i>Chars</i> array must be large enough to handle all the bytes that will
' be decoded. To ensure the <i>Chars</i> array is large enough, use either GetCharCount or
' GetMaxCharCount to determine a size that will hold the decoded bytes.
'
Public Function GetCharsEx(ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, ByRef Chars() As Integer, ByVal CharIndex As Long) As Long
    GetCharsEx = InternalGetChars(Bytes, ByteIndex, ByteCount, Chars, CharIndex, Nothing)
End Function

''
' Returns the maximum number of characters than can be decoded from the number of bytes specified.
'
' @param ByteCount The number of bytes to be decoded.
' @return The maximum number of characters that can be decoded from the specified number of bytes.
'
Public Function GetMaxCharCount(ByVal ByteCount As Long) As Long
    If ByteCount < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "ByteCount", ByteCount)
        
    GetMaxCharCount = ByteCount
End Function

''
' Returns the maximum number of bytes that can be created from a specific number of characters.
'
' @param CharCount The number of characters to be encoded.
' @return The maximum number of bytes that can be generated from the specified number of characters.
'
Public Function GetMaxByteCount(ByVal CharCount As Long) As Long
    If CharCount < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "CharCount", CharCount)
    
    GetMaxByteCount = 2 + (-(Int(-((CharCount * 16) / 6))))
End Function

''
' Returns an array of bytes that represents this encoding.
'
' @return A byte array containg the bytes to be used to identify this encoding type.
' @remarks <p>There is no identifier for UTF7 so an empty array is returned.</p>
'
Public Function GetPreamble() As Byte()
    GetPreamble = cArray.CreateInstance(ciByte)
End Function

''
' Returns an encoder that maintains state.
'
' @return The stateful encoder.
' @remarks Unlike UTF7Encoding, the encoder maintains state between encoding calls. This allows
' for a set of characters to be split up but still be encoded as a single set of characters. This
' will allow Modified Base64 encoding to continue between encoding calls.
'
Public Function GetEncoder() As Encoder
    Dim ret As New UTF7Encoder
    ret.Init Me
    Set GetEncoder = ret
End Function

''
' Returns a decoder that maintains state.
'
' @return The stateful decoder.
' @remarks Unlike UTF8Encoding, the decoder maintains state between decoding calls. this allows
' for a sequence of bytes to not have to be in the same call in order to be decoded to the
' representive character. All characters other than directly encodable characters are encoded
' in a Modified Base64 encoding, which can be require 3 bytes to decode to a single character.
' This decoder allows the Base64 encoded byte set to be decoded across multiple calls and still
' be decoded as a single set of bytes. A multi-byte encoded characters does not need to be in
' the same call in order to be decoded properly.
'
Public Function GetDecoder() As Decoder
    Dim ret As New UTF7Decoder
    ret.Init Me
    Set GetDecoder = ret
End Function

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Public Function ToString() As String
    ToString = App.Title & "." & TypeName(Me)
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equality to.
' @return Boolean indicating equality.
Public Function Equals(ByRef Value As Variant) As Boolean
    If IsObject(Value) Then
        Equals = (Value Is Me)
    End If
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long
    GetHashCode = ObjPtr(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal AllowOptionals As Boolean)
    Dim i As Long
    
    If AllowOptionals Then
        For i = 1 To 20
            mDirectlyEncodable(Choose(i, 33, 34, 35, 36, 37, 38, 42, 59, 60, 61, 62, 64, 91, 93, 94, 95, 96, 123, 124, 125)) = True
        Next i
    End If
    mAllowOptionals = AllowOptionals
End Sub

Friend Function InternalGetBytes(ByRef Chars() As Integer, ByVal Index As Long, ByVal Count As Long, ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByRef Encoder As UTF7Encoder) As Long
    Dim bits As Long
    Dim bitCount As Long
    Dim maxCharIndex As Long
    Dim maxByteIndex As Long
    Dim ch As Long
    Dim ByteStartIndex As Long
    Dim directlyEncodable As Boolean
    
    modPublicFunctions.Encoding.GetBytesArgumentCheck Chars, Index, Count, Bytes, ByteIndex
    
    If Not Encoder Is Nothing Then
        With Encoder
            bits = .bits
            bitCount = .bitCount
        End With
    Else
        bitCount = -1
    End If
    
    ByteStartIndex = ByteIndex
    maxCharIndex = Index + Count - 1
    maxByteIndex = UBound(Bytes)
    Do While Index <= maxCharIndex
        AsWord(ch) = Chars(Index)
        directlyEncodable = False
        If ch < &H80 Then
            If mDirectlyEncodable(ch) Then
                If bitCount >= 0 Then
                    If bitCount > 0 Then
                        If ByteIndex > maxByteIndex Then SmallBufferError "Bytes"
                        Bytes(ByteIndex) = Base64Bytes(Helper.ShiftLeft(bits, 6 - bitCount) And &H3F)
                        ByteIndex = ByteIndex + 1
                    End If
                    If ByteIndex > maxByteIndex Then SmallBufferError "Bytes"
                    Bytes(ByteIndex) = CHAR_MINUS_SIGN
                    ByteIndex = ByteIndex + 1
                End If
                If ByteIndex > maxByteIndex Then SmallBufferError "Bytes"
                Bytes(ByteIndex) = ch
                ByteIndex = ByteIndex + 1
                bitCount = -1
                directlyEncodable = True
            End If
        End If
        If Not directlyEncodable Then
            If ch = CHAR_PLUS_SIGN And bitCount < 0 Then
                If ByteIndex > maxByteIndex Then SmallBufferError "Bytes"
                Bytes(ByteIndex) = CHAR_PLUS_SIGN
                Bytes(ByteIndex + 1) = CHAR_MINUS_SIGN
                ByteIndex = ByteIndex + 2
            Else
                If bitCount < 0 Then
                    If ByteIndex > maxByteIndex Then SmallBufferError "Bytes"
                    Bytes(ByteIndex) = CHAR_PLUS_SIGN
                    ByteIndex = ByteIndex + 1
                    bitCount = 0
                End If
                bits = (bits * &H10000) Or ch
                bitCount = bitCount + 16
                Do While bitCount >= 6
                    bitCount = bitCount - 6
                    If ByteIndex > maxByteIndex Then SmallBufferError "Bytes"
                    Bytes(ByteIndex) = Base64Bytes(Helper.ShiftRight(bits, bitCount) And &H3F)
                    ByteIndex = ByteIndex + 1
                Loop
                bits = bits And (Powers(bitCount) - 1)
            End If
        End If
        Index = Index + 1
    Loop
    
    Dim Flush As Boolean
    If Encoder Is Nothing Then
        Flush = True
    Else
        Flush = Encoder.Flush
    End If
    
    If Flush And bitCount >= 0 Then
        If bitCount > 0 Then
            If ByteIndex > maxByteIndex Then _
                Throw Cor.NewArgumentException("Bytes conversion buffer is too small.", "Bytes")
            
            Bytes(ByteIndex) = Base64Bytes(Helper.ShiftLeft(bits, 6 - bitCount) And &H3F)
            ByteIndex = ByteIndex + 1
        End If
        If ByteIndex > maxByteIndex Then _
            Throw Cor.NewArgumentException("Bytes conversion buffer is too small.", "Bytes")
        
        Bytes(ByteIndex) = CHAR_MINUS_SIGN
        ByteIndex = ByteIndex + 1
        bitCount = -1
    End If
    
    If Not Encoder Is Nothing Then
        With Encoder
            .bits = bits
            .bitCount = bitCount
        End With
    End If
    
    InternalGetBytes = ByteIndex - ByteStartIndex
End Function

Friend Function InternalGetByteCount(ByRef Chars() As Integer, ByVal Index As Long, ByVal Count As Long, ByRef Encoder As UTF7Encoder) As Long
    Dim bitCount As Long
    Dim maxCharIndex As Long
    Dim ch As Long
    Dim directlyEncodable As Boolean
    Dim ByteIndex As Long
    
    If cArray.IsNull(Chars) Then _
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Array), "Chars")
    If Index < LBound(Chars) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "Index", Index)
    If Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "Count", Count)
    If Index + Count - 1 > UBound(Chars) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "Count")
    
    If Not Encoder Is Nothing Then
        bitCount = Encoder.bitCount
    Else
        bitCount = -1
    End If
    
    maxCharIndex = Index + Count - 1
    Do While Index <= maxCharIndex
        AsWord(ch) = Chars(Index)
        directlyEncodable = False
        If ch < &H80 Then
            If mDirectlyEncodable(ch) Then
                If bitCount >= 0 Then
                    If bitCount > 0 Then ByteIndex = ByteIndex + 1
                    ByteIndex = ByteIndex + 1
                End If
                ByteIndex = ByteIndex + 1
                bitCount = -1
                directlyEncodable = True
            End If
        End If
        If Not directlyEncodable Then
            If ch = CHAR_PLUS_SIGN And bitCount < 0 Then
                ByteIndex = ByteIndex + 2
            Else
                If bitCount < 0 Then
                    ByteIndex = ByteIndex + 1
                    bitCount = 0
                End If
                bitCount = bitCount + 16
                Do While bitCount >= 6
                    bitCount = bitCount - 6
                    ByteIndex = ByteIndex + 1
                Loop
            End If
        End If
        Index = Index + 1
    Loop
    
    Dim Flush As Boolean
    If Encoder Is Nothing Then
        Flush = True
    Else
        Flush = Encoder.Flush
    End If
    
    If Flush And bitCount >= 0 Then
        If bitCount > 0 Then
            ByteIndex = ByteIndex + 1
        End If
        ByteIndex = ByteIndex + 1
        bitCount = -1
    End If
    
    InternalGetByteCount = ByteIndex
End Function

Friend Function InternalGetCharCount(ByRef Bytes() As Byte, ByVal Index As Long, ByVal Count As Long) As Long
    Dim InBase64 As Boolean
    Dim maxindex As Long
    Dim CharCount As Long
    Dim b As Byte
    Dim bits As Currency
    
    If Index < LBound(Bytes) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "vIndex", Index)
    If Index + Count - 1 > UBound(Bytes) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "vCount")
    
    maxindex = Index + Count
    Do While Index < maxindex
        b = Bytes(Index)
        If Not InBase64 Then
            If b = CHAR_PLUS_SIGN Then
                InBase64 = True
            Else
                CharCount = CharCount + 1
            End If
        ElseIf b = CHAR_MINUS_SIGN Then
            If Bytes(Index - 1) = CHAR_PLUS_SIGN Then CharCount = CharCount + 1
            CharCount = CharCount + Int(bits / 16@)
            bits = 0@
            InBase64 = False
        Else
            bits = bits + 6@
        End If
        Index = Index + 1
    Loop
    
    InternalGetCharCount = Int(bits / 16@) + CharCount
End Function

Friend Function InternalGetChars(ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, ByRef Chars() As Integer, ByVal CharIndex As Long, ByRef Decoder As UTF7Decoder) As Long
    Dim bits As Long
    Dim bitCount As Long
    Dim b As Byte
    Dim InBase64 As Boolean
    Dim HasPlus As Boolean
    Dim CharStartIndex As Long
    Dim maxCharIndex As Long
    Dim maxByteIndex As Long
    
    modPublicFunctions.Encoding.GetCharsArgumentCheck Bytes, ByteIndex, ByteCount, Chars, CharIndex
    
    If Not Decoder Is Nothing Then
        With Decoder
            bits = .bits
            bitCount = .bitCount
            InBase64 = .InBase64
            HasPlus = .HasPlus
        End With
    End If
    
    CharStartIndex = CharIndex
    maxCharIndex = UBound(Chars)
    maxByteIndex = ByteIndex + ByteCount - 1
    Do While ByteIndex <= maxByteIndex
        b = Bytes(ByteIndex)
        Select Case b
            Case CHAR_MINUS_SIGN
                If HasPlus Then
                    If CharIndex > maxCharIndex Then SmallBufferError "Chars"
                    Chars(CharIndex) = CHAR_PLUS_SIGN
                    CharIndex = CharIndex + 1
                ElseIf InBase64 Then
                    bitCount = 0
                    bits = 0
                    InBase64 = False
                Else
                    If CharIndex > maxCharIndex Then SmallBufferError "Chars"
                    Chars(CharIndex) = CHAR_MINUS_SIGN
                    CharIndex = CharIndex + 1
                End If
                HasPlus = False
            Case CHAR_PLUS_SIGN
                If InBase64 Then
                    bits = (bits * &H40) Or Base64CharToBits(b)
                    bitCount = bitCount + 6
                    HasPlus = False
                Else
                    bits = 0
                    bitCount = 0
                    InBase64 = True
                    HasPlus = True
                End If
            Case Is < &H80
                If InBase64 Then
                    If Base64CharToBits(b) >= 0 Then
                        bits = (bits * &H40) Or Base64CharToBits(b)
                        bitCount = bitCount + 6
                    Else
                        bitCount = 0
                        bits = 0
                        InBase64 = False
                    End If
                Else
                    If CharIndex > maxCharIndex Then SmallBufferError "Chars"
                    Chars(CharIndex) = b
                    CharIndex = CharIndex + 1
                End If
                HasPlus = False
            Case Else
                If CharIndex > maxCharIndex Then SmallBufferError "Chars"
                Chars(CharIndex) = b
                CharIndex = CharIndex + 1
                bitCount = 0
                bits = 0
                InBase64 = False
                HasPlus = False
        End Select
        If bitCount >= 16 Then
            If CharIndex > maxCharIndex Then SmallBufferError "Chars"
            Chars(CharIndex) = AsWord(Helper.ShiftRight(bits, bitCount - 16) And &HFFFF&)
            CharIndex = CharIndex + 1
            bitCount = bitCount - 16
        End If
        ByteIndex = ByteIndex + 1
        bits = bits And &HFFFFFF
    Loop
    
    If Not Decoder Is Nothing Then
        With Decoder
            .bitCount = bitCount
            .bits = bits
            .HasPlus = HasPlus
            .InBase64 = InBase64
        End With
    End If
    
    InternalGetChars = CharIndex - CharStartIndex
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SmallBufferError(ByVal ParamName As String)
    Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_SmallConversionBuffer), ParamName)
End Sub

Private Sub AttachChars(ByRef Chars As Variant, ByRef vIndex As Variant, ByRef IndexOut As Long, ByRef vCount As Variant, ByRef CountOut As Long)
    Dim lb As Long
    Dim l As Long
    
    Select Case VarType(Chars)
        Case vbString
            l = Len(Chars)
            mCharsSA.cElements = l
            mCharsSA.pvData = StrPtr(Chars)
            mCharsSA.cbElements = 2
            mCharsSA.cDims = 1
            
            SAPtr(mChars) = VarPtr(mCharsSA)
        Case INTEGER_ARRAY
            SAPtr(mChars) = cArray.GetValidArrayPointer(Chars)
            lb = LBound(mChars)
            l = UBound(mChars) - lb + 1
        Case Else
            Throw Cor.NewArgumentException("Chars must be a String or Integer array.", "Chars")
    End Select
    
    If IsMissing(vIndex) Then
        IndexOut = lb
    Else
        IndexOut = vIndex
    End If
    If IsMissing(vCount) Then
        CountOut = l
    Else
        CountOut = vCount
    End If
End Sub

Private Sub AssignByteIndexAndCount(ByRef Bytes() As Byte, ByRef vIndex As Variant, ByRef Index As Long, ByRef vCount As Variant, ByRef Count As Long)
    If cArray.IsNull(Bytes) Then _
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Array), "Bytes")
    
    If IsMissing(vIndex) Then
        Index = LBound(Bytes)
    Else
        Index = vIndex
    End If
    If IsMissing(vCount) Then
        Count = cArray.GetLength(Bytes)
    Else
        Count = vCount
    End If
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Dim i As Long
    
    ReDim mDirectlyEncodable(0 To 127)
    For i = 0 To 25
        mDirectlyEncodable(65 + i) = True
        mDirectlyEncodable(97 + i) = True
    Next i
    For i = 0 To 9
        mDirectlyEncodable(48 + i) = True
    Next i
    For i = 1 To 13
        mDirectlyEncodable(Choose(i, 9, 10, 13, 32, 39, 40, 41, 44, 45, 46, 47, 58, 63)) = True
    Next i
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    Init PropBag.ReadProperty("AllowOptionals", False)
End Sub

Private Sub Class_Terminate()
    SAPtr(mChars) = 0
End Sub


Private Sub Class_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "AllowOptionals", mAllowOptionals
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   cObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function cObject_Equals(Value As Variant) As Boolean
    cObject_Equals = Equals(Value)
End Function

Private Function cObject_GetHashCode() As Long
    cObject_GetHashCode = GetHashCode
End Function

Private Function cObject_ToString() As String
    cObject_ToString = ToString
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Encoding Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get Encoding_BodyName() As String
    Encoding_BodyName = BodyName
End Property

Private Property Get Encoding_CodePage() As Long
    Encoding_CodePage = CodePage
End Property

Private Property Get Encoding_EncodingName() As String
    Encoding_EncodingName = EncodingName
End Property

Private Function Encoding_Equals(Value As Variant) As Boolean
    Encoding_Equals = Equals(Value)
End Function

Private Function Encoding_GetByteCount(Chars As Variant, Optional vIndex As Variant, Optional vCount As Variant) As Long
    Encoding_GetByteCount = GetByteCount(Chars, vIndex, vCount)
End Function

Private Function Encoding_GetBytes(Chars As Variant, Optional vIndex As Variant, Optional vCount As Variant) As Byte()
    Encoding_GetBytes = GetBytes(Chars, vIndex, vCount)
End Function

Private Function Encoding_GetBytesEx(Chars As Variant, ByVal CharIndex As Long, ByVal CharCount As Long, Bytes() As Byte, ByVal ByteIndex As Long) As Long
    Encoding_GetBytesEx = GetBytesEx(Chars, CharIndex, CharCount, Bytes, ByteIndex)
End Function

Private Function Encoding_GetCharCount(Bytes() As Byte, Optional Index As Variant, Optional Count As Variant) As Long
    Encoding_GetCharCount = GetCharCount(Bytes, Index, Count)
End Function

Private Function Encoding_GetChars(Bytes() As Byte, Optional vIndex As Variant, Optional vCount As Variant) As Integer()
    Encoding_GetChars = GetChars(Bytes, vIndex, vCount)
End Function

Private Function Encoding_GetCharsEx(Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, Chars() As Integer, ByVal CharIndex As Long) As Long
    Encoding_GetCharsEx = GetCharsEx(Bytes, ByteIndex, ByteCount, Chars, CharIndex)
End Function

Private Function Encoding_GetDecoder() As Decoder
    Set Encoding_GetDecoder = GetDecoder
End Function

Private Function Encoding_GetEncoder() As Encoder
    Set Encoding_GetEncoder = GetEncoder
End Function

Private Function Encoding_GetHashCode() As Long
    Encoding_GetHashCode = GetHashCode
End Function

Private Function Encoding_GetMaxByteCount(ByVal CharCount As Long) As Long
    Encoding_GetMaxByteCount = GetMaxByteCount(CharCount)
End Function

Private Function Encoding_GetMaxCharCount(ByVal ByteCount As Long) As Long
    Encoding_GetMaxCharCount = GetMaxCharCount(ByteCount)
End Function

Private Function Encoding_GetPreamble() As Byte()
    Encoding_GetPreamble = GetPreamble
End Function

Private Function Encoding_GetString(Bytes() As Byte, Optional vIndex As Variant, Optional vCount As Variant) As String
    Encoding_GetString = GetString(Bytes, vIndex, vCount)
End Function

Private Property Get Encoding_HeaderName() As String
    Encoding_HeaderName = HeaderName
End Property

Private Property Get Encoding_IsBrowserDisplay() As Boolean
    Encoding_IsBrowserDisplay = IsBrowserDisplay
End Property

Private Property Get Encoding_IsBrowserSave() As Boolean
    Encoding_IsBrowserSave = IsBrowserSave
End Property

Private Property Get Encoding_IsMailNewsDisplay() As Boolean
    Encoding_IsMailNewsDisplay = IsMailNewsDisplay
End Property

Private Property Get Encoding_IsMailNewsSave() As Boolean
    Encoding_IsMailNewsSave = IsMailNewsSave
End Property

Private Function Encoding_ToString() As String
    Encoding_ToString = ToString
End Function

Private Property Get Encoding_WebName() As String
    Encoding_WebName = WebName
End Property

Private Property Get Encoding_WindowsCodePage() As Long
    Encoding_WindowsCodePage = WindowsCodePage
End Property
