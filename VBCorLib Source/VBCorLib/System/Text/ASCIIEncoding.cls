VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ASCIIEncoding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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
'    Module: ASCIIEncoding
'
''
' Provides a set of methods used to encode and decode chars to and from bytes in ASCII format.
'
' @remarks
' Encoding characters to bytes using ASCII encoding will produce a 1:1 ratio for each
' character-to-byte encoding. If the character is greater than 127, then it will be
' replaced with character 63 (?).
'
' @see Encoding
' @see UTF8Encoding
' @see UTF7Encoding
' @see UnicodeEncoding
'
Option Explicit
Implements cObject
Implements Encoding

Private Const CODE_PAGE As Long = 20127
Private Const WINDOWS_CODE_PAGE As Long = 1252
Private Const BYTE_QUESTION_MARK As Byte = 63

Private mChars() As Integer
Private mCharsSA As SafeArray1d



''
' Returns the encoding name to be used in with the mail agent body tags.
'
' @return The body tag compatible encoding name.
'
Public Property Get BodyName() As String
    BodyName = "us-ascii"
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
' The name of this encoding.
'
' @return The encoding's name.
'
Public Property Get EncodingName() As String
    EncodingName = "US-ASCII"
End Property

''
' Returns the encoding name to be used in with the mail agent header tags.
'
' @return The header tag compatible encoding name.
'
Public Property Get HeaderName() As String
    HeaderName = "us-ascii"
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
' Returns the encoding name registered with the Internet Assigned Numbers Authority.
'
' @return The registered encoding name.
'
Public Property Get WebName() As String
    WebName = "us-ascii"
End Property

''
' Returns the Windows Operating System's code page for this encoding.
'
' @return The Windows code page for this encoding.
'
Public Property Get WindowsCodePage() As Long
    WindowsCodePage = WINDOWS_CODE_PAGE
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
    If Index + Count - 1 > UBound(mChars) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "vCount")
    
    GetByteCount = Count
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
    Dim ret() As Byte
    Dim Index As Long
    Dim Count As Long
    
    AttachChars Chars, vIndex, Index, vCount, Count
    ReDim ret(0 To GetByteCount(Chars, Index, Count) - 1)
    InternalGetBytes mChars, Index, Count, ret, 0
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
Public Function GetBytesEx(ByRef Chars As Variant, ByVal CharIndex As Long, ByVal CharCount As Long, ByRef bytes() As Byte, ByVal ByteIndex As Long) As Long
    AttachChars Chars, CharIndex, 0, CharCount, 0
    GetBytesEx = InternalGetBytes(mChars, CharIndex, CharCount, bytes, ByteIndex)
End Function

''
' Returns the number of characters that would be produced by decoding a byte array.
'
' @param Bytes The bytes to use when calculating the number of characters.
' @param vIndex The starting index in <i>Bytes</i> to being calculating from.
' @param vCount The number of bytes to be used in the calculation.
' @return The number of characters that would be decoded from the set of bytes.
'
Public Function GetCharCount(ByRef bytes() As Byte, Optional ByRef vIndex As Variant, Optional ByRef vCount As Variant) As Long
    Dim Index As Long
    Dim Count As Long
    
    AssignByteIndexAndCount bytes, vIndex, Index, vCount, Count
    If Index + Count - 1 > UBound(bytes) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "Count")
    
    GetCharCount = Count
End Function

''
' Decodes a set of bytes into a set of characters.
'
' @param Bytes The set of bytes to be decoded into characters.
' @param vIndex The index of the first byte to be used in decoding.
' @param vCount The number of bytes to be used in decoding.
' @return An array of characters decoded from the set of bytes.
'
Public Function GetChars(ByRef bytes() As Byte, Optional ByRef vIndex As Variant, Optional ByRef vCount As Variant) As Integer()
    Dim Index As Long
    Dim Count As Long
    Dim ret() As Integer
    
    AssignByteIndexAndCount bytes, vIndex, Index, vCount, Count
    ReDim ret(0 To GetCharCount(bytes, Index, Count) - 1)
    InternalGetChars bytes, Index, Count, ret, 0
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
Public Function GetCharsEx(ByRef bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, ByRef Chars() As Integer, ByVal CharIndex As Long) As Long
    GetCharsEx = InternalGetChars(bytes, ByteIndex, ByteCount, Chars, CharIndex)
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
    
    GetMaxByteCount = CharCount
End Function

''
' Returns an encoder that maintains state.
'
' @return The stateful encoder.
' @remarks Since encoding Unicode characters equates to exactly 2 bytes per
' character, there is no state that is maintained between calls. This method
' functions identically to GetBytes.
'
Public Function GetEncoder() As Encoder
    Dim ret As New DefaultEncoder
    ret.Init Me
    Set GetEncoder = ret
End Function

''
' Returns a decoder that maintains state.
'
' @return The stateful decoder.
' @remarks Unlike UnicodeEncoding, the decoder maintains state between decoding calls.
' Since a single unicode character required exactly 2 bytes to be decoded, it is possible
' that an odd number of bytes may be attempted to be decoding, creating an orphaned byte
' which represents only half of a unicode character. The remaining byte is held until
' the next decoding call, and the byte is then inserted at the beginning of the next set
' of bytes to be decoded, picking up byte decoding where the orphaned byte left off.
'
Public Function GetDecoder() As Decoder
    Dim ret As New DefaultDecoder
    ret.Init Me
    Set GetDecoder = ret
End Function

''
' Returns an array of bytes that represents this encoding.
'
' @return A byte array containg the bytes to be used to identify this encoding type.
' @remarks <p>If the encoding was created setting the <i>BigEndian</i> paramter
' to true, then this will return the Unicode byte array identifer &HFEFF otherwise
' &hFFFE for a non-bigEndian encoding. If no identifier is to be emitted, then an
' empty byte array is returned (not a null array).
'
Public Function GetPreamble() As Byte()
    GetPreamble = cArray.CreateInstance(ciByte)
End Function

''
' Decodes a set of bytes into a String.
'
' @param Bytes The set of bytes to be decoded into a string.
' @param vIndex The index of the first byte to be decoded.
' @param vCount The number of bytes to be used in the decoding.
' @return A string containing the decoded set of bytes.
'
Public Function GetString(ByRef bytes() As Byte, Optional ByRef vIndex As Variant, Optional ByRef vCount As Variant) As String
    Dim Size As Long
    Dim Index As Long
    Dim Count As Long
    
    AssignByteIndexAndCount bytes, vIndex, Index, vCount, Count
    Size = GetCharCount(bytes, Index, Count)
    GetString = SysAllocStringLen(0, Size)
    With mCharsSA
        .pvData = StrPtr(GetString)
        .lLbound = 0
        .cElements = Size
    End With
    SAPtr(mChars) = VarPtr(mCharsSA)
    InternalGetChars bytes, Index, Count, mChars, 0
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
Friend Function InternalGetBytes(ByRef Chars() As Integer, ByVal Index As Long, ByVal Count As Long, ByRef bytes() As Byte, ByVal ByteIndex As Long) As Long
    Dim i As Long
    Dim ch As Integer
    
    modPublicFunctions.Encoding.GetBytesArgumentCheck Chars, Index, Count, bytes, ByteIndex
    If ByteIndex + Count - 1 > UBound(bytes) Then SmallBufferError "Bytes"
    
    For i = 0 To Count - 1
        ch = Chars(Index + i)
        Select Case ch
            ' we need to have a range because by using Integers to
            ' represent unicode characters, we could end up with
            ' negative numbers.
            Case 0 To 127:  bytes(ByteIndex + i) = ch
            Case Else:      bytes(ByteIndex + i) = BYTE_QUESTION_MARK
        End Select
    Next i
    InternalGetBytes = Count
End Function

Friend Function InternalGetChars(ByRef bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, ByRef Chars() As Integer, ByVal CharIndex As Long) As Long
    Dim i As Long
    Dim b As Byte
    
    modPublicFunctions.Encoding.GetCharsArgumentCheck bytes, ByteIndex, ByteCount, Chars, CharIndex
    If CharIndex + ByteCount - 1 > UBound(Chars) Then SmallBufferError "Chars"
    
    For i = 0 To ByteCount - 1
        b = bytes(ByteIndex + i)
        If b < 128 Then
            Chars(CharIndex + i) = b
        Else
            Chars(CharIndex + i) = BYTE_QUESTION_MARK
        End If
    Next i
    InternalGetChars = ByteCount
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
            Throw Cor.NewArgumentException("Chars must be a String or Integer array, not " & TypeName(Chars) & ".", "Chars")
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

Private Sub AssignByteIndexAndCount(ByRef bytes() As Byte, ByRef vIndex As Variant, ByRef Index As Long, ByRef vCount As Variant, ByRef Count As Long)
    If cArray.IsNull(bytes) Then _
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Array), "Bytes")
    
    If IsMissing(vIndex) Then
        Index = LBound(bytes)
    Else
        Index = vIndex
    End If
    If IsMissing(vCount) Then
        Count = cArray.GetLength(bytes)
    Else
        Count = vCount
    End If
End Sub

Private Sub Class_Terminate()
    SAPtr(mChars) = 0
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   cObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function cObject_Equals(Value As Variant) As Boolean
    cObject_Equals = Equals(Value)
End Function

Private Function cObject_GetHashcode() As Long
    cObject_GetHashcode = GetHashCode
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

Private Function Encoding_GetBytesEx(Chars As Variant, ByVal CharIndex As Long, ByVal CharCount As Long, bytes() As Byte, ByVal ByteIndex As Long) As Long
    Encoding_GetBytesEx = GetBytesEx(Chars, CharIndex, CharCount, bytes, ByteIndex)
End Function

Private Function Encoding_GetCharCount(bytes() As Byte, Optional Index As Variant, Optional Count As Variant) As Long
    Encoding_GetCharCount = GetCharCount(bytes, Index, Count)
End Function

Private Function Encoding_GetChars(bytes() As Byte, Optional vIndex As Variant, Optional vCount As Variant) As Integer()
    Encoding_GetChars = GetChars(bytes, vIndex, vCount)
End Function

Private Function Encoding_GetCharsEx(bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, Chars() As Integer, ByVal CharIndex As Long) As Long
    Encoding_GetCharsEx = GetCharsEx(bytes, ByteIndex, ByteCount, Chars, CharIndex)
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

Private Function Encoding_GetString(bytes() As Byte, Optional vIndex As Variant, Optional vCount As Variant) As String
    Encoding_GetString = GetString(bytes, vIndex, vCount)
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
