VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CodePageEncoding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
'    Module: CodePageEncoding
'

''
' An encoding class that uses the underlying encoding support of Windows.
'

Option Explicit
Implements cObject
Implements Encoding

Private Declare Function GetCPInfo Lib "kernel32.dll" (ByVal CodePage As Long, ByRef lpCPInfo As CPINFO) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpWideCharStr As Integer, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Byte, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Byte, ByVal cchMultiByte As Long, ByRef lpWideCharStr As Integer, ByVal cchWideChar As Long) As Long

Private Const MAX_DEFAULTCHAR As Long = 2
Private Const MAX_LEADBYTES As Long = 12

Private Type CPINFO
    MaxCharSize As Long ' max length (Byte) of a char
    DefaultChar(MAX_DEFAULTCHAR) As Byte ' default character
    LeadByte(MAX_LEADBYTES) As Byte ' lead byte ranges
End Type


Private mCodePage As Long
Private mMaxCharSize As Long

Private mChars() As Integer
Private mCharsSA As SafeArray1d



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal CodePage As Long)
    Dim info As CPINFO
    
    If GetCPInfo(CodePage, info) = BOOL_FALSE Then _
        Throw Cor.NewNotSupportedException("Either the Code Page is invalid or the Encoding is not installed.")
    
    mCodePage = CodePage
    mMaxCharSize = info.MaxCharSize
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get Table() As EncodingTable
    Set Table = modPublicFunctions.Encoding.EncodingTable
End Property

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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
    SAPtr(mChars) = 0
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   cObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function cObject_Equals(Value As Variant) As Boolean
    Dim ec As Encoding
    
    If IsObject(Value) Then
        If TypeOf Value Is Encoding Then
            Set ec = Value
            cObject_Equals = (ec.CodePage = mCodePage)
        End If
    End If
End Function

Private Function cObject_GetHashcode() As Long
    cObject_GetHashcode = mCodePage
End Function

Private Function cObject_ToString() As String
    cObject_ToString = "VBCorLib.Encoding"
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Encoding Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get Encoding_BodyName() As String
    Encoding_BodyName = Table.BodyName(mCodePage)
End Property

Private Property Get Encoding_CodePage() As Long
    Encoding_CodePage = mCodePage
End Property

Private Property Get Encoding_EncodingName() As String
    Encoding_EncodingName = Table.EncodingName(mCodePage)
End Property

Private Function Encoding_Equals(Value As Variant) As Boolean
    Encoding_Equals = cObject_Equals(Value)
End Function

Private Function Encoding_GetByteCount(Chars As Variant, Optional vIndex As Variant, Optional vCount As Variant) As Long
    Dim Index As Long
    Dim Count As Long
    
    AttachChars Chars, vIndex, Index, vCount, Count
    If Index < LBound(mChars) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "vIndex", Index)
    If Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "vCount", Count)
    If Index + Count - 1 > UBound(mChars) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "vCount")
    
    If mMaxCharSize = 1 Or Count = 0 Then
        Encoding_GetByteCount = Count
    Else
        Encoding_GetByteCount = WideCharToMultiByte(mCodePage, 0, mChars(Index), Count, vbNullString, 0, vbNullString, 0)
    End If
    SAPtr(mChars) = 0
End Function

Private Function Encoding_GetBytes(Chars As Variant, Optional vIndex As Variant, Optional vCount As Variant) As Byte()
    Dim Index As Long
    Dim Count As Long
    Dim Size As Long
    Dim ret() As Byte
    
    AttachChars Chars, vIndex, Index, vCount, Count
    Size = Encoding_GetMaxByteCount(Count)
    ReDim ret(0 To Size - 1)
    modPublicFunctions.Encoding.GetBytesArgumentCheck mChars, Index, Count, ret, 0

    Count = WideCharToMultiByte(mCodePage, 0, mChars(Index), Count, ret(0), Size, vbNullString, 0)
    If Count <> Size Then ReDim Preserve ret(0 To Count - 1)
    SAPtr(mChars) = 0
    Encoding_GetBytes = ret
End Function

Private Function Encoding_GetBytesEx(Chars As Variant, ByVal CharIndex As Long, ByVal CharCount As Long, bytes() As Byte, ByVal ByteIndex As Long) As Long
    AttachChars Chars, CharIndex, 0, CharCount, 0
    modPublicFunctions.Encoding.GetBytesArgumentCheck mChars, CharIndex, CharCount, bytes, ByteIndex
    Encoding_GetBytesEx = WideCharToMultiByte(mCodePage, 0, mChars(CharIndex), CharCount, bytes(ByteIndex), cArray.GetLength(bytes), vbNullString, 0)
    SAPtr(mChars) = 0
End Function

Private Function Encoding_GetCharCount(bytes() As Byte, Optional vIndex As Variant, Optional vCount As Variant) As Long
    Dim Index As Long
    Dim Count As Long
    
    AssignByteIndexAndCount bytes, vIndex, Index, vCount, Count
    If Index < LBound(bytes) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "vIndex", Index)
    If Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "vCount", Count)
    If Index + Count - 1 > UBound(bytes) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "vCount")
    
    If Count = 0 Then Exit Function
    Encoding_GetCharCount = MultiByteToWideChar(mCodePage, 0, bytes(Index), Count, 0, 0)
End Function

Private Function Encoding_GetChars(bytes() As Byte, Optional vIndex As Variant, Optional vCount As Variant) As Integer()
    Dim Index As Long
    Dim Count As Long
    Dim ret() As Integer
    Dim Size As Long
    
    Size = Encoding_GetMaxCharCount(cArray.GetLength(bytes))
    AssignByteIndexAndCount bytes, vIndex, Index, vCount, Count
    ReDim ret(0 To Size - 1)
    Size = Encoding_GetCharsEx(bytes, Index, Count, ret, 0)
    ReDim Preserve ret(0 To Size - 1)
    SAPtr(mChars) = 0
    Encoding_GetChars = ret
End Function

Private Function Encoding_GetCharsEx(bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, Chars() As Integer, ByVal CharIndex As Long) As Long
    modPublicFunctions.Encoding.GetCharsArgumentCheck bytes, ByteIndex, ByteCount, Chars, CharIndex
    Encoding_GetCharsEx = MultiByteToWideChar(mCodePage, 0, bytes(ByteIndex), ByteCount, Chars(CharIndex), UBound(Chars) - CharIndex + 1)
End Function

Private Function Encoding_GetDecoder() As Decoder
    Dim ret As New CodePageDecoder
    ret.Init mMaxCharSize, mCodePage
    Set Encoding_GetDecoder = ret
End Function

Private Function Encoding_GetEncoder() As Encoder
    Dim ret As New DefaultEncoder
    ret.Init Me
    Set Encoding_GetEncoder = ret
End Function

Private Function Encoding_GetHashCode() As Long
    Encoding_GetHashCode = mCodePage
End Function

Private Function Encoding_GetMaxByteCount(ByVal CharCount As Long) As Long
    If CharCount < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "CharCount", CharCount)
        
    Encoding_GetMaxByteCount = CharCount * mMaxCharSize
End Function

Private Function Encoding_GetMaxCharCount(ByVal ByteCount As Long) As Long
    If ByteCount < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "ByteCount", ByteCount)
    
    Encoding_GetMaxCharCount = ByteCount
End Function

Private Function Encoding_GetPreamble() As Byte()
    Encoding_GetPreamble = cArray.CreateInstance(ciByte, 0)
End Function

Private Function Encoding_GetString(bytes() As Byte, Optional vIndex As Variant, Optional vCount As Variant) As String
    Dim ret As String
    Dim Size As Long
    Dim Index As Long
    Dim Count As Long
    
    AssignByteIndexAndCount bytes, vIndex, Index, vCount, Count
    Size = Encoding_GetMaxCharCount(Count)
    ret = SysAllocStringLen(0, Size)
    With mCharsSA
        .pvData = StrPtr(ret)
        .lLbound = 0
        .cElements = Size
    End With
    SAPtr(mChars) = VarPtr(mCharsSA)
    Size = Encoding_GetCharsEx(bytes, Index, Count, mChars, 0)
    SAPtr(mChars) = 0
    Encoding_GetString = Left$(ret, Size)
End Function

Private Property Get Encoding_HeaderName() As String
    Encoding_HeaderName = Table.HeaderName(mCodePage)
End Property

Private Property Get Encoding_IsBrowserDisplay() As Boolean
    Encoding_IsBrowserDisplay = Table.IsBrowserDisplay(mCodePage)
End Property

Private Property Get Encoding_IsBrowserSave() As Boolean
    Encoding_IsBrowserSave = Table.IsBrowserSave(mCodePage)
End Property

Private Property Get Encoding_IsMailNewsDisplay() As Boolean
    Encoding_IsMailNewsDisplay = Table.IsMailNewsDisplay(mCodePage)
End Property

Private Property Get Encoding_IsMailNewsSave() As Boolean
    Encoding_IsMailNewsSave = Table.IsMailNewsSave(mCodePage)
End Property

Private Function Encoding_ToString() As String
    Encoding_ToString = "VBCorLib.Encoding"
End Function

Private Property Get Encoding_WebName() As String
    Encoding_WebName = Table.WebName(mCodePage)
End Property

Private Property Get Encoding_WindowsCodePage() As Long
    Encoding_WindowsCodePage = Table.WindowsCodePage(mCodePage)
End Property
