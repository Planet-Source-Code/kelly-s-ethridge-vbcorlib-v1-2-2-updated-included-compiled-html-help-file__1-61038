VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CodePageDecoder"
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
'    Module: CodePageDecoder
'

''
' This class is a general decoding wrapper around the Windows decoding APIs.
'
Option Explicit
Implements cObject
Implements Decoder

Private Declare Function IsDBCSLeadByteEx Lib "kernel32.dll" (ByVal CodePage As Long, ByVal TestChar As Byte) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Byte, ByVal cchMultiByte As Long, ByRef lpWideCharStr As Integer, ByVal cchWideChar As Long) As Long

Private mCodePage As Long
Private mMaxCharSize As Long
Private mLastByte As Long
Private mTemp(1) As Byte



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal MaxCharSize As Long, ByVal CodePage As Long)
    mLastByte = -1
    mMaxCharSize = MaxCharSize
    mCodePage = CodePage
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BytesAreCharacterAligned(ByRef Bytes() As Byte, ByVal Index As Long, ByVal Count As Long) As Boolean
    Dim e As Long
    Dim i As Long
    
    e = Index + Count
    Do While i > Index
        If Not IsDBCSLeadByteEx(mCodePage, Bytes(i - 1)) Then Exit Do
        i = i - 1
    Loop
    BytesAreCharacterAligned = (((e - i) And 1) = 0)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   cObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function cObject_Equals(Value As Variant) As Boolean
    If IsObject(Value) Then
        cObject_Equals = (Value Is Me)
    End If
End Function

Private Function cObject_GetHashCode() As Long
    cObject_GetHashCode = ObjPtr(Me)
End Function

Private Function cObject_ToString() As String
    cObject_ToString = App.Title & "." & TypeName(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Decoder Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Decoder_Equals(Value As Variant) As Boolean
    Decoder_Equals = cObject_Equals(Value)
End Function

Private Function Decoder_GetCharCount(Bytes() As Byte, ByVal Index As Long, ByVal Count As Long) As Long
    Dim offset As Long
    
    If cArray.IsNull(Bytes) Then _
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Array), "Bytes")
    If Index < LBound(Bytes) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "Index", Index)
    If Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "Count", Count)
    If Index + Count - 1 > UBound(Bytes) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "Count")
    
    If mLastByte >= 0 Then
        Index = Index + 1
        Count = Count - 1
        offset = 1
        If Count = 0 Then
            Decoder_GetCharCount = 1
            Exit Function
        End If
    End If
    
    If Not BytesAreCharacterAligned(Bytes, Index, Count) Then
        Count = Count - 1
        If Count = 0 Then
            Decoder_GetCharCount = offset
            Exit Function
        End If
    End If
    
    Decoder_GetCharCount = offset + MultiByteToWideChar(mCodePage, 0, Bytes(Index), Count, 0, 0)
End Function

Private Function Decoder_GetChars(Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, Chars() As Integer, ByVal CharIndex As Long) As Long
    Dim offset As Long
    
    Encoding.GetCharsArgumentCheck Bytes, ByteIndex, ByteCount, Chars, CharIndex
    
    If ByteCount = 0 Then Exit Function
    If mLastByte >= 0 Then
        If CharIndex > UBound(Chars) Then _
            Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_SmallConversionBuffer), "Chars")
        
        mTemp(0) = mLastByte
        mTemp(1) = Bytes(ByteIndex)
        MultiByteToWideChar mCodePage, 0, mTemp(0), 2, Chars(CharIndex), 1
        ByteIndex = ByteIndex + 1
        ByteCount = ByteCount - 1
        CharIndex = CharIndex + 1
        mLastByte = -1
        offset = 1
        If ByteCount = 0 Then
            Decoder_GetChars = 1
            Exit Function
        End If
    End If
    
    If mMaxCharSize > 1 Then
        If Not BytesAreCharacterAligned(Bytes, ByteIndex, ByteCount) Then
            ByteCount = ByteCount - 1
            mLastByte = Bytes(ByteIndex + ByteCount)
            If ByteCount = 0 Then
                Decoder_GetChars = offset
                Exit Function
            End If
        End If
    End If

    If CharIndex > UBound(Chars) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_SmallConversionBuffer), "Chars")
    
    Decoder_GetChars = offset + MultiByteToWideChar(mCodePage, 0, Bytes(ByteIndex), ByteCount, Chars(CharIndex), UBound(Chars) - CharIndex + 1)
End Function

Private Function Decoder_GetHashCode() As Long
    Decoder_GetHashCode = cObject_GetHashCode
End Function

Private Function Decoder_ToString() As String
    Decoder_ToString = cObject_ToString
End Function
