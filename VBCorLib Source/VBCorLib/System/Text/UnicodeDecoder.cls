VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UnicodeDecoder"
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
'    Module: UnicodeDecoder
'
Option Explicit
Implements cObject
Implements Decoder

Private Const CHAR_SIZE As Long = 2


Private mIsBigEndian As Boolean
Private mLastByte As Long



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal IsBigEndian As Boolean)
    mIsBigEndian = IsBigEndian
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mLastByte = -1
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   cObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function cObject_ToString() As String
    cObject_ToString = App.Title & "." & TypeName(Me)
End Function

Private Function cObject_Equals(ByRef value As Variant) As Boolean
    If IsObject(value) Then
        cObject_Equals = (value Is Me)
    End If
End Function

Private Function cObject_GetHashcode() As Long
    cObject_GetHashcode = ObjPtr(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Decoder Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Decoder_Equals(value As Variant) As Boolean
    Decoder_Equals = cObject_Equals(value)
End Function

Private Function Decoder_GetCharCount(Bytes() As Byte, ByVal Index As Long, ByVal Count As Long) As Long
    If cArray.IsNull(Bytes) Then _
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Array), "Bytes")
    If Index < LBound(Bytes) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "vIndex", Index)
    If Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "vCount", Count)
    If Index + Count - 1 > UBound(Bytes) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "vCount")
    
    If mLastByte >= 0 Then Count = Count + 1
    Decoder_GetCharCount = Count \ CHAR_SIZE
End Function

Private Function Decoder_GetChars(Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, Chars() As Integer, ByVal CharIndex As Long) As Long
    Dim i As Long
    Dim ch As Long
    Dim TotalBytes As Long
    Dim Offset As Long
    
    
    modPublicFunctions.Encoding.GetCharsArgumentCheck Bytes, ByteIndex, TotalBytes, Chars, CharIndex
    If mIsBigEndian Then
        If mLastByte >= 0 Then
            TotalBytes = ByteCount + 1
        Else
            TotalBytes = ByteCount
        End If
        Do While i < TotalBytes
            If mLastByte >= 0 Then
                ch = mLastByte * &H100
                i = i - 1
                mLastByte = -1
            Else
                ch = Bytes(ByteIndex + i) * &H100
            End If
            i = i + 1
            ch = ch Or Bytes(ByteIndex + i)
            Chars(CharIndex) = AsWord(ch)
            CharIndex = CharIndex + 1
            i = i + 1
        Loop
        If (TotalBytes And 1) = 1 Then mLastByte = Bytes(ByteIndex + ByteCount - 1)
        Decoder_GetChars = TotalBytes \ 2
    Else
        If mLastByte >= 0 Then
            Chars(CharIndex) = mLastByte
            Offset = 1
        End If
        CopyMemory ByVal VarPtr(Chars(CharIndex)) + Offset, Bytes(ByteIndex), ByteCount - Offset
        mLastByte = -1
        If ((ByteCount + Offset) And 1) = 1 Then mLastByte = Bytes(ByteIndex + ByteCount - 1)
        Decoder_GetChars = (ByteCount + Offset) \ 2
    End If
End Function

Private Function Decoder_GetHashCode() As Long
    Decoder_GetHashCode = cObject_GetHashcode
End Function

Private Function Decoder_ToString() As String
    Decoder_ToString = cObject_ToString
End Function

