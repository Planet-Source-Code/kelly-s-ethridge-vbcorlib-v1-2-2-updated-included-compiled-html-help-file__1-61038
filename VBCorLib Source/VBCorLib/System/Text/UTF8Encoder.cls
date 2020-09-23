VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UTF8Encoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
'    Module: UTF8Encoder
'

''
' Private class returned by UTF8Encoding.GetEncoder
'

Option Explicit
Implements cObject
Implements Encoder

Private mEncoding As UTF8Encoding
Private mHighSurrogateChar As Long
Private mFlush As Boolean


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Encoding As UTF8Encoding)
    Set mEncoding = Encoding
End Sub

Friend Property Get Flush() As Boolean
    Flush = mFlush
End Property

Friend Property Get HighSurrogateChar() As Long
    HighSurrogateChar = mHighSurrogateChar
End Property

Friend Property Let HighSurrogateChar(ByVal RHS As Long)
    mHighSurrogateChar = RHS
End Property



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   cObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function cObject_ToString() As String
    cObject_ToString = App.Title & "." & TypeName(Me)
End Function

Private Function cObject_Equals(ByRef Value As Variant) As Boolean
    If IsObject(Value) Then
        cObject_Equals = (Value Is Me)
    End If
End Function

Private Function cObject_GetHashcode() As Long
    cObject_GetHashcode = ObjPtr(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Encoder Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Encoder_Equals(Value As Variant) As Boolean
    Encoder_Equals = cObject_Equals(Value)
End Function

Private Function Encoder_GetByteCount(Chars() As Integer, ByVal Index As Long, ByVal Count As Long, ByVal Flush As Boolean) As Long
    mFlush = Flush
    Encoder_GetByteCount = mEncoding.InternalGetByteCount(Chars, Index, Count, Me)
End Function

Private Function Encoder_GetBytes(Chars() As Integer, ByVal CharIndex As Long, ByVal CharCount As Long, Bytes() As Byte, ByVal ByteIndex As Long, ByVal Flush As Boolean) As Long
    mFlush = Flush
    Encoder_GetBytes = mEncoding.InternalGetBytes(Chars, CharIndex, CharCount, Bytes, ByteIndex, Me)
End Function

Private Function Encoder_GetHashCode() As Long
    Encoder_GetHashCode = cObject_GetHashcode
End Function

Private Function Encoder_ToString() As String
    Encoder_ToString = cObject_ToString
End Function
