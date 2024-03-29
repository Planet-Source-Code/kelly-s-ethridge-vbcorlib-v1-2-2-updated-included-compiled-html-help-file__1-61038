VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NullTextReader"
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
'    Module: NullTextReader
'

''
' Provides a TextReader with no data backing.
'
' @see TextReader
'
Option Explicit
Implements cObject
Implements TextReader


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   cObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function cObject_Equals(Value As Variant) As Boolean
    If IsObject(Value) Then
        cObject_Equals = Value Is Me
    End If
End Function

Private Function cObject_GetHashCode() As Long
    cObject_GetHashCode = ObjPtr(Me)
End Function

Private Function cObject_ToString() As String
    cObject_ToString = App.Title & "." & TypeName(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   TextReader Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TextReader_CloseReader()
' do nothing
End Sub

Private Function TextReader_Equals(Value As Variant) As Boolean
    TextReader_Equals = cObject_Equals(Value)
End Function

Private Function TextReader_GetHashCode() As Long
    TextReader_GetHashCode = cObject_GetHashCode
End Function

Private Function TextReader_Peek() As Long
    TextReader_Peek = -1
End Function

Private Function TextReader_Read() As Long
    TextReader_Read = -1
End Function

Private Function TextReader_ReadBlock(Chars() As Integer, ByVal Index As Long, ByVal Count As Long) As Long
' do nothing
End Function

Private Function TextReader_ReadLine() As String
    TextReader_ReadLine = vbNullString
End Function

Private Function TextReader_ReadToEnd() As String
    TextReader_ReadToEnd = vbNullString
End Function

Private Function TextReader_ToString() As String
    TextReader_ToString = cObject_ToString
End Function
