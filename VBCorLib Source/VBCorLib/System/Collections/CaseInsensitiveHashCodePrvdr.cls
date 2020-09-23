VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CaseInsensitiveHashCodePrvdr"
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
'    Module: CaseInsensitiveHashCodePrvdr
'

''
' Generates hashcodes for values as normal, but ignores the case for strings.
'
Option Explicit
Implements cObject
Implements IHashcodeProvider

Private mStringHashChars As WordBuffer


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
' @param value The value to compare equalit to.
' @return Boolean indicating equality.
Public Function Equals(ByRef Value As Variant) As Boolean
    If IsObject(Value) Then
        Equals = (Value Is Me)
    End If
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @param Value The value or variable to generate a hashcode for.
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode(Optional ByRef Value As Variant) As Long
    If Not IsMissing(Value) Then
        GetHashCode = IHashcodeProvider_GetHashCode(Value)
    Else
        GetHashCode = ObjPtr(Me)
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function HashString(ByRef Key As Variant) As Long
    Dim i As Long
    Dim l As Long
    Dim ch As Integer
    
    l = Len(Key)
    With mStringHashChars.SA
        .cElements = l
        .pvData = StrPtr(Key)
    End With
    For i = 0 To l - 1
        ch = mStringHashChars.Data(i)
        If ch >= 97 And ch <= 122 Then ch = ch - 32
        HashString = ((HashString * 16) + ch) And &H3FFFFFF
    Next i
End Function


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
'   IHashCodeProvider Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IHashcodeProvider_GetHashCode(Value As Variant) As Long
    Select Case VarType(Value)
        Case vbLong, vbInteger, vbByte, vbBoolean
            IHashcodeProvider_GetHashCode = CLng(Value)
        Case vbString
            IHashcodeProvider_GetHashCode = HashString(Value)
        Case vbDouble, vbDate, vbCurrency
            Dim ptr As Long
            If VariantType(Value) And VT_BYREF Then
                ptr = MemLong(VarPtr(Value) + 8)
            Else
                ptr = VarPtr(Value) + 8
            End If
            With AsDLong(ByVal ptr)
                IHashcodeProvider_GetHashCode = .LoDWord Xor .HiDWord
            End With
        Case vbSingle
            IHashcodeProvider_GetHashCode = AsLong(CSng(Value))
        Case vbObject, vbDataObject
            If TypeOf Value Is cObject Then
                Dim o As cObject
                Set o = Value
                IHashcodeProvider_GetHashCode = o.GetHashCode
            Else
                IHashcodeProvider_GetHashCode = ObjPtr(Value)
            End If
        Case Else
            Throw Cor.NewArgumentException("The specific value type is not supported.", "Key")
    End Select
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    InitWordBuffer mStringHashChars, 0, &H7FFFFFFF
End Sub
