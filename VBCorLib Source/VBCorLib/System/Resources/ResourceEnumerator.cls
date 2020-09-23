VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResourceEnumerator"
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
'    Module: ResourceEnumerator
'

Option Explicit
Implements cObject
Implements IEnumerator
Implements IDictionaryEnumerator

Private mResReader As Object
Private mResourceSets As Hashtable
Private mResIterator As IDictionaryEnumerator
Private mLangGroupIterator As IDictionaryEnumerator
Private mResTypeIterator As IDictionaryEnumerator




Friend Sub Init(ByVal Reader As Object, ByVal ResourceSets As Hashtable)
    Set mResReader = Reader
    Set mResourceSets = ResourceSets
    Set mLangGroupIterator = mResourceSets.GetEnumerator
End Sub


''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Private Function ToString() As String
    ToString = App.Title & "." & TypeName(Me)
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equalit to.
' @return Boolean indicating equality.
Private Function Equals(ByRef Value As Variant) As Boolean
    If IsObject(Value) Then
        Equals = (Value Is Me)
    End If
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Private Function GetHashCode() As Long
    GetHashCode = ObjPtr(Me)
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
'   IDictionaryEnumerator Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get IDictionaryEnumerator_Current() As Variant
    Set IDictionaryEnumerator_Current = IDictionaryEnumerator_Entry
End Property

Private Property Get IDictionaryEnumerator_Entry() As DictionaryEntry
    Dim res As Win32Resource
    
    If mLangGroupIterator Is Nothing Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumFinished))
    If mResIterator Is Nothing Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumNotStarted))

    Set res = mResIterator.Value
    Set IDictionaryEnumerator_Entry = Cor.NewDictionaryEntry(res.Ordinal, res)
End Property

Private Property Get IDictionaryEnumerator_Key() As Variant
    Dim res As Win32Resource
    
    If mLangGroupIterator Is Nothing Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumFinished))
    If mResIterator Is Nothing Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumNotStarted))
    
    Set res = mResIterator.Value
    IDictionaryEnumerator_Key = res.Ordinal
End Property

Private Function IDictionaryEnumerator_MoveNext() As Boolean
    Dim ht As Hashtable
    
    If mLangGroupIterator Is Nothing Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumFinished))
    
    Do
        Do
            If mResTypeIterator Is Nothing Then
                If Not mLangGroupIterator.MoveNext Then
                    Set mLangGroupIterator = Nothing
                    Set mResTypeIterator = Nothing
                    Set mResIterator = Nothing
                    Exit Function
                End If
                Set ht = mLangGroupIterator.Value
                Set mResTypeIterator = ht.GetEnumerator
            End If
            
            If mResIterator Is Nothing Then
                If mResTypeIterator.MoveNext Then
                    Set ht = mResTypeIterator.Value
                    Set mResIterator = ht.GetEnumerator
                    Exit Do
                Else
                    Set mResTypeIterator = Nothing
                End If
            Else
                Exit Do
            End If
        Loop
        
        If mResIterator.MoveNext Then Exit Do
        Set mResIterator = Nothing
    Loop
    
    IDictionaryEnumerator_MoveNext = True
End Function

Private Sub IDictionaryEnumerator_Reset()
    mLangGroupIterator.Reset
    Set mResIterator = Nothing
    Set mResTypeIterator = Nothing
End Sub

Private Property Get IDictionaryEnumerator_value() As Variant
    Dim res As Win32Resource
    
    Set res = mResIterator.Value
    Helper.MoveVariant IDictionaryEnumerator_value, res.Value
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEnumerator Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get IEnumerator_Current() As Variant
    VariantCopyInd IEnumerator_Current, IDictionaryEnumerator_Current
End Property

Private Function IEnumerator_MoveNext() As Boolean
    IEnumerator_MoveNext = IDictionaryEnumerator_MoveNext
End Function

Private Sub IEnumerator_Reset()
    IDictionaryEnumerator_Reset
End Sub
