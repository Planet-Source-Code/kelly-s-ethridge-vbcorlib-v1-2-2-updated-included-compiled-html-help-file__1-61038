VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SortedValueList"
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
'    Module: SortedValueList
'
Option Explicit
Implements cObject
Implements IList
Implements ICollection
Implements IEnumerable

Private mList As SortedList


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = CreateEnumerator(mList.GetValueEnumerator)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal list As SortedList)
    Set mList = list
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   cObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function cObject_Equals(value As Variant) As Boolean
    If IsObject(value) Then
        cObject_Equals = (value Is Me)
    End If
End Function

Private Function cObject_GetHashcode() As Long
    cObject_GetHashcode = ObjPtr(Me)
End Function

Private Function cObject_ToString() As String
    cObject_ToString = App.Title & "." & TypeName(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICollection Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ICollection_CopyTo(arr As Variant, ByVal index As Long)
    mList.CopyValuesTo arr, index
End Sub

Private Property Get ICollection_Count() As Long
    ICollection_Count = mList.Count
End Property

Private Function ICollection_GetEnumerator() As IEnumerator
    Set ICollection_GetEnumerator = mList.GetValueEnumerator
End Function

Private Function ICollection_NewEnum() As stdole.IUnknown
    Set ICollection_NewEnum = NewEnum
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEnumerable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IEnumerable_GetEnumerator() As IEnumerator
    Set IEnumerable_GetEnumerator = mList.GetValueEnumerator
End Function

Private Function IEnumerable_NewEnum() As stdole.IUnknown
    Set IEnumerable_NewEnum = NewEnum
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IList Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IList_Add(value As Variant) As Long
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_ReadOnlyCollection))
End Function

Private Sub IList_Clear()
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_ReadOnlyCollection))
End Sub

Private Function IList_Contains(value As Variant, Optional ByVal comparer As IComparer) As Boolean
    IList_Contains = mList.ContainsValue(value, comparer)
End Function

Private Sub IList_CopyTo(arr As Variant, ByVal index As Long)
    mList.CopyValuesTo arr, index
End Sub

Private Property Get IList_Count() As Long
    IList_Count = mList.Count
End Property

Private Function IList_GetEnumerator() As IEnumerator
    Set IList_GetEnumerator = mList.GetValueEnumerator
End Function

Private Function IList_IndexOf(value As Variant, Optional ByVal comparer As IComparer) As Long
    IList_IndexOf = mList.IndexOfValue(value, comparer)
End Function

Private Sub IList_Insert(ByVal index As Long, value As Variant)
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_ReadOnlyCollection))
End Sub

Private Property Get IList_IsFixedSize() As Boolean
    IList_IsFixedSize = True
End Property

Private Property Get IList_IsReadOnly() As Boolean
    IList_IsReadOnly = True
End Property

Private Property Set IList_Item(ByVal index As Long, RHS As Variant)
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_ReadOnlyCollection))
End Property

Private Property Let IList_Item(ByVal index As Long, RHS As Variant)
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_ReadOnlyCollection))
End Property

Private Property Get IList_Item(ByVal index As Long) As Variant
    Helper.MoveVariant IList_Item, mList.GetByIndex(index)
End Property

Private Function IList_NewEnum() As stdole.IUnknown
    Set IList_NewEnum = NewEnum
End Function

Private Sub IList_Remove(value As Variant, Optional ByVal comparer As IComparer)
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_ReadOnlyCollection))
End Sub

Private Sub IList_RemoveAt(ByVal index As Long)
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_ReadOnlyCollection))
End Sub

