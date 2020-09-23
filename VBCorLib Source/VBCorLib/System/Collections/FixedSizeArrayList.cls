VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FixedSizeArrayList"
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
'    Module: FixedSizeArrayList
'
Option Explicit
Implements cObject
Implements ArrayList
Implements ICollection
Implements IEnumerable
Implements ICloneable
Implements IList

Private mList As ArrayList


Friend Sub Init(ByVal list As ArrayList)
    Set mList = list
End Sub



Private Function ArrayList_Add(value As Variant) As Long
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_FixedSizeCollection))
End Function

Private Sub ArrayList_AddRange(c As Variant)
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_FixedSizeCollection))
End Sub

Private Function ArrayList_BinarySearch(value As Variant, Optional ByVal StartIndex As Long, Optional ByRef Count As Variant, Optional ByVal Comparer As IComparer) As Long
    ArrayList_BinarySearch = mList.BinarySearch(value, StartIndex, Count, Comparer)
End Function

Private Property Let ArrayList_Capacity(ByVal RHS As Long)
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_FixedSizeCollection))
End Property

Private Property Get ArrayList_Capacity() As Long
    ArrayList_Capacity = mList.capacity
End Property

Private Sub ArrayList_Clear()
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_FixedSizeCollection))
End Sub

Private Function ArrayList_Clone() As Object
    Dim ret As New FixedSizeArrayList
    ret.Init mList.Clone
    Set ArrayList_Clone = ret
End Function

Private Function ArrayList_Contains(value As Variant, Optional ByVal Comparer As IComparer) As Boolean
    ArrayList_Contains = mList.Contains(value, Comparer)
End Function

Private Sub ArrayList_CopyTo(dstArray As Variant, Optional arrayindex As Variant)
    mList.CopyTo dstArray, arrayindex
End Sub

Private Sub ArrayList_CopyToEx(ByVal Index As Long, dstArray As Variant, ByVal arrayindex As Long, ByVal Count As Long)
    mList.CopyToEx Index, dstArray, arrayindex, Count
End Sub

Private Property Get ArrayList_count() As Long
    ArrayList_count = mList.Count
End Property

Private Function ArrayList_Equals(value As Variant) As Boolean
    ArrayList_Equals = cObject_Equals(value)
End Function

Private Function ArrayList_GetEnumerator(Optional ByVal StartIndex As Long, Optional ByRef Count As Variant) As Object
    Dim cnt As Long
    
    If IsMissing(Count) Then
        cnt = mList.Count
    Else
        cnt = Count
    End If
    
    If StartIndex < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "startindex", StartIndex)
    If cnt < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "Count", Count)
    If StartIndex + Count > mList.Count Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "count")
    
    Dim ret As New ArrayListEnumerator
    ret.Init Me, StartIndex, cnt
    Set ArrayList_GetEnumerator = ret
End Function

Private Function ArrayList_GetHashCode() As Long
    ArrayList_GetHashCode = cObject_GetHashcode
End Function

Private Function ArrayList_GetRange(ByVal Index As Long, ByVal Count As Long) As ArrayList
    If Index < 0 Or Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), IIf(Index < 0, "index", "count"), IIf(Index < 0, Index, Count))
    If Index + Count > mList.Count Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "count")
    Dim ret As RangedArrayList
    Set ret = New RangedArrayList
    ret.Init Me, Index, Count
    Set ArrayList_GetRange = ret
End Function

Private Function ArrayList_IndexOf(value As Variant, Optional ByVal StartIndex As Long, Optional ByRef Count As Variant, Optional ByVal Comparer As IComparer) As Long
    ArrayList_IndexOf = mList.IndexOf(value, StartIndex, Count, Comparer)
End Function

Private Sub ArrayList_Insert(ByVal Index As Long, value As Variant)
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_FixedSizeCollection))
End Sub

Private Sub ArrayList_InsertRange(ByVal Index As Long, c As Variant)
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_FixedSizeCollection))
End Sub

Private Property Get ArrayList_IsFixedSize() As Boolean
    ArrayList_IsFixedSize = True
End Property

Private Property Get ArrayList_IsReadOnly() As Boolean
    ArrayList_IsReadOnly = False
End Property

Private Property Set ArrayList_Item(ByVal Index As Long, RHS As Variant)
    Set mList(Index) = RHS
End Property

Private Property Let ArrayList_Item(ByVal Index As Long, RHS As Variant)
    mList(Index) = RHS
End Property

Private Property Get ArrayList_Item(ByVal Index As Long) As Variant
    Helper.MoveVariant ArrayList_Item, mList(Index)
End Property

Private Function ArrayList_LastIndexOf(value As Variant, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant, Optional ByVal Comparer As IComparer) As Long
    ArrayList_LastIndexOf = mList.LastIndexOf(value, StartIndex, Count, Comparer)
End Function

Private Function ArrayList_NewEnum() As stdole.IUnknown
    Set ArrayList_NewEnum = CreateEnumerator(ArrayList_GetEnumerator)
End Function

Private Sub ArrayList_Remove(value As Variant, Optional ByVal Comparer As IComparer)
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_FixedSizeCollection))
End Sub

Private Sub ArrayList_RemoveAt(ByVal Index As Long)
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_FixedSizeCollection))
End Sub

Private Sub ArrayList_RemoveRange(ByVal Index As Long, ByVal Count As Long)
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_FixedSizeCollection))
End Sub

Private Sub ArrayList_Reverse(Optional ByVal Index As Long, Optional ByRef Count As Variant)
    mList.Reverse Index, Count
End Sub

Private Sub ArrayList_SetRange(ByVal Index As Long, c As Variant)
    mList.SetRange Index, c
End Sub

Private Sub ArrayList_Sort(Optional ByVal StartIndex As Long, Optional ByRef Count As Variant, Optional ByVal Comparer As IComparer)
    mList.Sort StartIndex, Count, Comparer
End Sub

Private Function ArrayList_ToArray(Optional ByVal arraytype As ciArrayTypes = 12&) As Variant
    ArrayList_ToArray = mList.ToArray(arraytype)
End Function

Private Function ArrayList_ToString() As String
    ArrayList_ToString = mList.ToString
End Function

Private Sub ArrayList_TrimToSize()
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_FixedSizeCollection))
End Sub

Private Property Get ArrayList_Version() As Long
    ArrayList_Version = mList.Version
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICloneable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = ArrayList_Clone
End Function


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
Private Sub ICollection_CopyTo(arr As Variant, ByVal Index As Long)
    ArrayList_CopyTo arr, Index
End Sub

Private Property Get ICollection_Count() As Long
    ICollection_Count = ArrayList_count
End Property

Private Function ICollection_GetEnumerator() As IEnumerator
    Set ICollection_GetEnumerator = ArrayList_GetEnumerator
End Function

Private Function ICollection_NewEnum() As stdole.IUnknown
    Set ICollection_NewEnum = ArrayList_NewEnum
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEnumerable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IEnumerable_GetEnumerator() As IEnumerator
    Set IEnumerable_GetEnumerator = ArrayList_GetEnumerator
End Function

Private Function IEnumerable_NewEnum() As stdole.IUnknown
    Set IEnumerable_NewEnum = ArrayList_NewEnum
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IList Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IList_Add(value As Variant) As Long
    IList_Add = ArrayList_Add(value)
End Function

Private Sub IList_Clear()
    ArrayList_Clear
End Sub

Private Function IList_Contains(value As Variant, Optional ByVal Comparer As IComparer) As Boolean
    IList_Contains = ArrayList_Contains(value, Comparer)
End Function

Private Sub IList_CopyTo(arr As Variant, ByVal Index As Long)
    ArrayList_CopyTo arr, Index
End Sub

Private Property Get IList_Count() As Long
    IList_Count = ArrayList_count
End Property

Private Function IList_GetEnumerator() As IEnumerator
    Set IList_GetEnumerator = ArrayList_GetEnumerator
End Function

Private Function IList_IndexOf(value As Variant, Optional ByVal Comparer As IComparer) As Long
    IList_IndexOf = ArrayList_IndexOf(value, , , Comparer)
End Function

Private Sub IList_Insert(ByVal Index As Long, value As Variant)
    ArrayList_Insert Index, value
End Sub

Private Property Get IList_IsFixedSize() As Boolean
    IList_IsFixedSize = ArrayList_IsFixedSize
End Property

Private Property Get IList_IsReadOnly() As Boolean
    IList_IsReadOnly = ArrayList_IsReadOnly
End Property

Private Property Set IList_Item(ByVal Index As Long, RHS As Variant)
    Set ArrayList_Item(Index) = RHS
End Property

Private Property Let IList_Item(ByVal Index As Long, RHS As Variant)
    ArrayList_Item(Index) = RHS
End Property

Private Property Get IList_Item(ByVal Index As Long) As Variant
    Helper.MoveVariant IList_Item, ArrayList_Item(Index)
End Property

Private Function IList_NewEnum() As stdole.IUnknown
    Set IList_NewEnum = ArrayList_NewEnum
End Function

Private Sub IList_Remove(value As Variant, Optional ByVal Comparer As IComparer)
    ArrayList_Remove value, Comparer
End Sub

Private Sub IList_RemoveAt(ByVal Index As Long)
    ArrayList_RemoveAt Index
End Sub

