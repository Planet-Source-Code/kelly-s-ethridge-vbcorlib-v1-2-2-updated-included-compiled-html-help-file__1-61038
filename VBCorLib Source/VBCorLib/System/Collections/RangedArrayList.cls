VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RangedArrayList"
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
'    Module: RangedArrayList
'

''
' This class is used to represent a sub-array view of the parent ArrayList object.
'

Option Explicit
Implements cObject
Implements ArrayList
Implements ICollection
Implements IEnumerable
Implements ICloneable
Implements IList

Private mList As ArrayList
Private mVersion As Long
Private mIndex As Long
Private mCount As Long



Friend Sub Init(ByVal list As ArrayList, ByVal Index As Long, ByVal Count As Long)
    Set mList = list
    mVersion = list.Version
    mIndex = Index
    mCount = Count
End Sub



Private Sub VerifyVersion()
    If mVersion <> mList.Version Then _
        Throw Cor.NewInvalidOperationException("The base ArrayList has been modified outside of the Ranged list.")
End Sub



Private Function ArrayList_Add(Value As Variant) As Long
    VerifyVersion
    mList.Insert mIndex + mCount, Value
    ArrayList_Add = mCount
    mCount = mCount + 1
    mVersion = mVersion + 1
End Function

Private Sub ArrayList_AddRange(c As Variant)
    VerifyVersion
    mList.InsertRange mIndex + mCount, c
    mCount = mCount + mList.GetCollectionAndSize(c)
    mVersion = mVersion + 1
End Sub

Private Function ArrayList_BinarySearch(Value As Variant, Optional ByVal Index As Long, Optional ByRef Count As Variant, Optional ByVal comparer As IComparer) As Long
    Dim cnt As Long
    
    VerifyVersion
    If Count < 0 Then Count = mCount
    If IsMissing(Count) Then
        cnt = mCount
    Else
        cnt = Count
    End If
    
    If Index < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "Index", Index)
    If cnt < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "Count", cnt)
    If Index + cnt > mCount Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "Count")
    
    Dim i As Long
    i = mList.BinarySearch(Value, mIndex + Index, cnt, comparer)
    If i >= 0 Then i = i - mIndex
    ArrayList_BinarySearch = i
End Function

Private Property Let ArrayList_Capacity(ByVal RHS As Long)
    VerifyVersion
    mList.Capacity = RHS
    mVersion = mVersion + 1
End Property

Private Property Get ArrayList_Capacity() As Long
    VerifyVersion
    ArrayList_Capacity = mList.Capacity
End Property

Private Sub ArrayList_Clear()
    VerifyVersion
    If mCount = 0 Then Exit Sub
    mList.RemoveRange mIndex, mCount
    mCount = 0
    mVersion = mVersion + 1
End Sub

Private Function ArrayList_Clone() As Object
    VerifyVersion
    Dim ret As New RangedArrayList
    ret.Init mList.Clone, mIndex, mCount
    Set ArrayList_Clone = ret
End Function

Private Function ArrayList_Contains(Value As Variant, Optional ByVal comparer As IComparer) As Boolean
    VerifyVersion
    ArrayList_Contains = (mList.IndexOf(Value, mIndex, mCount, comparer) >= 0)
End Function

Private Sub ArrayList_CopyTo(dstArray As Variant, Optional arrayindex As Variant)
    Dim Index As Long
    VerifyVersion
    If IsMissing(arrayindex) Then Index = LBound(dstArray) Else Index = arrayindex
    mList.CopyToEx mIndex, dstArray, Index, mCount
End Sub

Private Sub ArrayList_CopyToEx(ByVal Index As Long, dstArray As Variant, ByVal arrayindex As Long, ByVal Count As Long)
    VerifyVersion
    If Index < 0 Or Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), IIf(Index < 0, "index", "count"), IIf(Index < 0, Index, Count))
    If Index + Count > mCount Then _
        Throw Cor.NewArgumentException("Invalid offset length.", "count")
    mList.CopyToEx mIndex + Index, dstArray, arrayindex, Count
End Sub

Private Property Get ArrayList_count() As Long
    VerifyVersion
    ArrayList_count = mCount
End Property

Private Function ArrayList_Equals(Value As Variant) As Boolean
    ArrayList_Equals = cObject_Equals(Value)
End Function

Private Function ArrayList_GetEnumerator(Optional ByVal StartIndex As Long, Optional ByRef Count As Variant) As Object
    Dim cnt As Long
    
    VerifyVersion
    
    If IsMissing(Count) Then
        cnt = mCount
    Else
        cnt = Count
    End If
    
    If StartIndex < 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "startindex")
    If cnt < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "Count", cnt)
    If StartIndex + cnt > mCount Then _
        Throw Cor.NewArgumentException("Invalid count offset.", "Count")
    
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
    If Index + Count > mCount Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "count")
    Dim ret As RangedArrayList
    Set ret = New RangedArrayList
    ret.Init Me, Index, Count
    Set ArrayList_GetRange = ret
End Function

Private Function ArrayList_IndexOf(Value As Variant, Optional ByVal StartIndex As Long, Optional ByRef Count As Variant, Optional ByVal comparer As IComparer) As Long
    Dim cnt As Long
    
    VerifyVersion
    
    If IsMissing(Count) Then
        cnt = mCount
    Else
        cnt = Count
    End If
    
    If StartIndex < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "startindex", StartIndex)
    If cnt < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "Count", cnt)
    If StartIndex + cnt > mCount Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "Count")
    Dim i As Long
    i = mList.IndexOf(Value, mIndex + StartIndex, cnt, comparer)
    If i >= 0 Then i = i - mIndex
    ArrayList_IndexOf = i
End Function

Private Sub ArrayList_Insert(ByVal Index As Long, Value As Variant)
    VerifyVersion
    If Index < 0 Or Index > mCount Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Index), "index", Index)
    mList.Insert mIndex + Index, Value
    mCount = mCount + 1
    mVersion = mVersion + 1
End Sub

Private Sub ArrayList_InsertRange(ByVal Index As Long, c As Variant)
    Dim Size As Long
    VerifyVersion
    If Index < 0 Or Index > mCount Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Index), "index", Index)
    Size = mList.GetCollectionAndSize(c)
    mList.InsertRange mIndex + Index, c
    mCount = mCount + Size
    mVersion = mVersion + 1
End Sub

Private Property Get ArrayList_IsFixedSize() As Boolean
    ArrayList_IsFixedSize = mList.IsFixedSize
End Property

Private Property Get ArrayList_IsReadOnly() As Boolean
    ArrayList_IsReadOnly = mList.IsReadOnly
End Property

Private Property Set ArrayList_Item(ByVal Index As Long, RHS As Variant)
    VerifyVersion
    If Index < 0 Or Index >= mCount Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Index), "index", Index)
    Set mList(mIndex + Index) = RHS
    mVersion = mVersion + 1
End Property

Private Property Let ArrayList_Item(ByVal Index As Long, RHS As Variant)
    VerifyVersion
    If Index < 0 Or Index >= mCount Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Index), "index", Index)
    mList(mIndex + Index) = RHS
    mVersion = mVersion + 1
End Property

Private Property Get ArrayList_Item(ByVal Index As Long) As Variant
    VerifyVersion
    If Index < 0 Or Index >= mCount Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Index), "index", Index)
    Helper.MoveVariant ArrayList_Item, mList(mIndex + Index)
End Property

Private Function ArrayList_LastIndexOf(Value As Variant, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant, Optional ByVal comparer As IComparer) As Long
    Dim idx As Long
    Dim cnt As Long
    
    VerifyVersion
    
    If IsMissing(StartIndex) Then
        idx = mCount - 1
    Else
        idx = StartIndex
    End If
    If IsMissing(Count) Then
        cnt = mCount
    Else
        cnt = Count
    End If
    
    If idx - cnt + 1 < 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "count")
    Dim i As Long
    i = mList.LastIndexOf(Value, mIndex + idx, cnt, comparer)
    If i >= 0 Then i = i - mIndex
    ArrayList_LastIndexOf = i
End Function

Private Function ArrayList_NewEnum() As stdole.IUnknown
    VerifyVersion
    Set ArrayList_NewEnum = CreateEnumerator(ArrayList_GetEnumerator)
End Function

Private Sub ArrayList_Remove(Value As Variant, Optional ByVal comparer As IComparer)
    VerifyVersion
    Dim i As Long
    i = mList.IndexOf(Value, mIndex, mCount, comparer)
    If i >= 0 Then
        mList.RemoveAt i
        mCount = mCount - 1
        mVersion = mVersion + 1
    End If
End Sub

Private Sub ArrayList_RemoveAt(ByVal Index As Long)
    VerifyVersion
    If Index < 0 Or Index >= mCount Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Index), "index", Index)
    mList.RemoveAt mIndex + Index
    mCount = mCount - 1
    mVersion = mVersion + 1
End Sub

Private Sub ArrayList_RemoveRange(ByVal Index As Long, ByVal Count As Long)
    VerifyVersion
    If Index < 0 Or Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), IIf(Index < 0, "index", "count"), IIf(Index < 0, Index, Count))
    If Index + Count > mCount Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "count")
    mList.RemoveRange mIndex + Index, Count
    mCount = mCount - Count
    mVersion = mVersion + 1
End Sub

Private Sub ArrayList_Reverse(Optional ByVal Index As Long, Optional ByRef Count As Variant)
    Dim cnt As Long
    
    VerifyVersion
    
    If IsMissing(Count) Then
        cnt = mCount
    Else
        cnt = Count
    End If
    
    If Index < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "Index", Index)
    If Index + cnt > mCount Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "Count")
    mList.Reverse mIndex + Index, cnt
    mVersion = mVersion + 1
End Sub

Private Sub ArrayList_SetRange(ByVal Index As Long, c As Variant)
    Dim Size As Long
    
    VerifyVersion
    If Index < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "Iindex", Index)
    Size = mList.GetCollectionAndSize(c)
    If Index + Size > mCount Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LargerThanCollection), "index", Index)
    mList.SetRange mIndex + Index, c
    mVersion = mVersion + 1
End Sub

Private Sub ArrayList_Sort(Optional ByVal StartIndex As Long, Optional ByRef Count As Variant, Optional ByVal comparer As IComparer)
    Dim cnt As Long
    
    VerifyVersion
    
    If IsMissing(Count) Then
        cnt = mCount
    Else
        cnt = Count
    End If
    
    If StartIndex < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "startindex", StartIndex)
    If cnt < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "Count", Count)
    If StartIndex + cnt > mCount Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "count")
    mList.Sort mIndex + StartIndex, cnt, comparer
End Sub

Private Function ArrayList_ToArray(Optional ByVal arraytype As ciArrayTypes = 12&) As Variant
    VerifyVersion
    VariantType(ArrayList_ToArray) = arraytype Or vbArray
    MemLong(VarPtr(ArrayList_ToArray) + VARIANTDATA_OFFSET) = SafeArrayCreateVector(arraytype, 0, mCount)
    mList.CopyToEx mIndex, ArrayList_ToArray, 0, mCount
End Function

Private Function ArrayList_ToString() As String
    ArrayList_ToString = cObject_ToString
End Function

Private Sub ArrayList_TrimToSize()
    VerifyVersion
    mList.TrimToSize
    mVersion = mVersion + 1
End Sub

Private Property Get ArrayList_Version() As Long
    ArrayList_Version = mVersion
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
Private Function cObject_Equals(Value As Variant) As Boolean
    If IsObject(Value) Then
        cObject_Equals = (Value Is Me)
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
Private Function IList_Add(Value As Variant) As Long
    IList_Add = ArrayList_Add(Value)
End Function

Private Sub IList_Clear()
    ArrayList_Clear
End Sub

Private Function IList_Contains(Value As Variant, Optional ByVal comparer As IComparer) As Boolean
    IList_Contains = ArrayList_Contains(Value, comparer)
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

Private Function IList_IndexOf(Value As Variant, Optional ByVal comparer As IComparer) As Long
    IList_IndexOf = ArrayList_IndexOf(Value, , , comparer)
End Function

Private Sub IList_Insert(ByVal Index As Long, Value As Variant)
    ArrayList_Insert Index, Value
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

Private Sub IList_Remove(Value As Variant, Optional ByVal comparer As IComparer)
    ArrayList_Remove Value, comparer
End Sub

Private Sub IList_RemoveAt(ByVal Index As Long)
    ArrayList_RemoveAt Index
End Sub


