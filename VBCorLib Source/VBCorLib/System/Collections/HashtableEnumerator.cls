VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HashtableEnumerator"
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
'    Module: HashtableEnumerator
'
''
' Class derived from cObject.
'

Option Explicit
Implements cObject
Implements IDictionaryEnumerator
Implements IEnumerator

Public Enum HTEnumTypeEnum
    etEntries
    etKeys
    etValues
End Enum

Private mList As HashTable
Private mBuckets() As Bucket
Private mEnumType As HTEnumTypeEnum
Private mIndex As Long
Private mVersion As Long
Private mCount As Long
Private mEntries As Long
Private mCapacity As Long


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = CreateEnumerator(Me)
End Function

Friend Sub Init(ByVal list As HashTable, ByVal BucketsSAPtr As Long, ByVal HTEnumType As HTEnumTypeEnum)
    Set mList = list
    SAPtr(mBuckets) = BucketsSAPtr
    mEnumType = HTEnumType
    mIndex = -1
    mVersion = list.Version
    mCount = list.Count
    mCapacity = UBound(mBuckets) + 1
End Sub

Private Sub Class_Terminate()
    SAPtr(mBuckets) = 0
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
'   IDictionaryEnumerator Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get IDictionaryEnumerator_Current() As Variant
    If mList.Version <> mVersion Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_VersionError))
    If mIndex < 0 Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumNotStarted))
    If mEntries > mCount Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumFinished))
    
    Select Case mEnumType
        Case etEntries
            Set IDictionaryEnumerator_Current = Cor.NewDictionaryEntry(mBuckets(mIndex).Key, mBuckets(mIndex).value)
        Case etKeys
            VariantCopy IDictionaryEnumerator_Current, mBuckets(mIndex).Key
        Case etValues
            VariantCopy IDictionaryEnumerator_Current, mBuckets(mIndex).value
    End Select
End Property

Private Property Get IDictionaryEnumerator_Entry() As DictionaryEntry
    If mList.Version <> mVersion Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_VersionError))
    If mIndex < 0 Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumNotStarted))
    If mEntries > mCount Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumFinished))
    
    Set IDictionaryEnumerator_Entry = Cor.NewDictionaryEntry(mBuckets(mIndex).Key, mBuckets(mIndex).value)
End Property

Private Property Get IDictionaryEnumerator_Key() As Variant
    If mList.Version <> mVersion Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_VersionError))
    If mIndex < 0 Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumNotStarted))
    If mEntries > mCount Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumFinished))

    VariantCopy IDictionaryEnumerator_Key, mBuckets(mIndex).Key
End Property

Private Function IDictionaryEnumerator_MoveNext() As Boolean
    If mList.Version <> mVersion Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_VersionError))
    If mEntries > mCount Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumFinished))
    
    Do
        mIndex = mIndex + 1
        If mIndex = mCapacity Then
            mEntries = mCount + 1
            Exit Function
        End If
    Loop While mBuckets(mIndex).State <> bsOccupied
    mEntries = mEntries + 1
    IDictionaryEnumerator_MoveNext = True
End Function

Private Sub IDictionaryEnumerator_Reset()
    mIndex = -1
    mEntries = 0
End Sub

Private Property Get IDictionaryEnumerator_value() As Variant
    If mList.Version <> mVersion Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_VersionError))
    If mIndex < 0 Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumNotStarted))
    If mEntries > mCount Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumFinished))

    VariantCopy IDictionaryEnumerator_value, mBuckets(mIndex).value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEnumerator Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get IEnumerator_Current() As Variant
    Helper.MoveVariant IEnumerator_Current, IDictionaryEnumerator_Current
End Property

Private Function IEnumerator_MoveNext() As Boolean
    IEnumerator_MoveNext = IDictionaryEnumerator_MoveNext
End Function

Private Sub IEnumerator_Reset()
    IDictionaryEnumerator_Reset
End Sub
