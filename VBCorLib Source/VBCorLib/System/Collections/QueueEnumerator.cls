VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QueueEnumerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) s2004 Kelly Ethridge
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
'    Module: QueueEnumerator
'

''
' Class derived from cObject.
'

Option Explicit
Implements cObject
Implements IEnumerator
Implements ICloneable

Private mQueue As Queue
Private mItems() As Variant
Private mHead As Long
Private mCapacity As Long
Private mIndex As Long
Private mEndIndex As Long
Private mVersion As Long


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = CreateEnumerator(Me)
End Function

Friend Sub Init(ByVal q As Queue, ByVal head As Long, ByVal ItemsSAPtr As Long)
    Set mQueue = q
    mHead = head
    SAPtr(mItems) = ItemsSAPtr
    mCapacity = UBound(mItems) + 1
    mEndIndex = q.Count - 1
    mVersion = q.Version
    mIndex = -1
End Sub

Friend Sub CloneHelper(ByVal q As Queue, ByVal head As Long, ByVal ItemsSAPtr As Long, ByVal index As Long)
    Init q, head, ItemsSAPtr
    mIndex = index
End Sub


Private Sub Class_Terminate()
    SAPtr(mItems) = 0
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
'   ICloneable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ICloneable_Clone() As Object
    Dim ret As New QueueEnumerator
    ret.CloneHelper mQueue, mHead, SAPtr(mItems), mIndex
    Set ICloneable_Clone = ret
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEnumerator Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get IEnumerator_Current() As Variant
    If mQueue.Version <> mVersion Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_VersionError))
    If mIndex < 0 Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumNotStarted))
    If mIndex > mEndIndex Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumFinished))
    
    VariantCopy IEnumerator_Current, mItems((mHead + mIndex) Mod mCapacity)
End Property

Private Function IEnumerator_MoveNext() As Boolean
    If mQueue.Version <> mVersion Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_VersionError))
    If mIndex > mEndIndex Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumFinished))
    
    If mIndex = mEndIndex Then
        mIndex = mEndIndex + 1
        Exit Function
    End If
    mIndex = mIndex + 1
    IEnumerator_MoveNext = True
End Function

Private Sub IEnumerator_Reset()
    mIndex = -1
End Sub
