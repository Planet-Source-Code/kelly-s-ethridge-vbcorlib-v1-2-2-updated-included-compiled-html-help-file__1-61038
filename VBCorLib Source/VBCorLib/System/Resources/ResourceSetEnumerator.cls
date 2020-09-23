VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResourceSetEnumerator"
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
'    Module: ResourceSetEnumerator
'

Option Explicit
Implements IDictionaryEnumerator
Implements IEnumerator

Private mResourceSet As ResourceSet
Private mResources As Hashtable
Private mIterator As IDictionaryEnumerator
Private mCurrentResType As IDictionaryEnumerator




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interfac
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal ResSet As ResourceSet)
    Set mResourceSet = ResSet
    Set mResources = ResSet.GetResourceSet
    Set mIterator = mResources.GetEnumerator
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IDictionaryEnumerator Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get IDictionaryEnumerator_Current() As Variant
    Set IDictionaryEnumerator_Current = IDictionaryEnumerator_Entry
End Property

Private Property Get IDictionaryEnumerator_Entry() As DictionaryEntry
    Dim Res As Win32Resource
    
    If mIterator Is Nothing Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumFinished))
    If mCurrentResType Is Nothing Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumNotStarted))
        
    Set Res = mCurrentResType.value
    Set IDictionaryEnumerator_Entry = Cor.NewDictionaryEntry(Res.Ordinal, Res)
End Property

Private Property Get IDictionaryEnumerator_Key() As Variant
    Dim Res As Win32Resource
    
    If mIterator Is Nothing Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumFinished))
    If mCurrentResType Is Nothing Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumNotStarted))
    
    Set Res = mCurrentResType.value
    IDictionaryEnumerator_Key = Res.Ordinal
End Property

Private Function IDictionaryEnumerator_MoveNext() As Boolean
    Dim ht As Hashtable
    
    If mIterator Is Nothing Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumFinished))

    Do
        If mCurrentResType Is Nothing Then
            If Not mIterator.MoveNext Then
                Set mIterator = Nothing
                Exit Function
            End If
            Set ht = mIterator.value
            Set mCurrentResType = ht.GetEnumerator
        End If
        
        If mCurrentResType.MoveNext Then Exit Do
        Set mCurrentResType = Nothing
    Loop
    
    IDictionaryEnumerator_MoveNext = True
End Function

Private Sub IDictionaryEnumerator_Reset()
    mIterator.Reset
    Set mCurrentResType = Nothing
End Sub

Private Property Get IDictionaryEnumerator_value() As Variant
    Dim Res As Win32Resource
    
    Set Res = mCurrentResType.value
    Helper.MoveVariant IDictionaryEnumerator_value, Res.value
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
