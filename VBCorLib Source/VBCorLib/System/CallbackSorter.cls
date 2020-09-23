VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CallbackSorter"
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
'    Module: CallbackSorter
'

Option Explicit

Private mDelegator As FunctionDelegator
Private mComparer As ITwoRefReturnLong
Private mHasSortItems As Boolean
Private mSortItems As SortItems
Private mKeyItems As SortItems
Private mPVData As Long
Private mElemSize As Long
Private mPivotBuffer As Long

Private mInt8() As Byte
Private mInt16() As Integer
Private mInt32() As Long
Private mInt64() As Currency
Private mInt128() As Variant



Friend Sub Sort(ByVal pSAKeys As Long, ByVal pSAItems As Long, ByVal Left As Long, ByVal Right As Long, ByVal AddressOfCallback As Long)
    mDelegator.pfn = AddressOfCallback
    If pSAItems <> 0 Then
        InitSortItems pSAItems
    Else
        mHasSortItems = False
    End If
    
    Select Case SafeArrayGetElemsize(pSAKeys)
        Case 1
            SAPtr(mInt8) = pSAKeys
            SortInt8 Left, Right
            SAPtr(mInt8) = 0
        Case 2
            SAPtr(mInt16) = pSAKeys
            SortInt16 Left, Right
            SAPtr(mInt16) = 0
        Case 4
            SAPtr(mInt32) = pSAKeys
            SortInt32 Left, Right
            SAPtr(mInt32) = 0
        Case 8
            SAPtr(mInt64) = pSAKeys
            SortInt64 Left, Right
            SAPtr(mInt64) = 0
        Case 16
            SAPtr(mInt128) = pSAKeys
            SortInt128 Left, Right
            SAPtr(mInt128) = 0
        Case Else
            With mKeyItems
                CopyMemory .SA, ByVal pSAKeys, SIZEOF_SAFEARRAY1D
                mElemSize = .SA.cbElements
                mPivotBuffer = CoTaskMemAlloc(mElemSize)
                .Buffer = CoTaskMemAlloc(mElemSize)
                mPVData = .SA.pvData
                SortAny Left, Right
                CoTaskMemFree .Buffer
                CoTaskMemFree mPivotBuffer
            End With
    End Select
    
    If mSortItems.Buffer Then
        CoTaskMemFree mSortItems.Buffer
        mSortItems.Buffer = 0
    End If
End Sub

Private Sub InitSortItems(ByVal pSA As Long)
    CopyMemory mSortItems.SA, ByVal pSA, SIZEOF_SAFEARRAY1D
    Select Case mSortItems.SA.cbElements
        Case 1, 2, 4, 8, 16
        Case Else: mSortItems.Buffer = CoTaskMemAlloc(mSortItems.SA.cbElements)
    End Select
    mHasSortItems = True
End Sub

Private Sub SortInt8(ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, x As Byte, t As Byte
    Do While Left < Right
        i = Left: j = Right: x = mInt8((i + j) \ 2)
        Do
            Do While mComparer.Call(mInt8(i), x) < 0: i = i + 1: Loop
            Do While mComparer.Call(mInt8(j), x) > 0: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then t = mInt8(i): mInt8(i) = mInt8(j): mInt8(j) = t: If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then SortInt8 Left, j
            Left = i
        Else
            If i < Right Then SortInt8 i, Right
            Right = j
        End If
    Loop
End Sub

Private Sub SortInt16(ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, x As Integer, t As Integer
    Do While Left < Right
        i = Left: j = Right: x = mInt16((i + j) \ 2)
        Do
            Do While mComparer.Call(mInt16(i), x) < 0: i = i + 1: Loop
            Do While mComparer.Call(mInt16(j), x) > 0: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then t = mInt16(i): mInt16(i) = mInt16(j): mInt16(j) = t: If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then SortInt16 Left, j
            Left = i
        Else
            If i < Right Then SortInt16 i, Right
            Right = j
        End If
    Loop
End Sub

Private Sub SortInt32(ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, x As Long, t As Long
    Do While Left < Right
        i = Left: j = Right: x = mInt32((i + j) \ 2)
        Do
            Do While mComparer.Call(mInt32(i), x) < 0: i = i + 1: Loop
            Do While mComparer.Call(mInt32(j), x) > 0: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then t = mInt32(i): mInt32(i) = mInt32(j): mInt32(j) = t: If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then SortInt32 Left, j
            Left = i
        Else
            If i < Right Then SortInt32 i, Right
            Right = j
        End If
    Loop
End Sub

Private Sub SortInt64(ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, x As Currency, t As Currency
    Do While Left < Right
        i = Left: j = Right: x = mInt64((i + j) \ 2)
        Do
            Do While mComparer.Call(mInt64(i), x) < 0: i = i + 1: Loop
            Do While mComparer.Call(mInt64(j), x) > 0: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then t = mInt64(i): mInt64(i) = mInt64(j): mInt64(j) = t: If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then SortInt64 Left, j
            Left = i
        Else
            If i < Right Then SortInt64 i, Right
            Right = j
        End If
    Loop
End Sub

Private Sub SortInt128(ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, x As Variant
    Do While Left < Right
        i = Left: j = Right: CopyMemory x, mInt128((i + j) \ 2), 16
        Do
            Do While mComparer.Call(mInt128(i), x) < 0: i = i + 1: Loop
            Do While mComparer.Call(mInt128(j), x) > 0: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then Helper.Swap16 mInt128(i), mInt128(j): If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then SortInt128 Left, j
            Left = i
        Else
            If i < Right Then SortInt128 i, Right
            Right = j
        End If
        VariantType(x) = 0
    Loop
End Sub

Private Sub SortAny(ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long: Dim PivotBuffer As Long
    'PivotBuffer = CoTaskMemAlloc(mElemSize)
    Do While Left < Right
        i = Left: j = Right: CopyMemory ByVal mPivotBuffer, ByVal mPVData + ((i + j) \ 2) * mElemSize, mElemSize
        Do
            Do While mComparer.Call(ByVal mPVData + i * mElemSize, ByVal mPivotBuffer) < 0: i = i + 1: Loop
            Do While mComparer.Call(ByVal mPVData + j * mElemSize, ByVal mPivotBuffer) > 0: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then SwapSortItems mKeyItems, i, j: If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then SortAny Left, j
            Left = i
        Else
            If i < Right Then SortAny i, Right
            Right = j
        End If
    Loop
    'CoTaskMemFree PivotBuffer
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mComparer = InitDelegator(mDelegator)
End Sub
