Attribute VB_Name = "modArrayHelpers"
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
'    Module: modArrayHelpers
'
Option Explicit
Public Const SIZEOF_SAFEARRAY               As Long = 16
Public Const SIZEOF_SAFEARRAYBOUND          As Long = 8
Public Const SIZEOF_SAFEARRAY1D             As Long = SIZEOF_SAFEARRAY + SIZEOF_SAFEARRAYBOUND
Public Const SIZEOF_GUID                    As Long = 16
Public Const SIZEOF_GUIDSAFEARRAY1D         As Long = SIZEOF_SAFEARRAY1D + SIZEOF_GUID
Public Const SAFEARRAYDATAPOINTER_OFFSET    As Long = 12
Public Const VARIANTDATA_OFFSET             As Long = 8
Public Const VT_BYREF                       As Long = &H4000
Public Const CBELEMENTS_OFFSET              As Long = 2

Public Type SortItems
    SA As SafeArray1d
    Buffer As Long
End Type

Private mSortItems As SortItems
Private mHasSortItems As Boolean
Private mSortKeys As SortItems
Public SortComparer As IComparer



' Compare routines.
Public Function CompareLongs(ByRef x As Long, ByRef y As Long) As Long
    If x > y Then CompareLongs = 1: Exit Function
    If x < y Then CompareLongs = -1
End Function
Public Function CompareIntegers(ByRef x As Integer, ByRef y As Integer) As Long
    If x > y Then CompareIntegers = 1: Exit Function
    If x < y Then CompareIntegers = -1
End Function
Public Function CompareStrings(ByRef x As String, ByRef y As String) As Long
    If x > y Then CompareStrings = 1: Exit Function
    If x < y Then CompareStrings = -1
End Function
Public Function CompareDoubles(ByRef x As Double, ByRef y As Double) As Long
    If x > y Then CompareDoubles = 1: Exit Function
    If x < y Then CompareDoubles = -1
End Function
Public Function CompareSingles(ByRef x As Single, ByRef y As Single) As Long
    If x > y Then CompareSingles = 1: Exit Function
    If x < y Then CompareSingles = -1
End Function
Public Function CompareBytes(ByRef x As Byte, ByRef y As Byte) As Long
    If x > y Then CompareBytes = 1: Exit Function
    If x < y Then CompareBytes = -1
End Function
Public Function CompareBooleans(ByRef x As Boolean, ByRef y As Boolean) As Long
    If x > y Then CompareBooleans = 1: Exit Function
    If x < y Then CompareBooleans = -1
End Function
Public Function CompareDates(ByRef x As Date, ByRef y As Date) As Long
    CompareDates = DateDiff("s", y, x)
End Function
Public Function CompareCurrencies(ByRef x As Currency, ByRef y As Currency) As Long
    If x > y Then CompareCurrencies = 1: Exit Function
    If x < y Then CompareCurrencies = -1
End Function
Public Function CompareIComparable(ByRef x As Object, ByRef y As Variant) As Long
    Dim comparableX As IComparable
    Set comparableX = x
    CompareIComparable = comparableX.CompareTo(y)
End Function
Public Function CompareVariants(ByRef x As Variant, ByRef y As Variant) As Long
    Dim comparable As IComparable
    
    Select Case VarType(x)
        Case vbNull
            If IsNull(y) Then Exit Function
            CompareVariants = -1
            Exit Function
        Case vbEmpty
            If IsEmpty(y) Then Exit Function
            CompareVariants = -1
            Exit Function
        Case vbObject, vbDataObject
            If TypeOf x Is IComparable Then
                Set comparable = x
                CompareVariants = comparable.CompareTo(y)
                Exit Function
            End If
        Case VarType(y)
            If x = y Then Exit Function
            If x < y Then
                CompareVariants = -1
            Else
                CompareVariants = 1
            End If
            Exit Function
    End Select
    Select Case VarType(y)
        Case vbNull, vbEmpty
            CompareVariants = 1
        Case vbObject, vbDataObject
            If TypeOf y Is IComparable Then
                Set comparable = y
                CompareVariants = -comparable.CompareTo(x)
                Exit Function
            Else
                Throw Cor.NewArgumentException("Object must implement IComparable interface.")
            End If
        Case Else
            Throw Cor.NewInvalidOperationException("Specified IComparer failed.")
    End Select
End Function


' Functions used to test for equality.
Public Function EqualsLong(ByRef x As Long, ByRef y As Long) As Boolean: EqualsLong = (x = y): End Function
Public Function EqualsString(ByRef x As String, ByRef y As String) As Boolean: EqualsString = (x = y): End Function
Public Function EqualsDouble(ByRef x As Double, ByRef y As Double) As Boolean: EqualsDouble = (x = y): End Function
Public Function EqualsInteger(ByRef x As Integer, ByRef y As Integer) As Boolean: EqualsInteger = (x = y): End Function
Public Function EqualsSingle(ByRef x As Single, ByRef y As Single) As Boolean: EqualsSingle = (x = y): End Function
Public Function EqualsDate(ByRef x As Date, ByRef y As Date) As Boolean: EqualsDate = (DateDiff("s", x, y) = 0): End Function
Public Function EqualsByte(ByRef x As Byte, ByRef y As Byte) As Boolean: EqualsByte = (x = y): End Function
Public Function EqualsBoolean(ByRef x As Boolean, ByRef y As Boolean) As Boolean: EqualsBoolean = (x = y): End Function
Public Function EqualsCurrency(ByRef x As Currency, ByRef y As Currency) As Boolean: EqualsCurrency = (x = y): End Function
Public Function EqualsObject(ByRef x As Object, ByRef y As Object) As Boolean
    Dim o As cObject
    If TypeOf x Is cObject Then
        Set o = x
        EqualsObject = o.Equals(y)
    Else
        EqualsObject = x Is y
    End If
End Function

Public Function EqualsVariants(ByRef x As Variant, ByRef y As Variant) As Boolean
    Dim o As cObject
    Select Case VarType(x)
        Case vbObject
            If x Is Nothing Then
                If IsObject(y) Then
                    EqualsVariants = (y Is Nothing)
                End If
            ElseIf TypeOf x Is cObject Then
                Set o = x
                EqualsVariants = o.Equals(y)
            ElseIf IsObject(y) Then
                If y Is Nothing Then Exit Function
                If TypeOf y Is cObject Then
                    Set o = y
                    EqualsVariants = o.Equals(x)
                Else
                    EqualsVariants = (x Is y)
                End If
            End If
        Case vbNull
            EqualsVariants = IsNull(y)
        Case VarType(y)
            EqualsVariants = (x = y)
    End Select
End Function

' Functions used to assign variables to wider variables
Public Sub WidenLongToDouble(ByRef x As Double, ByRef y As Long): x = y: End Sub
Public Sub WidenLongToString(ByRef x As String, ByRef y As Long): x = y: End Sub
Public Sub WidenLongToCurrency(ByRef x As Currency, ByRef y As Long): x = y: End Sub
Public Sub WidenLongToVariant(ByRef x As Variant, ByRef y As Long): x = y: End Sub
Public Sub WidenIntegerToLong(ByRef x As Long, ByRef y As Integer): x = y: End Sub
Public Sub WidenIntegerToSingle(ByRef x As Single, ByRef y As Integer): x = y: End Sub
Public Sub WidenIntegerToDouble(ByRef x As Double, ByRef y As Integer): x = y: End Sub
Public Sub WidenIntegerToString(ByRef x As String, ByRef y As Integer): x = y: End Sub
Public Sub WidenIntegerToCurrency(ByRef x As Currency, ByRef y As Integer): x = y: End Sub
Public Sub WidenIntegerToVariant(ByRef x As Variant, ByRef y As Integer): x = y: End Sub
Public Sub WidenByteToInteger(ByRef x As Integer, ByRef y As Byte): x = y: End Sub
Public Sub WidenByteToLong(ByRef x As Long, ByRef y As Byte): x = y: End Sub
Public Sub WidenByteToSingle(ByRef x As Single, ByRef y As Byte): x = y: End Sub
Public Sub WidenByteToDouble(ByRef x As Double, ByRef y As Byte): x = y: End Sub
Public Sub WidenByteToString(ByRef x As String, ByRef y As Byte): x = y: End Sub
Public Sub WidenByteToCurrency(ByRef x As Currency, ByRef y As Byte): x = y: End Sub
Public Sub WidenByteToVariant(ByRef x As Variant, ByRef y As Byte): x = y: End Sub
Public Sub WidenSingleToDouble(ByRef x As Double, ByRef y As Single): x = y: End Sub
Public Sub WidenSingleToString(ByRef x As String, ByRef y As Single): x = y: End Sub
Public Sub WidenSingleToVariant(ByRef x As Variant, ByRef y As Single): x = y: End Sub
Public Sub WidenDateToDouble(ByRef x As Double, ByRef y As Date): x = y: End Sub
Public Sub WidenDateToString(ByRef x As String, ByRef y As Date): x = y: End Sub
Public Sub WidenDateToVariant(ByRef x As Variant, ByRef y As Date): x = y: End Sub
Public Sub WidenObjectToVariant(ByRef x As Variant, ByRef y As Object): Set x = y: End Sub
Public Sub WidenCurrencyToString(ByRef x As String, ByRef y As Currency): x = y: End Sub
Public Sub WidenCurrencyToVariant(ByRef x As Variant, ByRef y As Currency): x = y: End Sub
Public Sub WidenStringToVariant(ByRef x As Variant, ByRef y As String): x = y: End Sub

' Functions used to assign variants to narrower variables.
Public Sub NarrowVariantToLong(ByRef x As Long, ByRef y As Variant): x = y: End Sub
Public Sub NarrowVariantToInteger(ByRef x As Integer, ByRef y As Variant): x = y: End Sub
Public Sub NarrowVariantToDouble(ByRef x As Double, ByRef y As Variant): x = y: End Sub
Public Sub NarrowVariantToString(ByRef x As String, ByRef y As Variant): x = y: End Sub
Public Sub NarrowVariantToSingle(ByRef x As Single, ByRef y As Variant): x = y: End Sub
Public Sub NarrowVariantToByte(ByRef x As Byte, ByRef y As Variant): x = y: End Sub
Public Sub NarrowVariantToDate(ByRef x As Date, ByRef y As Variant): x = y: End Sub
Public Sub NarrowVariantToBoolean(ByRef x As Boolean, ByRef y As Variant): x = y: End Sub
Public Sub NarrowVariantToCurrency(ByRef x As Currency, ByRef y As Variant): x = y: End Sub
Public Sub NarrowVariantToObject(ByRef x As Object, ByRef y As Variant): Set x = y: End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Optimized sort routines. There could have been one
'   all-purpose sort routine, but it would be too slow.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetSortKeys(ByVal pSA As Long)
    CopyMemory mSortKeys.SA, ByVal pSA, SIZEOF_SAFEARRAY1D
    Select Case mSortKeys.SA.cbElements
        Case 1, 2, 4, 8, 16
        Case Else: mSortKeys.Buffer = CoTaskMemAlloc(mSortKeys.SA.cbElements)
    End Select
End Sub
Public Sub ClearSortKeys()
    If mSortKeys.Buffer Then CoTaskMemFree (mSortKeys.Buffer)
    mSortKeys.Buffer = 0
End Sub
Public Sub SetSortItems(ByVal pSA As Long)
    CopyMemory mSortItems.SA, ByVal pSA, SIZEOF_SAFEARRAY1D
    Select Case mSortItems.SA.cbElements
        Case 1, 2, 4, 8, 16
        Case Else: mSortItems.Buffer = CoTaskMemAlloc(mSortItems.SA.cbElements)
    End Select
    mHasSortItems = True
End Sub
Public Sub ClearSortItems()
    If mHasSortItems Then
        CoTaskMemFree mSortItems.Buffer
        mSortItems.Buffer = 0
        mHasSortItems = False
    End If
End Sub
Public Sub SwapSortItems(ByRef items As SortItems, ByVal i As Long, ByVal j As Long)
    With items.SA
        Select Case .cbElements
            Case 4:     Helper.Swap4 ByVal .pvData + i * 4, ByVal .pvData + j * 4
            Case 8:     Helper.Swap8 ByVal .pvData + i * 8, ByVal .pvData + j * 8
            Case 2:     Helper.Swap2 ByVal .pvData + i * 2, ByVal .pvData + j * 2
            Case 16:    Helper.Swap16 ByVal .pvData + i * 16, ByVal .pvData + j * 16
            Case 1:     Helper.Swap1 ByVal .pvData + i, ByVal .pvData + j
            Case Else
                ' primarily for UDTs
                CopyMemory ByVal items.Buffer, ByVal .pvData + i * .cbElements, .cbElements
                CopyMemory ByVal .pvData + i * .cbElements, ByVal .pvData + j * .cbElements, .cbElements
                CopyMemory ByVal .pvData + j * .cbElements, ByVal items.Buffer, .cbElements
        End Select
    End With
End Sub
Public Sub QuickSortLong(ByRef Keys() As Long, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, x As Long, t As Long
    Do While Left < Right
        i = Left: j = Right: x = Keys((i + j) \ 2)
        Do
            Do While Keys(i) < x: i = i + 1: Loop
            Do While Keys(j) > x: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then t = Keys(i): Keys(i) = Keys(j): Keys(j) = t: If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then QuickSortLong Keys, Left, j
            Left = i
        Else
            If i < Right Then QuickSortLong Keys, i, Right
            Right = j
        End If
    Loop
End Sub
Public Sub QuickSortString(ByRef Keys() As String, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, x As String
    Do While Left < Right
        i = Left: j = Right: x = STRINGREF(Keys((i + j) \ 2))
        Do
            Do While Keys(i) < x: i = i + 1: Loop
            Do While Keys(j) > x: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then Helper.Swap4 Keys(i), Keys(j): If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then QuickSortString Keys, Left, j
            Left = i
        Else
            If i < Right Then QuickSortString Keys, i, Right
            Right = j
        End If
        StringPtr(x) = 0
    Loop
End Sub
Public Sub QuickSortObject(ByRef Keys() As Object, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, x As Variant, Key As IComparable
    Do While Left < Right
        i = Left: j = Right: Set x = Keys((i + j) \ 2)
        Do
            Set Key = Keys(i): Do While Key.CompareTo(x) < 0: i = i + 1: Set Key = Keys(i): Loop
            Set Key = Keys(j): Do While Key.CompareTo(x) > 0: j = j - 1: Set Key = Keys(j): Loop
            If i > j Then Exit Do
            If i < j Then Helper.Swap4 Keys(i), Keys(j): If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then QuickSortObject Keys, Left, j
            Left = i
        Else
            If i < Right Then QuickSortObject Keys, i, Right
            Right = j
        End If
    Loop
End Sub
Public Sub QuickSortInteger(ByRef Keys() As Integer, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, x As Integer, t As Integer
    Do While Left < Right
        i = Left: j = Right: x = Keys((i + j) \ 2)
        Do
            Do While Keys(i) < x: i = i + 1: Loop
            Do While Keys(j) > x: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then t = Keys(i): Keys(i) = Keys(j): Keys(j) = t: If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then QuickSortInteger Keys, Left, j
            Left = i
        Else
            If i < Right Then QuickSortInteger Keys, i, Right
            Right = j
        End If
    Loop
End Sub
Public Sub QuickSortByte(ByRef Keys() As Byte, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, x As Byte, t As Byte
    Do While Left < Right
        i = Left: j = Right: x = Keys((i + j) \ 2)
        Do
            Do While Keys(i) < x: i = i + 1: Loop
            Do While Keys(j) > x: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then t = Keys(i): Keys(i) = Keys(j): Keys(j) = t: If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then QuickSortByte Keys, Left, j
            Left = i
        Else
            If i < Right Then QuickSortByte Keys, i, Right
            Right = j
        End If
    Loop
End Sub
Public Sub QuickSortDouble(ByRef Keys() As Double, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, x As Double, t As Double
    Do While Left < Right
        i = Left: j = Right: x = Keys((i + j) \ 2)
        Do
            Do While Keys(i) < x: i = i + 1: Loop
            Do While Keys(j) > x: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then t = Keys(i): Keys(i) = Keys(j): Keys(j) = t: If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then QuickSortDouble Keys, Left, j
            Left = i
        Else
            If i < Right Then QuickSortDouble Keys, i, Right
            Right = j
        End If
    Loop
End Sub
Public Sub QuickSortSingle(ByRef Keys() As Single, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, x As Single, t As Single
    Do While Left < Right
        i = Left: j = Right: x = Keys((i + j) \ 2)
        Do
            Do While Keys(i) < x: i = i + 1: Loop
            Do While Keys(j) > x: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then t = Keys(i): Keys(i) = Keys(j): Keys(j) = t: If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then QuickSortSingle Keys, Left, j
            Left = i
        Else
            If i < Right Then QuickSortSingle Keys, i, Right
            Right = j
        End If
    Loop
End Sub
Public Sub QuickSortCurrency(ByRef Keys() As Currency, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, x As Currency, t As Currency
    Do While Left < Right
        i = Left: j = Right: x = Keys((i + j) \ 2)
        Do
            Do While Keys(i) < x: i = i + 1: Loop
            Do While Keys(j) > x: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then t = Keys(i): Keys(i) = Keys(j): Keys(j) = t: If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then QuickSortCurrency Keys, Left, j
            Left = i
        Else
            If i < Right Then QuickSortCurrency Keys, i, Right
            Right = j
        End If
    Loop
End Sub
Public Sub QuickSortBoolean(ByRef Keys() As Boolean, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, x As Boolean, t As Boolean
    Do While Left < Right
        i = Left: j = Right: x = Keys((i + j) \ 2)
        Do
            Do While Keys(i) < x: i = i + 1: Loop
            Do While Keys(j) > x: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then t = Keys(i): Keys(i) = Keys(j): Keys(j) = t: If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then QuickSortBoolean Keys, Left, j
            Left = i
        Else
            If i < Right Then QuickSortBoolean Keys, i, Right
            Right = j
        End If
    Loop
End Sub
Public Sub QuickSortVariant(ByRef Keys() As Variant, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, x As Variant
    Do While Left < Right
        i = Left: j = Right: VariantCopyInd x, Keys((i + j) \ 2)
        Do
            Do While CompareVariants(Keys(i), x) < 0: i = i + 1: Loop
            Do While CompareVariants(Keys(j), x) > 0: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then Helper.Swap16 Keys(i), Keys(j): If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then QuickSortVariant Keys, Left, j
            Left = i
        Else
            If i < Right Then QuickSortVariant Keys, i, Right
            Right = j
        End If
    Loop
End Sub
Public Sub QuickSortGeneral(ByRef Keys As Variant, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, x As Variant
    Do While Left < Right
        i = Left: j = Right: AssignVariant x, Keys((i + j) \ 2)
        Do
            Do While SortComparer.Compare(Keys(i), x) < 0: i = i + 1: Loop
            Do While SortComparer.Compare(Keys(j), x) > 0: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then SwapSortItems mSortKeys, i, j: If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then QuickSortGeneral Keys, Left, j
            Left = i
        Else
            If i < Right Then QuickSortGeneral Keys, i, Right
            Right = j
        End If
    Loop
End Sub

Private Sub AssignVariant(ByRef dst As Variant, ByRef src As Variant)
    VariantCopyInd dst, src
End Sub

