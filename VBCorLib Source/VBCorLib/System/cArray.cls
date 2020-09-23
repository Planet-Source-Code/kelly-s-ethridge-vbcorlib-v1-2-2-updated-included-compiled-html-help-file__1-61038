VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
'    Module: cArray
'

''
' Provides methods for manipulating, searching and sorting arrays.
'
' @remarks This class cannot be directly instantiated. To access the methods
' use the class name directly.
' <pre>
' Debug.Print cArray.GetLength(MyArray)
' </pre>
'
Option Explicit

Private Const PVDATA_OFFSET         As Long = 12
Private Const CBELEMENTS_OFFSET     As Long = 4
Private Const FFEATURES_OFFSET      As Long = 2
Private Const LBOUND_OFFSET         As Long = 20


' extended safearray descriptor the includes the 16 bytes
' preceding the array descriptor. The 16 bytes contains
' a GUID for objects and user-defined types. This is to
' help in the correct copying of those types of elements.
Private Type GuidSafeArray1d
    Guid As VBGUID
    SA As SafeArray1d
End Type

''
' Represents the available array types that can be created using
' CreateInstance. The values corrispond the vbVarType enum.
'
' @param ciLong Same as vbLong
' @param ciInteger Same as vbInteger
' @param ciByte Same as vbByte
' @param ciDouble Same as vbDouble
' @param ciSingle Same as vbSingle
' @param ciBoolean Same as vbBoolean
' @param ciDate Same as vbDate
' @param ciString Same as vbString
' @param ciCurrency Same as vbCurrency
' @param ciVariant Same as vbVariant
' @param ciObject Same as vbObject
'
Public Enum ciArrayTypes
    ciLong = vbLong
    ciInteger = vbInteger
    ciByte = vbByte
    ciDouble = vbDouble
    ciSingle = vbSingle
    ciBoolean = vbBoolean
    ciDate = vbDate
    ciString = vbString
    ciCurrency = vbCurrency
    ciVariant = vbVariant
    ciObject = vbObject
End Enum

' for internal use only.
Private Enum CopyTypes
    CopyError
    CopyValue
    CopyReference
    CopyValueCast
    CopyObjectCast
    CopyCustom
End Enum


Private mBinarySearchDelegator As FunctionDelegator
Private mBinarySearchComparer As ITwoArgReturnLong

Private mIndexOfDelegator As FunctionDelegator
Private mIndexOfComparer As ITwoRefReturnBool

Private mAssignDelegator As FunctionDelegator
Private mAssigner As ITwoArgReturnVoid

Private mSortDelegator As FunctionDelegator
Private mSorter As ISortRoutine

Private mFindDelegator As FunctionDelegator
Private mFinder As IOneArgReturnBool

Private mCallbackSorter As CallbackSorter


''
' Uses a callback method to check if an element in an array exists.
'
' @param arr The array to check if an element exists according to the callback criteria.
' @param AddressOfMatch The AddressOf the callback method used to determine if
' an element matches the specified criteria.
' @return Indication of the element existing in the array.
'
Public Function Exists(ByRef arr As Variant, ByVal AddressOfMatch As Long) As Boolean
    Exists = (FindIndex(arr, AddressOfMatch) >= LBound(arr))
End Function

''
' Uses a callback method to search an array for the first element that matches the criteria.
'
' @param arr The array to search.
' @param AddressOfMatch The callback method address used to determine if an element
' matches the specified criteria.
' @return The element value if a match is found, otherwise a default for that element type.
' @include "..\Includes\cArray.Find.txt"
Public Function Find(ByRef arr As Variant, ByVal AddressOfMatch As Long) As Variant
    Dim i As Long
    
    i = FindIndex(arr, AddressOfMatch)
    If i >= LBound(arr) Then
        VariantCopyInd Find, arr(i)
    Else
        AssignDefault arr, Find
    End If
End Function

''
' Finds the index of the first array element that matches the criteria
' specified by the callback function.
'
' @param arr The array to be searched.
' @param AddressOfMatch The address of the callback function used to match the array elements.
' @param StartIndex The starting element of the array to beging the matching process.
' @param Count The number of elements to include in the array search.
' @return The index of the matched element, or LowerBound - 1 if none is found.
' @include "..\Includes\cArray.Find.txt"
Public Function FindIndex(ByRef arr As Variant, ByVal AddressOfMatch As Long, Optional ByVal StartIndex As Variant, Optional ByVal Count As Variant) As Long
    Dim si As Long
    Dim ct As Long
    Dim i As Long
    Dim pSA As Long
    Dim ElemSize As Long
    Dim PtrElem As Long
    Dim lb As Long
    
    pSA = GetValidArrayPointer(arr)
    lb = LBound(arr)
    
    If IsMissing(StartIndex) Then si = lb Else si = StartIndex
    If IsMissing(Count) Then ct = GetLength(arr) Else ct = Count
    
    If si < lb Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "StartIndex", si)
    If ct < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "Count", ct)
    If si + ct - 1 > UBound(arr) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "Count")
    
    ElemSize = MemLong(pSA + 4)
    PtrElem = MemLong(pSA + 12) + (si - lb) * ElemSize
    mFindDelegator.pfn = AddressOfMatch
    
    For i = 1 To ct
        If mFinder.Call(PtrElem) = True Then
            FindIndex = lb + i - 1
            Exit Function
        End If
        PtrElem = PtrElem + ElemSize
    Next i
    FindIndex = lb - 1
End Function

''
' Finds all matching elements in an array, returning an array of the matched elements.
'
' @param arr The array to find all matching elements.
' @param AddressOfMatch The callback address of the matching criteria method.
' @return An array of all matched elements.
' @include "..\Includes\cArray.Find.txt"
Public Function FindAll(ByRef arr As Variant, ByVal AddressOfMatch As Long) As Variant
    Dim Indexes() As Long
    Dim pSA As Long
    Dim ElemSize As Long
    Dim PtrElem As Long
    Dim i As Long
    Dim j As Long
    Dim ct As Long
    Dim ret As Variant
    Dim lb As Long
    
    
    pSA = GetValidArrayPointer(arr)
    ElemSize = MemLong(pSA + CBELEMENTS_OFFSET)
    PtrElem = MemLong(pSA + PVDATA_OFFSET)
    mFindDelegator.pfn = AddressOfMatch
    ct = GetLength(arr)
    lb = LBound(arr)
    
    ' just redim to the maximum number of indexes possible.
    ' may have to revisit this later to be more memory effecient.
    ReDim Indexes(0 To ct - 1)
    
    ' iterate each array element, passing the memory address
    ' to the user supplied AddressOf callback. This will appear
    ' as a ByRef parameter of the array's datatype in the callback.
    For i = lb To lb + ct - 1
        If mFinder.Call(PtrElem) = True Then
            Indexes(j) = i
            j = j + 1
        End If
        PtrElem = PtrElem + ElemSize
    Next i
    
    Dim vt As Long
    vt = VarType(arr) And &HFF
    Select Case vt
        Case vbObject
            ' as an array objects, we create a new array using the
            ' same Guid value as that of the original array.
            VariantType(ret) = vbObject Or vbArray
            MemLong(VarPtr(ret) + VARIANTDATA_OFFSET) = SafeArrayCreateVectorEx(vbObject, 0, j, pSA - SIZEOF_GUID)
            For i = 0 To j - 1
                Set ret(i) = arr(Indexes(i))
            Next i
            
        Case vbUserDefinedType
            ' oh fun in creating an array of UDTs.
            Dim rec As IRecordInfo
            Dim PtrDstElem As Long
            Dim pvData  As Long
            
            Set rec = SafeArrayGetRecordInfo(pSA)
            pvData = MemLong(pSA + PVDATA_OFFSET)
            
            pSA = SafeArrayCreateVectorEx(vbUserDefinedType, 0, j, ObjPtr(rec))
            PtrDstElem = MemLong(pSA + PVDATA_OFFSET)
            
            For i = 0 To j - 1
                rec.RecordCopy pvData + (Indexes(i) - lb) * ElemSize, PtrDstElem
                PtrDstElem = PtrDstElem + ElemSize
            Next i
            
            VariantType(ret) = vbUserDefinedType Or vbArray
            MemLong(VarPtr(ret) + VARIANTDATA_OFFSET) = pSA
            
        Case Else
            ret = CreateInstance(vt, j)
            For i = 0 To j - 1
                ret(i) = arr(Indexes(i))
            Next i
    End Select
    
    Helper.MoveVariant FindAll, ret
End Function

''
' Finds the last occurence of a matched element in the array.
'
' @param arr The array to search.
' @param AddressOfMatch A callback address of the matching criteria method.
' @return The last element in the array to match the criteria, or the
' default for the array type if no elements matched.
' @include "..\Includes\cArray.Find.txt"
Public Function FindLast(ByRef arr As Variant, ByVal AddressOfMatch As Long) As Variant
    Dim i As Long
    
    i = FindLastIndex(arr, AddressOfMatch)
    If i >= LBound(arr) Then
        VariantCopyInd FindLast, arr(i)
    Else
        AssignDefault arr, FindLast
    End If
End Function

''
' Finds the index of the last occurence of a matched element in the array.
'
' @param arr The array to search.
' @param AddressOfMatch The callback address of the matching criteria method.
' @param StartIndex The starting index in the array to begin the backward search.
' @param Count The number of elements to search.
' @return The index of the last occurence of the matched element.
' @include "..\Includes\cArray.Find.txt"
Public Function FindLastIndex(ByRef arr As Variant, ByVal AddressOfMatch As Long, Optional ByVal StartIndex As Variant, Optional ByVal Count As Variant) As Long
    Dim si As Long
    Dim ct As Long
    Dim i As Long
    Dim pSA As Long
    Dim ElemSize As Long
    Dim PtrElem As Long
    Dim lb As Long
    Dim ub As Long
    
    pSA = GetValidArrayPointer(arr)
    lb = LBound(arr)
    ub = UBound(arr)
    
    If IsMissing(StartIndex) Then si = UBound(arr) Else si = StartIndex
    If IsMissing(Count) Then ct = GetLength(arr) Else ct = Count
    
    If si > ub Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_UBound), "StartIndex", si)
    If ct < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "Count", ct)
    If si - ct + 1 < LBound(arr) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "Count")
    
    ElemSize = MemLong(pSA + CBELEMENTS_OFFSET)
    PtrElem = MemLong(pSA + PVDATA_OFFSET) + (si - lb) * ElemSize
    mFindDelegator.pfn = AddressOfMatch
    
    For i = ct To 1 Step -1
        If mFinder.Call(PtrElem) = True Then
            FindLastIndex = lb + i - 1
            Exit Function
        End If
        PtrElem = PtrElem - ElemSize
    Next i
    FindLastIndex = lb - 1
End Function

''
' Returns a boolean indicating if all elements in the array matched the criteria.
'
' @param arr The array to compare against the criteria.
' @param AddressOfMatch The callback address used to match the criteria.
' @Return Indication of the entire array matching the criteria or not.
' @include "..\Includes\cArray.Find.txt"
Public Function TrueForAll(ByRef arr As Variant, ByVal AddressOfMatch As Long) As Boolean
    Dim i As Long
    Dim pSA As Long
    Dim ElemSize As Long
    Dim PtrElem As Long
    Dim lb As Long
    
    pSA = GetValidArrayPointer(arr)
    
    ElemSize = MemLong(pSA + CBELEMENTS_OFFSET)
    PtrElem = MemLong(pSA + PVDATA_OFFSET)
    mFindDelegator.pfn = AddressOfMatch
    
    For i = LBound(arr) To UBound(arr)
        If mFinder.Call(PtrElem) = False Then Exit Function
        PtrElem = PtrElem + ElemSize
    Next i
    TrueForAll = True
End Function


''
' Performs a binary search on a given array. A subportion of the array can
' be searched using the startindex and length parameters. A custom user
' comparer can optionally be supplied to perform special comparisons
' between elements in the array.
'
' @param arr The array to search for a specific value.
' @param value The value to search for in the array.
' @param vStartIndex The starting index in the array to begin searching.
' @param vLength The number of elements to search, starting at startindex.
' @param comparer A user supplied object to compare elements within the array.
' @return The index at which the value was found.
' @remarks If the returned value is less than the lower bound of the array, then
' the value indicates where the value would have been found in the array. The
' following is how to convert the return value to the array index:<br><br>
' Lowerbound of 0:  return = Not return<br>
' Other lowerbound: return = (Not (return - lowerbound)) + lowerbound<br>
' @include "..\Includes\cArray_BinarySearch.txt"
Public Function BinarySearch(ByRef arr As Variant, ByRef Value As Variant, Optional ByVal vStartIndex As Variant, Optional ByVal vLength As Variant, Optional ByVal comparer As IComparer) As Long
    Dim md As Long
    Dim lo As Long
    Dim hi As Long
    Dim lb As Long
    Dim pSA As Long
    Dim StartIndex As Long
    Dim Length As Long
    
    pSA = GetValidArrayPointer(arr)
    If SafeArrayGetDim(pSA) > 1 Then Throw Cor.NewRankException(Environment.GetResourceString(Rank_MultiDimension))
    
    lb = LBound(arr)
    If IsMissing(vStartIndex) Then StartIndex = lb Else StartIndex = vStartIndex
    If IsMissing(vLength) Then Length = cArray.GetLength(arr) Else Length = vLength
    
    hi = StartIndex + Length - 1
    lo = StartIndex
    
    ' do some argument checking
    If StartIndex < lb Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "StartIndex", StartIndex)
    If Length < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "Length", Length)
    If hi > UBound(arr) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "Length")
        
    On Error GoTo errTrap
    If comparer Is Nothing Then
        If TrySZBinarySearch(pSA, Value, StartIndex, Length, BinarySearch) Then Exit Function
        Set comparer = VBCorLib.comparer.Default
    End If
    
    ' perform a search using comparer.
    Do While lo <= hi
        md = (lo + hi) \ 2
        Select Case comparer.Compare(arr(md), Value)
            Case 0
                BinarySearch = md
                Exit Function
            Case Is > 0
                hi = md - 1
            Case Else
                lo = md + 1
        End Select
    Loop
    
    BinarySearch = (Not (lo - lb)) + lb
    Exit Function
errTrap:
    Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_Comparer_Arg, Err.Description))
End Function

''
'   Clears a portion of the elements in an array.
'
' @param arr The array to clear elements from.
' @param startindex The starting element to being clearing.
' @param length The number of elements to be cleared.
'
Public Sub Clear(ByRef arr As Variant, ByVal StartIndex As Long, ByVal Length As Long)
    Dim pSA As Long
    Dim lb As Long
    Dim IIDOffset As Long
    
    pSA = GetValidArrayPointer(arr)
    lb = LBound(arr)
    If StartIndex < lb Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "startindex", StartIndex)
    If Length < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "Length", Length)
    If StartIndex + Length - 1 > UBound(arr) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset))
    If SafeArrayGetDim(pSA) > 1 Then _
        Throw Cor.NewRankException(Environment.GetResourceString(Rank_MultiDimension))
    
    If Length = 0 Then Exit Sub
    
    Dim SA As GuidSafeArray1d
'    Select Case VarType(arr) And &HFF
'        Case vbObject
'            CopyMemory SA, ByVal pSA - SIZEOF_GUID, SIZEOF_GUIDSAFEARRAY1D
'        Case vbUserDefinedType
'            CopyMemory ByVal VarPtr(SA.SA) - 4, ByVal pSA - 4, SIZEOF_SAFEARRAY1D + 4
'        Case Else
'            CopyMemory SA.SA, ByVal pSA, SIZEOF_SAFEARRAY1D
'    End Select
    Select Case VarType(arr) And &HFF
        Case vbObject: IIDOffset = SIZEOF_GUID
        Case vbUserDefinedType: IIDOffset = 4
    End Select
    
    CopyMemory ByVal VarPtr(SA.SA) - IIDOffset, ByVal pSA - IIDOffset, Len(SA.SA) + IIDOffset
    
    With SA.SA
        .fFeatures = .fFeatures Or FADF_FIXEDSIZE Or FADF_STATIC Or FADF_AUTO
        .pvData = .pvData + (StartIndex - lb) * .cbElements
        .cElements = Length
    End With
    SafeArrayDestroyData VarPtr(SA.SA)
End Sub

''
' Copies one Array to another Array and performs type casting as necessary.
'
' @param sourcearray The array from which to copy the elements.
' @param destinationarray The array in which to place the elements.
' @param length The number of elements to copy.
' @param AddressOfCopier A callback address to a method used to assign
' elements from the source array to the destination array.
' @remarks The AddressOfCopier method signature must be two parameters
' of ByRef for the datatypes of each of the arrays. The destination
' array element is the first parameter.
' <pre>
' Public Sub CopyCallback(ByRef dst As String, ByRef src As MyClass)
'     dst = src.Name
' End Sub
' </pre>
Public Sub Copy(ByRef sourceArray As Variant, ByRef destinationarray As Variant, ByVal Length As Long, Optional ByVal AddressOfCopier As Long = 0)
    CopyEx sourceArray, LBound(sourceArray), destinationarray, LBound(destinationarray), Length, AddressOfCopier
End Sub

''
' Copies a section of one Array to another Array and performs type casting as necessary.
'
' @param sourcearray The array from which to copy the elements.
' @param sourceindex The starting element in the source array to begin copying from.
' @param destinationarray The array in which to place the elements.
' @param destinationindex The starting element in the destination array to place elements.
' @param length The number of elements to copy.
' @param AddressOfCopier A callback address to a method used to assign
' elements from the source array to the destination array.
' @remarks The AddressOfCopier method signature must be two parameters
' of ByRef for the datatypes of each of the arrays. The destination
' array element is the first parameter.
' <pre>
' Public Sub CopyCallback(ByRef dst As String, ByRef src As MyClass)
'     dst = src.Name
' End Sub
' </pre>
Public Sub CopyEx(ByRef sourceArray As Variant, ByVal sourceindex As Long, ByRef destinationarray As Variant, ByVal destinationindex As Long, ByVal Length As Long, Optional ByVal AddressOfCopier As Long = 0)
    Dim srcpSA As Long
    Dim dstpSA As Long
    Dim srcType As VbVarType
    Dim dstType As VbVarType
    Dim IsRefType As Boolean
    Dim CopyType As CopyTypes
    
    srcpSA = GetValidArrayPointer(sourceArray)
    dstpSA = GetValidArrayPointer(destinationarray)
    
    If SafeArrayGetDim(srcpSA) > 1 Or SafeArrayGetDim(dstpSA) > 1 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Rank_MultiDimension))
    If sourceindex < LBound(sourceArray) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "sourceindex", sourceindex)
    If destinationindex < LBound(destinationarray) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "destinationindex", destinationindex)
    If Length < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "Length", Length)
    If (sourceindex + Length - 1 > UBound(sourceArray)) Or (destinationindex + Length - 1 > UBound(destinationarray)) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "Length")
    
    If Length = 0 Then Exit Sub
    
    srcType = VarType(sourceArray) And &HFF
    dstType = VarType(destinationarray) And &HFF
    
    If AddressOfCopier = 0 Then
        If srcType = dstType Then
            Select Case srcType
                Case vbObject
                    If IsEqualGUID(ByVal srcpSA - SIZEOF_GUID, ByVal dstpSA - SIZEOF_GUID) <> BOOL_FALSE Then
                        CopyType = CopyReference
                    Else
                        CopyType = CopyObjectCast
                    End If
                    
                Case vbString
                    CopyType = CopyReference
                    
                Case vbUserDefinedType
                    Dim SrcRec As IRecordInfo
                    Dim DstRec As IRecordInfo
                    
                    Set SrcRec = SafeArrayGetRecordInfo(srcpSA)
                    Set DstRec = SafeArrayGetRecordInfo(dstpSA)
                    
                    If SrcRec.IsMatchingType(DstRec) Then
                        CopyType = CopyReference
                    Else
                        CopyType = CopyError
                    End If
                    
                Case Else
                    CopyType = CopyValue
            End Select
        Else
            CopyType = CopyValueCast
        End If
    Else
        CopyType = CopyCustom
    End If
    
    Select Case CopyType
        Case CopyValue
            CopyValueTypes srcpSA, sourceindex, dstpSA, destinationindex, Length
            
        Case CopyReference
            CopyReferenceTypes srcpSA, sourceindex, dstpSA, destinationindex, Length
            
        Case CopyValueCast
            Dim pfn As Long
            pfn = GetAssigningFunction(srcType, dstType)
            If pfn = 0 Then _
                Throw Cor.NewArrayTypeMismatchException(Environment.GetResourceString(ArrayTypeMismatch_Incompatible))
            
            mAssignDelegator.pfn = pfn
            CopyWithCast srcpSA, sourceindex, dstpSA, destinationindex, Length

        Case CopyObjectCast
            CopyObjectTypesCast srcpSA, sourceindex, dstpSA, destinationindex, Length
            
        Case CopyCustom
            mAssignDelegator.pfn = AddressOfCopier
            CopyWithCast srcpSA, sourceindex, dstpSA, destinationindex, Length
            
        Case CopyError
            Throw Cor.NewArrayTypeMismatchException(Environment.GetResourceString(ArrayTypeMismatch_Incompatible))
    End Select
End Sub

''
' Initializes a new array of the specified type in up to 3 dimensions.
'
' @param ArrayType The data type the array represents.
' @param length1 The number of elements in the first dimension.
' @param length2 The number of elements in the second dimension.
' @param length3 The number of elements in the third dimension.
' @returns An array of the specified type with the specified dimensions.
' @remarks If a length is zero, then an empty array is returned, not a Null array.
' If one length is defined as zero, then all other defined lengths must be zero.
' This does not mean that lenghts 2 and 3 need to be set to zero if only an empty
' one-dimensional array is being created.
Public Function CreateInstance(ByVal arraytype As ciArrayTypes, Optional ByVal length1 As Long = 0, Optional ByRef length2 As Variant, Optional ByRef length3 As Variant) As Variant
    Dim lengths(2) As Long
    Dim bounds() As SafeArrayBound
    Dim cDims As Long
    
    cDims = 1
    If length1 < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "length1", length1)
    lengths(0) = length1
    
    If Not IsMissing(length2) Then
        cDims = 2
        lengths(1) = length2
        If lengths(1) < 0 Then _
            Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "length2", length2)
        If lengths(1) = 0 And lengths(0) <> 0 Then _
            Throw Cor.NewArgumentException("If defining one dimension as 0, then all other defined dimensions must be 0.", "length2")
    End If
    If Not IsMissing(length3) Then
        If cDims <> 2 Then Throw Cor.NewArgumentException("Cannot specifiy dimension without dimension 2.", "length3")
        cDims = 3
        lengths(2) = length3
        If lengths(2) < 0 Then _
            Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "length3", length3)
        If lengths(2) = 0 And lengths(1) <> 0 Then _
            Throw Cor.NewArgumentException("If defining one dimension as 0, then all other defined dimensions must be 0.", "length2")
    End If
    
    ReDim bounds(0 To cDims - 1)
    Dim i As Long
    For i = 0 To cDims - 1
        bounds(i).cElements = lengths(i)
        bounds(i).lLbound = 0
    Next i
    
    VariantType(CreateInstance) = arraytype Or vbArray
    MemLong(VarPtr(CreateInstance) + VARIANTDATA_OFFSET) = SafeArrayCreate(arraytype, cDims, bounds(0))
End Function

''
' Creates an array of values.
'
' @param arraytype The target datatype of the resultant array and the values to be converted to.
' @param arrayelements The values to be put into the new array.
' @return An array of the specified datatype and containing the array elements with conversion.
Public Function NewArray(ByVal arraytype As ciArrayTypes, ParamArray arrayelements() As Variant) As Variant
    Dim ub As Long
    Dim args() As Variant
    
    ub = UBound(arrayelements)
    If ub >= 0 Then
        VariantType(NewArray) = arraytype Or vbArray
        MemLong(VarPtr(NewArray) + VARIANTDATA_OFFSET) = SafeArrayCreateVector(arraytype, 0, ub + 1)
        Helper.Swap4 ByVal ArrPtr(args), ByVal Helper.DerefEBP(16)
        CopyEx args, 0, NewArray, 0, ub + 1
    End If
End Function

''
' Returns the number of elements in the specified dimension. If 0 is
' specified, this returns the total number of elements in the array.
'
' @param arr The array to retrieve the number of elements from.
' @param dimension (Optional) the specific dimension to get the number of elements.
' @return The number of elements in the specified dimension or the entire array.
Public Function GetLength(ByRef arr As Variant, Optional ByVal dimension As Long = 0) As Long
    Dim pSA As Long
    Dim rank As Long
    Dim i As Long
    
    pSA = GetValidArrayPointer(arr)
    rank = MemWord(pSA)
    If dimension > rank Then _
        Throw Cor.NewIndexOutOfRangeException(Environment.GetResourceString(IndexOutOfRange_Dimension))
    
    pSA = pSA + SIZEOF_SAFEARRAY
    If dimension < 1 Then
        ' optimize for 99% of calls
        If rank = 1 Then
            GetLength = MemLong(pSA)
            Exit Function
        End If
        GetLength = 1
        For i = pSA To pSA + (SIZEOF_SAFEARRAYBOUND * (rank - 1)) Step SIZEOF_SAFEARRAYBOUND
            GetLength = GetLength * MemLong(i)
        Next i
    Else
        pSA = pSA + (rank - dimension) * SIZEOF_SAFEARRAYBOUND
        GetLength = MemLong(pSA)
    End If
End Function

''
' Returns the number of dimensions in the array.
'
' @param arr The array to retrieve the number of dimensions from.
' @return The number of dimension in the array.
Public Function GetRank(ByRef arr As Variant) As Long
    GetRank = MemWord(GetValidArrayPointer(arr))
End Function

''
' Returns the index of the first occurrence of a value in a one-dimensional Array or in a portion of the Array.
'
' @param arr The array to search.
' @param value The value to search for in the array.
' @param vstartindex The index to the element in the array to begin search from.
' @param vcount The number of elements to search, starting from startindex.
' @param comparer An optional comparer, primarily used for User-Defined types.
' @return Value indicating the index the value was found. If a value of less that
' the lower-bounds is returned, then the value was not found in the array.
Public Function IndexOf(ByRef arr As Variant, ByRef Value As Variant, Optional ByRef vStartIndex As Variant, Optional ByRef vCount As Variant, Optional ByVal comparer As IComparer) As Long
    Dim pSA As Long
    Dim StartIndex As Long
    Dim Count As Long
    
    pSA = GetValidArrayPointer(arr)
    If IsMissing(vStartIndex) Then StartIndex = LBound(arr) Else StartIndex = vStartIndex
    If IsMissing(vCount) Then Count = cArray.GetLength(arr) Else Count = vCount

    If StartIndex < LBound(arr) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "startindex", StartIndex)
    If Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "Count", Count)
    If StartIndex + Count - 1 > UBound(arr) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "Count")
    If SafeArrayGetDim(pSA) > 1 Then _
        Throw Cor.NewRankException(Environment.GetResourceString(Rank_MultiDimension))

    On Error GoTo errTrap
    If comparer Is Nothing Then
        ' if no comparer was supplied, then try some optimized routines.
        If TrySZIndexOf(pSA, Value, StartIndex, Count, IndexOf) Then Exit Function
        Set comparer = VBCorLib.comparer.Default
    End If
    ' perform a search using comparer.
    Do While Count > 0
        If comparer.Compare(arr(StartIndex), Value) = 0 Then
            IndexOf = StartIndex
            Exit Function
        End If
        StartIndex = StartIndex + 1
        Count = Count - 1
    Loop
    IndexOf = LBound(arr) - 1
    Exit Function
errTrap:
    Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_Comparer_Arg, Err.Description))
End Function

''
' Returns if an array variable is uninitialized.
'
' @param arr The array to test for initialization.
' @return Value indicating if the array is initialized.
Public Function IsNull(ByRef arr As Variant) As Boolean
    IsNull = (GetArrayPointer(arr) = 0)
End Function

''
' Returns the index of the last occurrence of a value in a one-dimensional Array or in a portion of the Array.
'
' @param arr The array to search.
' @param value The value to search for in the array.
' @param vstartindex The index to the element in the array to begin search from.
' @param vcount The number of elements to search, starting from startindex.
' @param comparer An optional comparer, primarily used for User-Defined types.
' @return Value indicating the index the value was found. If a value of less that
' the lower-bounds is returned, then the value was not found in the array.
Public Function LastIndexOf(ByRef arr As Variant, ByRef Value As Variant, Optional ByVal vStartIndex As Variant, Optional ByVal vCount As Variant, Optional ByVal comparer As IComparer) As Long
    Dim pSA As Long
    Dim StartIndex As Long
    Dim Count As Long
    
    pSA = GetValidArrayPointer(arr)
    If IsMissing(vStartIndex) Then StartIndex = UBound(arr) Else StartIndex = vStartIndex
    If IsMissing(vCount) Then Count = cArray.GetLength(arr) Else Count = vCount
    
    If StartIndex > UBound(arr) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_UBound), "startindex", StartIndex)
    If Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "Count", Count)
    If StartIndex - Count + 1 < LBound(arr) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "Count")
    If SafeArrayGetDim(pSA) > 1 Then _
        Throw Cor.NewRankException(Environment.GetResourceString(Rank_MultiDimension))

    On Error GoTo errTrap
    If comparer Is Nothing Then
        ' if no comparer was supplied, then try some optimized routines.
        If TrySZLastIndexOf(pSA, Value, StartIndex, Count, LastIndexOf) Then Exit Function
        Set comparer = VBCorLib.comparer.Default
    End If
    ' perform a search using comparer.
    Do While Count > 0
        If comparer.Compare(arr(StartIndex), Value) = 0 Then
            LastIndexOf = StartIndex
            Exit Function
        End If
        StartIndex = StartIndex - 1
        Count = Count - 1
    Loop
    LastIndexOf = LBound(arr) - 1
    Exit Function
errTrap:
    Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_Comparer_Arg, Err.Description))
End Function

''
' Reverses the elements in a subportion of an array.
'
' @param arr The array to reverse.
' @param vStartIndex The starting element to begin reversing elements.
' @param vLength The number of elements to reverse in the array.
'
Public Sub Reverse(ByRef arr As Variant, Optional ByVal vStartIndex As Variant, Optional ByVal vLength As Variant)
    Dim pSA As Long
    Dim items As SortItems
    Dim i As Long
    Dim j As Long
    Dim swapper As ISwap
    Dim StartIndex As Long
    Dim Length As Long
    
    pSA = GetValidArrayPointer(arr)
    
    If IsMissing(vStartIndex) Then StartIndex = LBound(arr) Else StartIndex = vStartIndex
    If IsMissing(vLength) Then Length = cArray.GetLength(arr) Else Length = vLength
    
    If StartIndex < LBound(arr) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "startindex", StartIndex)
    If Length < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "Length", Length)
    If StartIndex + Length - 1 > UBound(arr) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "length")
    If SafeArrayGetDim(pSA) > 1 Then _
        Throw Cor.NewRankException(Environment.GetResourceString(Rank_MultiDimension))
    
    CopyMemory items.SA, ByVal pSA, SIZEOF_SAFEARRAY1D
    With items.SA
        Select Case .cbElements
            Case 1, 2, 4, 8, 16
                Set swapper = GetSwapper(.cbElements)
            Case Else
                items.Buffer = CoTaskMemAlloc(.cbElements)
        End Select
        
        i = .pvData + StartIndex * .cbElements
        j = i + Length * .cbElements - .cbElements
        Do While i < j
            If Not swapper Is Nothing Then
                swapper.Call 0, ByVal i, ByVal j    ' the first argument is just a filler.
            Else
                CopyMemory ByVal items.Buffer, ByVal i, .cbElements
                CopyMemory ByVal i, ByVal j, .cbElements
                CopyMemory ByVal j, ByVal items.Buffer, .cbElements
            End If
            i = i + .cbElements
            j = j - .cbElements
        Loop
    End With
    If items.Buffer Then CoTaskMemFree items.Buffer
End Sub

''
' Sorts an entire array. An optionally supplied comparer object can be
' used to compare special elements, such as userdefined values.
'
' @param arr The array to search for a value.
' @param comparer (Optional) An IComparer object to perform custom comparisons.
' @include "..\Includes\cArray_Sort.txt"
Public Sub Sort(ByRef arr As Variant, Optional ByVal comparer As Variant)
    InternalSort arr, Empty, LBound(arr), UBound(arr) - LBound(arr) + 1, comparer
End Sub

''
' Sorts an array, or subportion, given a startindex and length. An optionally
' supplied comparer object can be used to compare special elements, such as
' userdefined values.
'
' @param arr The array to search for a value.
' @param startindex The starting index in the array to start sorting.
' @param length The number of elements to be sorted.
' @param comparer (Optional) An IComparer object to perform custom comparisons.
'
Public Sub SortEx(ByRef arr As Variant, ByVal StartIndex As Long, ByVal Length As Long, Optional ByVal comparer As Variant)
    InternalSort arr, Empty, StartIndex, Length, comparer
End Sub

''
' Sorts an entire array based on an array of keys. An optionally supplied
' comparer object can be used to compare special elements, such as userdefined values.
'
' @param keys An array the sorting is based on.
' @param items An array that is sorted based on the sorting of keys.
' @param comparer (Optional) An IComparer object to perform custom comparisons.
' @include "..\Includes\cArray.SortKey.txt"
Public Sub SortKey(ByRef Keys As Variant, ByRef items As Variant, Optional ByVal comparer As Variant)
    InternalSort Keys, items, LBound(Keys), UBound(Keys) - LBound(Keys) + 1, comparer
End Sub

''
' Sorts an array based on an array of keys. A subportion of the array can
' be sorted given a startindex and length. An optionally, supplied comparer
' object can be used to compare special elements, such as userdefined values.
'
' @param keys An array the sorting is based on.
' @param items An array that is sorted based on the sorting of keys.
' @param startindex The starting index in the array to start sorting.
' @param length The number of elements to be sorted.
' @param comparer (Optional) An IComparer object to perform custom comparisons.
'
Public Sub SortKeyEx(ByRef Keys As Variant, ByRef items As Variant, ByVal StartIndex As Long, ByVal Length As Long, Optional ByVal comparer As Variant)
    InternalSort Keys, items, StartIndex, Length, comparer
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mBinarySearchComparer = InitDelegator(mBinarySearchDelegator)
    Set mIndexOfComparer = InitDelegator(mIndexOfDelegator)
    Set mAssigner = InitDelegator(mAssignDelegator)
    Set mSorter = InitDelegator(mSortDelegator)
    Set mFinder = InitDelegator(mFindDelegator)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CallbackSorter() As CallbackSorter
    If mCallbackSorter Is Nothing Then Set mCallbackSorter = New CallbackSorter
    Set CallbackSorter = mCallbackSorter
End Function

Private Sub AssignDefault(ByRef arr As Variant, ByRef RetVal As Variant)
    Select Case VarType(arr) And &HFF
        Case vbLong:        RetVal = 0&
        Case vbInteger:     RetVal = 0
        Case vbByte:        RetVal = CByte(0)
        Case vbString:      RetVal = vbNullString
        Case vbObject:      Set RetVal = Nothing
        Case vbDouble:      RetVal = 0#
        Case vbSingle:      RetVal = 0!
        Case vbDate:        RetVal = #12:00:00 AM#
        Case vbCurrency:    RetVal = 0@
        Case vbBoolean:     RetVal = False
        Case vbVariant:     RetVal = Empty
        Case vbUserDefinedType
            Dim rec As IRecordInfo
            Set rec = SafeArrayGetRecordInfo(GetArrayPointer(arr))
            VariantType(RetVal) = vbUserDefinedType
            MemLong(VarPtr(RetVal) + 8) = rec.RecordCreate
            MemLong(VarPtr(RetVal) + 12) = ObjectPtr(rec)
            ObjectPtr(rec) = 0
    End Select
End Sub

Private Sub CopyValueTypes(ByVal srcpSA As Long, ByVal sourceindex As Long, ByVal dstpSA As Long, ByVal destinationindex As Long, ByVal Length As Long)
    Dim pSrcElem As Long
    Dim pDstElem As Long
    Dim ElemSize As Long
    
    ElemSize = MemLong(srcpSA + CBELEMENTS_OFFSET)
    pSrcElem = MemLong(srcpSA + PVDATA_OFFSET) + (sourceindex - MemLong(srcpSA + LBOUND_OFFSET)) * ElemSize
    pDstElem = MemLong(dstpSA + PVDATA_OFFSET) + (destinationindex - MemLong(dstpSA + LBOUND_OFFSET)) * ElemSize
    
    CopyMemory ByVal pDstElem, ByVal pSrcElem, Length * ElemSize
End Sub

Private Sub CopyReferenceTypes(ByVal srcpSA As Long, ByVal sourceindex As Long, ByVal dstpSA As Long, ByVal destinationindex As Long, ByVal Length As Long)
    Dim srcSA As GuidSafeArray1d
    Dim dstSA As GuidSafeArray1d
    Dim IIDOffset As Long
    
    Select Case SafeArrayGetVartype(srcpSA)
        Case vbObject:              IIDOffset = SIZEOF_GUID
        Case vbUserDefinedType:     IIDOffset = 4
    End Select
    
    CopyMemory ByVal VarPtr(srcSA.SA) - IIDOffset, ByVal srcpSA - IIDOffset, LenB(srcSA.SA) + IIDOffset
    CopyMemory ByVal VarPtr(dstSA.SA) - IIDOffset, ByVal dstpSA - IIDOffset, LenB(dstSA.SA) + IIDOffset
    
    If srcpSA = dstpSA Then
        ' test for overlapping source and destination portion of the same array
        If destinationindex > sourceindex And destinationindex < sourceindex + Length Then
            Dim ElemSize As Long
            Dim pvData As Long
            Dim lb As Long
            Dim ClearSA As GuidSafeArray1d
            
            ClearSA = srcSA
            With ClearSA.SA
                .pvData = .pvData + (sourceindex + Length) * .cbElements
                .cElements = destinationindex - sourceindex
            End With
            SafeArrayDestroyData VarPtr(ClearSA.SA)
            
            With srcSA.SA
                ElemSize = .cbElements
                pvData = .pvData
                lb = .lLbound
            End With

            CopyMemory ByVal pvData + (destinationindex - lb) * ElemSize, ByVal pvData + (sourceindex - lb) * ElemSize, Length * ElemSize

            ' if we have moved reference type values, then we need to
            ' fill in the uncovered portion with the original values
            ' with copies. First we have to zero out the uncovered portion
            ' so the copy won't release any references, since we merely
            ' moved the references without making copies to objects or strings.
            ' Once the uncovered portion is set up, we can copy backwards using
            ' the standard reference copy method without stepping on ourselves.
            ZeroMemory ByVal pvData + (sourceindex - lb) * ElemSize, (destinationindex - sourceindex) * ElemSize
            
            Length = destinationindex - sourceindex
            Dim t As Long
            t = sourceindex
            sourceindex = destinationindex
            destinationindex = t
        End If
    End If
    
    SetCopyArrayDescriptor srcSA.SA, sourceindex, Length
    SetCopyArrayDescriptor dstSA.SA, destinationindex, Length
    
    SafeArrayCopyData VarPtr(srcSA.SA), VarPtr(dstSA.SA)
End Sub

Private Sub CopyWithCast(ByVal srcpSA As Long, ByVal sourceindex As Long, ByVal dstpSA As Long, ByVal destinationindex As Long, ByVal Length As Long)
    Dim src As Long
    Dim dst As Long
    Dim srcLb As Long
    Dim dstLb As Long
    Dim srcSize As Long
    Dim dstSize As Long
    
    srcSize = SafeArrayGetElemsize(srcpSA)
    dstSize = SafeArrayGetElemsize(dstpSA)
    srcLb = SafeArrayGetLBound(srcpSA, 1)
    dstLb = SafeArrayGetLBound(dstpSA, 1)
    
    src = MemLong(srcpSA + PVDATA_OFFSET) + (sourceindex - srcLb) * srcSize
    dst = MemLong(dstpSA + PVDATA_OFFSET) + (destinationindex - dstLb) * dstSize
    
    Do While Length > 0
        mAssigner.Call dst, src
        src = src + srcSize
        dst = dst + dstSize
        Length = Length - 1
    Loop
End Sub

Private Sub CopyObjectTypesCast(ByVal srcpSA As Long, ByVal sourceindex As Long, ByVal dstpSA As Long, ByVal destinationindex As Long, ByVal Length As Long)
    Dim src As Long
    Dim dst As Long
    Dim srcLb As Long
    Dim dstLb As Long
    Dim srcObj As IVBUnknown
    Dim dstObj As IVBUnknown
    Dim pDstGuid As Long
    Dim pObj As Long
    
    srcLb = SafeArrayGetLBound(srcpSA, 1)
    dstLb = SafeArrayGetLBound(dstpSA, 1)
    
    src = MemLong(srcpSA + PVDATA_OFFSET) + (sourceindex - srcLb) * 4
    dst = MemLong(dstpSA + PVDATA_OFFSET) + (destinationindex - dstLb) * 4
    pDstGuid = dstpSA - SIZEOF_GUID
    
    Do While Length > 0
        ObjectPtr(srcObj) = MemLong(src)
        pObj = 0
        If Not srcObj Is Nothing Then
            If srcObj.QueryInterface(ByVal pDstGuid, pObj) = E_NOINTERFACE Then
                ObjectPtr(srcObj) = 0
                ObjectPtr(dstObj) = 0
                Throw Cor.NewArrayTypeMismatchException(Environment.GetResourceString(ArrayTypeMismatch_Incompatible))
            End If
        End If
        
        ObjectPtr(dstObj) = MemLong(dst)
        If Not dstObj Is Nothing Then dstObj.Release
        MemLong(dst) = pObj
        
        src = src + 4
        dst = dst + 4
        Length = Length - 1
    Loop
    
    ObjectPtr(srcObj) = 0
    ObjectPtr(dstObj) = 0
End Sub

Private Function GetAssigningFunction(ByVal srcType As VbVarType, ByVal dstType As VbVarType) As Long
    Dim ret As Long
    Select Case srcType
        Case vbLong
            Select Case dstType
                Case vbDouble:      ret = FuncAddr(AddressOf WidenLongToDouble)
                Case vbString:      ret = FuncAddr(AddressOf WidenLongToString)
                Case vbCurrency:    ret = FuncAddr(AddressOf WidenLongToCurrency)
                Case vbVariant:     ret = FuncAddr(AddressOf WidenLongToVariant)
            End Select
        Case vbInteger, vbBoolean
            Select Case dstType
                Case vbLong:        ret = FuncAddr(AddressOf WidenIntegerToLong)
                Case vbDouble:      ret = FuncAddr(AddressOf WidenIntegerToDouble)
                Case vbString:      ret = FuncAddr(AddressOf WidenIntegerToString)
                Case vbSingle:      ret = FuncAddr(AddressOf WidenIntegerToSingle)
                Case vbCurrency:    ret = FuncAddr(AddressOf WidenIntegerToCurrency)
                Case vbVariant:     ret = FuncAddr(AddressOf WidenIntegerToVariant)
            End Select
        Case vbByte
            Select Case dstType
                Case vbLong:        ret = FuncAddr(AddressOf WidenByteToLong)
                Case vbInteger:     ret = FuncAddr(AddressOf WidenByteToInteger)
                Case vbDouble:      ret = FuncAddr(AddressOf WidenByteToDouble)
                Case vbString:      ret = FuncAddr(AddressOf WidenByteToString)
                Case vbSingle:      ret = FuncAddr(AddressOf WidenByteToSingle)
                Case vbCurrency:    ret = FuncAddr(AddressOf WidenByteToCurrency)
                Case vbVariant:     ret = FuncAddr(AddressOf WidenByteToVariant)
            End Select
        Case vbSingle
            Select Case dstType
                Case vbDouble:      ret = FuncAddr(AddressOf WidenSingleToDouble)
                Case vbString:      ret = FuncAddr(AddressOf WidenSingleToString)
                Case vbVariant:     ret = FuncAddr(AddressOf WidenSingleToVariant)
            End Select
        Case vbDate
            Select Case dstType
                Case vbDouble:      ret = FuncAddr(AddressOf WidenDateToDouble)
                Case vbString:      ret = FuncAddr(AddressOf WidenDateToString)
                Case vbVariant:     ret = FuncAddr(AddressOf WidenDateToVariant)
            End Select
        Case vbObject, vbDataObject
            Select Case dstType
                Case vbVariant:     ret = FuncAddr(AddressOf WidenObjectToVariant)
            End Select
        Case vbCurrency
            Select Case dstType
                Case vbString:      ret = FuncAddr(AddressOf WidenCurrencyToString)
                Case vbVariant:     ret = FuncAddr(AddressOf WidenCurrencyToVariant)
            End Select
        Case vbString
            Select Case dstType
                Case vbVariant:     ret = FuncAddr(AddressOf WidenStringToVariant)
            End Select
        Case vbVariant
            Select Case dstType
                Case vbLong:                    ret = FuncAddr(AddressOf NarrowVariantToLong)
                Case vbInteger:                 ret = FuncAddr(AddressOf NarrowVariantToInteger)
                Case vbString:                  ret = FuncAddr(AddressOf NarrowVariantToString)
                Case vbDouble:                  ret = FuncAddr(AddressOf NarrowVariantToDouble)
                Case vbObject, vbDataObject:    ret = FuncAddr(AddressOf NarrowVariantToObject)
                Case vbSingle:                  ret = FuncAddr(AddressOf NarrowVariantToSingle)
                Case vbByte:                    ret = FuncAddr(AddressOf NarrowVariantToByte)
                Case vbDate:                    ret = FuncAddr(AddressOf NarrowVariantToDate)
                Case vbBoolean:                 ret = FuncAddr(AddressOf NarrowVariantToBoolean)
                Case vbCurrency:                ret = FuncAddr(AddressOf NarrowVariantToCurrency)
            End Select
    End Select
    GetAssigningFunction = ret
End Function

''
' Perform the sorting. If no comparer is supplied, then optimized sorting routines
' will be attempted, otherwise, the general routine will be used with the comparer.
Private Sub InternalSort(ByRef Keys As Variant, ByRef items As Variant, ByVal Index As Long, ByVal Length As Long, ByRef comparer As Variant)
    Dim pKeysSA As Long
    Dim pItemsSA As Long
    
    pKeysSA = GetValidArrayPointer(Keys)
    If SafeArrayGetDim(pKeysSA) > 1 Then _
        Throw Cor.NewRankException(Environment.GetResourceString(Rank_MultiDimension))
    If Index < LBound(Keys) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "index", Index)
    If Length < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "Length", Length)
    If Index + Length - 1 > UBound(Keys) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "length")
    
    If Not IsEmpty(items) Then
        pItemsSA = GetValidArrayPointer(items)
        If SafeArrayGetDim(pItemsSA) > 1 Then _
            Throw Cor.NewRankException(Environment.GetResourceString(Rank_MultiDimension))
        If (LBound(Keys) <> LBound(items)) Or _
           (UBound(Keys) <> UBound(items)) Then _
            Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_MatchingBounds))
    End If
    
    Do
        Select Case VarType(comparer)
            Case vbObject
                SetSortKeys pKeysSA
                If pItemsSA <> 0 Then SetSortItems pItemsSA
                If comparer Is Nothing Then
                    Set SortComparer = VBCorLib.comparer.Default
                Else
                    Set SortComparer = comparer
                End If
                QuickSortGeneral Keys, Index, Index + Length - 1
                Exit Do
                
            Case vbLong
                Index = Index - LBound(Keys)
                CallbackSorter.Sort pKeysSA, pItemsSA, Index, Index + Length - 1, comparer
                Exit Sub
                
            Case vbError
                If pItemsSA <> 0 Then SetSortItems pItemsSA
                Index = Index - LBound(Keys)
                If TrySZSort(pKeysSA, Index, Index + Length - 1) Then Exit Do
                Set comparer = VBCorLib.comparer.Default
            
            Case Else
                Throw Cor.NewArgumentException("Invalid Comparer.", "Comparer")
        End Select
    Loop
    
errTrap:
    ClearSortKeys
    ClearSortItems
    Set SortComparer = Nothing
    
    Dim ex As Exception
    If Catch(ex, Err) Then
        If ex.HResult = 13 Then
            Throw Cor.NewInvalidCastException(Environment.GetResourceString(ArrayTypeMismatch_Compare))
        Else
            Throw ex
        End If
    End If
End Sub

Private Sub SetCopyArrayDescriptor(ByRef desc As SafeArray1d, ByVal Index As Long, ByVal Length As Long)
    With desc
        .pvData = .pvData + (Index - .lLbound) * .cbElements
        .cElements = Length
    End With
End Sub

''
' This is an optimized search routine that uses a function pointer
' to call a specific comparison routine.
Private Function SZBinarySearch(ByVal pSA As Long, ByVal pValue As Long, ByVal StartIndex As Long, ByVal Length As Long, ByVal pfn As Long) As Long
    Dim Size As Long
    Dim pvData As Long
    Dim lo As Long
    Dim hi As Long
    Dim md As Long
    Dim lb As Long
    
    mBinarySearchDelegator.pfn = pfn
    Size = SafeArrayGetElemsize(pSA)
    pvData = MemLong(pSA + 12)
    
    lb = SafeArrayGetLBound(pSA, 1)
    lo = StartIndex - lb
    hi = lo + Length - 1
    
    Do While lo <= hi
        md = (lo + hi) \ 2
        Select Case mBinarySearchComparer.Call(pvData + md * Size, pValue)
            Case 0
                SZBinarySearch = md + lb
                Exit Function
            Case Is > 0
                hi = md - 1
            Case Else
                lo = md + 1
        End Select
    Loop
    
    SZBinarySearch = (Not lo) + lb
End Function

Private Function SZIndexOf(ByVal pSA As Long, ByVal pValue As Long, ByVal StartIndex As Long, ByVal Count As Long, ByVal pfn As Long) As Long
    Dim Size As Long
    Dim pvData As Long
    Dim lb As Long
    
    mIndexOfDelegator.pfn = pfn
    Size = SafeArrayGetElemsize(pSA)
    pvData = MemLong(pSA + 12)
    lb = SafeArrayGetLBound(pSA, 1)
    
    StartIndex = (StartIndex - lb)
    Do While Count > 0
        If mIndexOfComparer.Call(pvData + StartIndex * Size, pValue) Then
            SZIndexOf = StartIndex + lb
            Exit Function
        End If
        Count = Count - 1
        StartIndex = StartIndex + 1
    Loop

    SZIndexOf = lb - 1
End Function

Private Function SZLastIndexOf(ByVal pSA As Long, ByVal pValue As Long, ByVal StartIndex As Long, ByVal Count As Long, ByVal pfn As Long) As Long
    Dim Size As Long
    Dim pvData As Long
    Dim lb As Long
    
    mIndexOfDelegator.pfn = pfn
    Size = SafeArrayGetElemsize(pSA)
    pvData = MemLong(pSA + 12)
    lb = SafeArrayGetLBound(pSA, 1)
    
    StartIndex = (StartIndex - lb)
    Do While Count > 0
        If mIndexOfComparer.Call(pvData + StartIndex * Size, pValue) Then
            SZLastIndexOf = StartIndex + lb
            Exit Function
        End If
        Count = Count - 1
        StartIndex = StartIndex - 1
    Loop

    SZLastIndexOf = lb - 1
End Function

''
' Attempt to use a specialized search for a specific data type.
Private Function TrySZBinarySearch(ByVal pSA As Long, ByRef Value As Variant, ByVal StartIndex As Long, ByVal Length As Long, ByRef RetVal As Long) As Boolean
    Select Case SafeArrayGetVartype(pSA)
        Case vbLong:                    RetVal = SZBinarySearch(pSA, VarPtr(CLng(Value)), StartIndex, Length, AddressOf CompareLongs)
        Case vbString:                  RetVal = SZBinarySearch(pSA, VarPtr(StrPtr(Value)), StartIndex, Length, AddressOf CompareStrings)
        Case vbDouble:                  RetVal = SZBinarySearch(pSA, VarPtr(CDbl(Value)), StartIndex, Length, AddressOf CompareDoubles)
        Case vbObject, vbDataObject:    RetVal = SZBinarySearch(pSA, VarPtr(Value), StartIndex, Length, AddressOf CompareIComparable)
        Case vbVariant:                 RetVal = SZBinarySearch(pSA, VarPtr(Value), StartIndex, Length, AddressOf CompareVariants)
        Case vbInteger:                 RetVal = SZBinarySearch(pSA, VarPtr(CInt(Value)), StartIndex, Length, AddressOf CompareIntegers)
        Case vbSingle:                  RetVal = SZBinarySearch(pSA, VarPtr(CSng(Value)), StartIndex, Length, AddressOf CompareSingles)
        Case vbCurrency:                RetVal = SZBinarySearch(pSA, VarPtr(CCur(Value)), StartIndex, Length, AddressOf CompareCurrencies)
        Case vbDate:                    RetVal = SZBinarySearch(pSA, VarPtr(CDate(Value)), StartIndex, Length, AddressOf CompareDates)
        Case vbBoolean:                 RetVal = SZBinarySearch(pSA, VarPtr(CBool(Value)), StartIndex, Length, AddressOf CompareBooleans)
        Case vbByte:                    RetVal = SZBinarySearch(pSA, VarPtr(CByte(Value)), StartIndex, Length, AddressOf CompareBytes)
        Case Else: Exit Function
    End Select
    TrySZBinarySearch = True
End Function

Private Function TrySZIndexOf(ByVal pSA As Long, ByRef Value As Variant, ByVal StartIndex As Long, ByVal Count As Long, ByRef RetVal As Long) As Boolean
    Select Case SafeArrayGetVartype(pSA) And &HFF
        Case vbLong:                    RetVal = SZIndexOf(pSA, VarPtr(CLng(Value)), StartIndex, Count, AddressOf EqualsLong)
        Case vbString:                  RetVal = SZIndexOf(pSA, VarPtr(StrPtr(Value)), StartIndex, Count, AddressOf EqualsString)
        Case vbDouble:                  RetVal = SZIndexOf(pSA, VarPtr(CDbl(Value)), StartIndex, Count, AddressOf EqualsDouble)
        Case vbDate:                    RetVal = SZIndexOf(pSA, VarPtr(CDate(Value)), StartIndex, Count, AddressOf EqualsDate)
        Case vbObject, vbDataObject:    RetVal = SZIndexOf(pSA, VarPtr(ObjPtr(Value)), StartIndex, Count, AddressOf EqualsObject)
        Case vbVariant:                 RetVal = SZIndexOf(pSA, VarPtr(Value), StartIndex, Count, AddressOf EqualsVariants)
        Case vbInteger:                 RetVal = SZIndexOf(pSA, VarPtr(CInt(Value)), StartIndex, Count, AddressOf EqualsInteger)
        Case vbSingle:                  RetVal = SZIndexOf(pSA, VarPtr(CSng(Value)), StartIndex, Count, AddressOf EqualsSingle)
        Case vbByte:                    RetVal = SZIndexOf(pSA, VarPtr(CSng(Value)), StartIndex, Count, AddressOf EqualsByte)
        Case vbBoolean:                 RetVal = SZIndexOf(pSA, VarPtr(CBool(Value)), StartIndex, Count, AddressOf EqualsBoolean)
        Case vbCurrency:                RetVal = SZIndexOf(pSA, VarPtr(CCur(Value)), StartIndex, Count, AddressOf EqualsCurrency)
        Case Else: Exit Function
    End Select
    TrySZIndexOf = True
End Function

Private Function TrySZLastIndexOf(ByVal pSA As Long, ByRef Value As Variant, ByVal StartIndex As Long, ByVal Count As Long, ByRef RetVal As Long) As Boolean
    Select Case SafeArrayGetVartype(pSA) And &HFF
        Case vbLong:                    RetVal = SZLastIndexOf(pSA, VarPtr(CLng(Value)), StartIndex, Count, AddressOf EqualsLong)
        Case vbString:                  RetVal = SZLastIndexOf(pSA, VarPtr(StrPtr(Value)), StartIndex, Count, AddressOf EqualsString)
        Case vbDouble:                  RetVal = SZLastIndexOf(pSA, VarPtr(CDbl(Value)), StartIndex, Count, AddressOf EqualsDouble)
        Case vbDate:                    RetVal = SZLastIndexOf(pSA, VarPtr(CDate(Value)), StartIndex, Count, AddressOf EqualsDate)
        Case vbObject, vbDataObject:    RetVal = SZLastIndexOf(pSA, VarPtr(ObjPtr(Value)), StartIndex, Count, AddressOf EqualsObject)
        Case vbVariant:                 RetVal = SZLastIndexOf(pSA, VarPtr(Value), StartIndex, Count, AddressOf EqualsVariants)
        Case vbInteger:                 RetVal = SZLastIndexOf(pSA, VarPtr(CInt(Value)), StartIndex, Count, AddressOf EqualsInteger)
        Case vbSingle:                  RetVal = SZLastIndexOf(pSA, VarPtr(CSng(Value)), StartIndex, Count, AddressOf EqualsSingle)
        Case vbByte:                    RetVal = SZLastIndexOf(pSA, VarPtr(CSng(Value)), StartIndex, Count, AddressOf EqualsByte)
        Case vbBoolean:                 RetVal = SZLastIndexOf(pSA, VarPtr(CBool(Value)), StartIndex, Count, AddressOf EqualsBoolean)
        Case vbCurrency:                RetVal = SZLastIndexOf(pSA, VarPtr(CCur(Value)), StartIndex, Count, AddressOf EqualsCurrency)
        Case Else: Exit Function
    End Select
    TrySZLastIndexOf = True
End Function

''
'   Attempt to locate an optimized sort routine for a specific data type.
Private Function TrySZSort(ByVal pSA As Long, ByVal Left As Long, ByVal Right As Long) As Boolean
    Dim pfn As Long
    Select Case SafeArrayGetVartype(pSA) And &HFF
        Case vbLong:                    pfn = FuncAddr(AddressOf QuickSortLong)
        Case vbString:                  pfn = FuncAddr(AddressOf QuickSortString)
        Case vbDouble, vbDate:          pfn = FuncAddr(AddressOf QuickSortDouble)
        Case vbObject, vbDataObject:    pfn = FuncAddr(AddressOf QuickSortObject)
        Case vbVariant:                 pfn = FuncAddr(AddressOf QuickSortVariant)
        Case vbInteger:                 pfn = FuncAddr(AddressOf QuickSortInteger)
        Case vbSingle:                  pfn = FuncAddr(AddressOf QuickSortSingle)
        Case vbByte:                    pfn = FuncAddr(AddressOf QuickSortByte)
        Case vbCurrency:                pfn = FuncAddr(AddressOf QuickSortCurrency)
        Case vbBoolean:                 pfn = FuncAddr(AddressOf QuickSortBoolean)
        Case Else: Exit Function
    End Select
    mSortDelegator.pfn = pfn
    mSorter.Call pSA, Left, Right
    TrySZSort = True
End Function

''
' This will verify that the array pointer is not Null. If it is null
' then an ArgumentNullException will be thrown.
Friend Function GetValidArrayPointer(ByRef arr As Variant) As Long
    GetValidArrayPointer = GetArrayPointer(arr)
    If GetValidArrayPointer = vbNullPtr Then _
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Array), "arr")
End Function

Private Function GetArrayPointer(ByRef arr As Variant) As Long
    Const BYREF_ARRAY As Long = VT_BYREF Or vbArray
    
    Dim vt As Long
    
    vt = VariantType(arr)
    Select Case vt And BYREF_ARRAY
        ' we have to double deref the original array pointer because
        ' the variant held a pointer to the original array variable.
        Case BYREF_ARRAY:   GetArrayPointer = MemLong(MemLong(VarPtr(arr) + VARIANTDATA_OFFSET))
        ' we won't need to deref again if the original array was dimensioned
        ' as a variant ie:
        '    Dim arr As Variant
        '    ReDim arr(1 To 10) As Long
        '
        ' The passed in variant will be the array variable, not a ByRef
        ' pointer to the array variable.
        Case vbArray:       GetArrayPointer = MemLong(VarPtr(arr) + VARIANTDATA_OFFSET)
        ' you bad person
        Case Else:          Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_ArrayRequired), "arr")
    End Select
    
    ' HACK HACK HACK
    '
    ' When an uninitialized array of objects or UDTs is passed into a
    ' function as a ByRef Variant, the array is initialized with just the
    ' SafeArrayDescriptor, at which point, it is a valid array and can
    ' be used by UBound and LBound after the call. So, now we're just
    ' going to assume that any object or UDT array that has just the descriptor
    ' allocated was Null to begin with. That means whenever an Object or UDT
    ' array is passed to any cArray method, it will technically never
    ' be uninitialized, just zero-length.
    Select Case vt And &HFF
        Case vbObject, vbUserDefinedType: If MemLong(GetArrayPointer + PVDATA_OFFSET) = 0 Then GetArrayPointer = 0
    End Select
End Function
