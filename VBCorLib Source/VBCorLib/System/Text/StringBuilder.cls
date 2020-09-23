VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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
'    Module: StringBuilder
'

''
' Allows easy manipulation, concatenation and removal of strings and characters.
'
' @remarks
' This class can be used to quickly minpulate strings. It allows for easy concatenation,
' removal and modification of the current underlying string value.
'
' @see Constructors
' @see NumberFormatInfo
' @see DateTimeFormatInfo
'
Option Explicit
Implements cObject

Private Const DEF_FILLCHAR          As Integer = 32
Private Const DEF_CAPACITY          As Long = 512
Private Const COLON                 As Long = 58
Private Const OPEN_BRACE            As Long = 123
Private Const CLOSE_BRACE           As Long = 125
Private Const COMMA                 As Long = 44
Private Const MINUS                 As Long = 45

' module wide scope
Private mStringPtr As Long
Private mStringSA As SafeArray1d
Private mString() As Integer
Private mCapacity As Long
Private mLength As Long
Private mFillChar As Integer

' Replace/AppendFormat command scope
Private OldStr() As Integer
Private BufStr() As Integer
Private OldStrSA As SafeArray1d
Private BufStrSA As SafeArray1d



''
' Appends a string to the current string value.
'
' @param s String to be appended.
' @return This instance of StringBuilder.
' @remarks This version of Append passes then string in by value to allow any conversion
' to take place without having to explicitly cast the value to a string. Such things as
' numbers will want to easily be converted.
'
Public Function Append(ByVal s As String) As StringBuilder
    Set Append = AppendString(s)
End Function

''
' Appends a character a specified number of times to the end of string value.
'
' @param char The AscW character code to append.
' @param count The number of times to append the character.
' @return This instance of StringBuilder.
' @remarks The character is a unicode value from -32767 to 65535. For values that
' are greater than 32767 (&H7FFF), 65536 (&H10000) is subtracted from it to map the
' character to a 16-bit Integer without overflowing.
'
Public Function AppendChar(ByVal Char As Long, Optional ByVal Count As Long = 1) As StringBuilder
    If Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Count cannot be negative.", "Count", Count)
    
    If mLength + Count > mCapacity Then EnsureCapacity mLength + Count
    If Char < -&H7FFF& Or Char > &HFFFF& Then _
        Throw Cor.NewArgumentOutOfRangeException("Char must be between -32767 and 65535 inclusively.", "Char", Char)
    If Char > &H7FFF& Then Char = Char - &H10000
    
    ' appending 1 character at a time happens 99%
    ' of the time, so optimize for it.
    If Count = 1 Then
        mString(mLength) = Char
    Else
        Fill mLength, Char, Count
    End If
    mLength = mLength + Count
    Set AppendChar = Me
End Function

''
' Appends an array of characters to the end of the string value.
'
' @param Chars The array of characters to be appended.
' @param startindex The index from Chars to start appending.
' @param count The number of characters to append.
' @return This instance of StringBuilder.
'
Public Function AppendChars(ByRef Chars() As Integer, Optional ByVal StartIndex As Long, Optional ByRef Count As Variant) As StringBuilder
    Dim lb As Long
    Dim cnt As Long
    

    If cArray.IsNull(Chars) Then _
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Array), "chars")
    
    lb = LBound(Chars)
    If IsMissing(Count) Then
        cnt = cArray.GetLength(Chars) - (StartIndex - lb)
    Else
        cnt = Count
    End If
    
    If StartIndex < lb Then _
        Throw Cor.NewArgumentOutOfRangeException("startindex cannot be less than LowerBound.", "startindex", StartIndex)
    If StartIndex + cnt - 1 > UBound(Chars) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "Count")
    
    If mLength + cnt > mCapacity Then EnsureCapacity mLength + cnt
    CopyMemory mString(mLength), Chars(StartIndex), cnt * 2
    mLength = mLength + cnt
    Set AppendChars = Me
End Function

''
' Appends a string with formatted arguments.
'
' @param format The string to append that contains formatting information.
' @param args The arguments to be formatted into the string.
' @return This instance of StringBuilder.
' @remarks
' <p>The format string can embed indexes into the args parameter to indicate where
' a value should be placed. The way to indicate an index is through an argument index enclosed
' in braces {N}. N is the parameter index beginning with zero.</p>
' <br><p>Some arguments may handle additional formatting commands. Please refer to their
' documentation for more information.</p>
' @see NumberFormatInfo
' @see DateTimeFormatInfo
' @include "..\..\Includes\StringBuilder.AppendFormat.txt"
Public Function AppendFormat(ByRef Format As String, ParamArray args() As Variant) As StringBuilder
    Dim vArgs() As Variant
    Helper.Swap4 ByVal ArrPtr(vArgs), ByVal Helper.DerefEBP(16)
    InternalAppendFormat Nothing, Format, vArgs
    Set AppendFormat = Me
End Function

''
' Appends a string with formatted arguments using the supplied provider.
'
' @param provider A custom formatting object that formats the arguments.
' @param Format The string to append that contains formatting information.
' @param args The arguments to be formatted into the string.
' @return This instance of StringBuilder.
' @remarks
' <p>The format string can embed indexes into the args parameter to indicate where
' a value should be placed. The way to indicate an index is through an argument index enclosed
' in braces {N}. N is the parameter index beginning with zero.</p>
' <br><p>Some arguments may handle additional formatting commands. Please refer to their
' documentation for more information.</p>
' <br><p>The supplied provider can format specific arguments unknown to the system. It can also
' delegate formatting of system-known arguments to another StringBuilder (usually cString.Format)
' to provide default formatting of known types and format only specific types.
' @see NumberFormatInfo
' @see DateTimeFormatInfo
'
Public Function AppendFormatEx(ByVal provider As IFormatProvider, ByRef Format As String, ParamArray args() As Variant) As StringBuilder
    Dim vArgs() As Variant
    Helper.Swap4 ByVal ArrPtr(vArgs), ByVal Helper.DerefEBP(20)
    InternalAppendFormat provider, Format, vArgs
    Set AppendFormatEx = Me
End Function

''
' Appends a string to the current string value.
'
' @param s String to be appended.
' @param startindex The index to the first character of string to start appending.
' @param count The number of character from the string to append.
' @return This instance of StringBuilder.
' @remarks<p>This version of append uses a by-reference string value to force a string type
' to be passed in. This reduces string duplication.</p>
' <p>The startindex is zero based, not one based like the Left,Right and Mid functions.
'
Public Function AppendString(ByRef s As String, Optional ByVal StartIndex As Long, Optional ByRef Count As Variant) As StringBuilder
    Dim ls As Long
    Dim cnt As Long
    
    ls = Len(s)
    If ls > 0 Then
        If StartIndex < 0 Then _
            Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "startindex", StartIndex)
            
        If IsMissing(Count) Then
            cnt = ls - StartIndex
        Else
            cnt = Count
        End If
        
        If StartIndex + cnt > ls Then _
            Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "Count")
        
        If mLength + cnt > mCapacity Then EnsureCapacity mLength + cnt
        CopyMemory ByVal mStringPtr + mLength * 2, ByVal StrPtr(s) + StartIndex * 2, cnt * 2
        mLength = mLength + cnt
    End If
    Set AppendString = Me
End Function

''
' Appends the supplied string and a Carriage-Return/Linefeed to the end of the StringBuilder.
'
' @param s The string to append to the builder, followed by a Carraige-Return/Linefeed.
' @return A reference to this instance of StringBuilder.
'
Public Function AppendLine(Optional ByVal s As String) As StringBuilder
    Set AppendLine = AppendString(s & vbCrLf)
End Function

''
' Copies a set of characters from the StringBuilder into a character array.
'
' @param sourceindex The starting index in the StringBuilder to be copying from (0-based).
' @param Destination The character array to copy to.
' @param destinationindex The start index to begin placing characters in to.
' @param Count The number of characters to copy.
'
Public Sub CopyTo(ByVal sourceindex As Long, ByRef Destination() As Integer, ByVal destinationindex As Long, ByVal Count As Long)
    If sourceindex < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "SourceIndex", sourceindex)
    If cArray.IsNull(Destination) Then _
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Array), "Destination")
    If destinationindex < LBound(Destination) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "DestinationIndex", destinationindex)
    If Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "Count", Count)
    If sourceindex + Count > mLength Then _
        Throw Cor.NewArgumentException("SourceIndex + Count is greater than the length of the internal string value.", "Count")
    If destinationindex + Count - 1 > UBound(Destination) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset))
    
    If Count > 0 Then
        CopyMemory Destination(destinationindex), ByVal mStringPtr + sourceindex * 2, Count * 2
    End If
End Sub

''
' Returns the buffer size of the string value.
'
' @return Value indicating the amount of memory allocated in characters.
'
Public Property Get Capacity() As Long
    Capacity = mCapacity
End Property

''
' Sets the amount of memory allocated in characters.
'
' @param RHS The number of characters to allocate.
'
Public Property Let Capacity(ByVal RHS As Long)
    Dim ptr As Long
    
    If RHS < mLength Then _
        Throw Cor.NewArgumentOutOfRangeException("Cannot set capacity less than the length of the current instance.", "Capacity", RHS)
    
    ptr = CoTaskMemAlloc(RHS * 2)
    If ptr = 0 Then Throw New OutOfMemoryException
    CopyMemory ByVal ptr, ByVal mStringPtr, mLength * 2
    CoTaskMemFree mStringPtr
    mStringPtr = ptr
    mStringSA.pvData = ptr
    mCapacity = RHS
End Property

''
' Returns a character from the string value.
'
' @param Index The index from the string value to return the character.
' @return The character in the string at the specified index.
' @remarks Index is zero-base.
'
Public Property Get Chars(ByVal Index As Long) As Integer
Attribute Chars.VB_UserMemId = 0
    If Index < 0 Or Index >= mLength Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Index), "index", Index)
    
    Chars = mString(Index)
End Property

''
' Sets a character in the string value.
'
' @param index The index into the string value to set the character.
' @param char The character to set.
' @remarks index is zero-based.
'
Public Property Let Chars(ByVal Index As Long, ByVal Char As Integer)
    If Index < 0 Or Index >= mLength Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Index), "index", Index)
    
    mString(Index) = Char
End Property

''
' Ensures that the internal buffer has atleast the requested amount.
'
' @param requiredCapacity The minimum amount of characters to be allocated.
' @return The capacity after the call is complete.
'
Public Function EnsureCapacity(ByVal RequiredCapacity As Long) As Long
    Dim NewCapacity As Long
    
    If RequiredCapacity <= mCapacity Then
        EnsureCapacity = mCapacity
        Exit Function
    End If
    
    NewCapacity = mCapacity * 2
    If RequiredCapacity > NewCapacity Then NewCapacity = RequiredCapacity
    Capacity = NewCapacity
    EnsureCapacity = NewCapacity
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to test equality on.
' @return Boolean indicating equality.
'
Public Function Equals(ByRef Value As Variant) As Boolean
    If IsObject(Value) Then
        Equals = (Value Is Me)
    End If
End Function

''
' Returns the fill character used to pad empty space.
'
' @return An Integer representing the AscW of the character.
'
Public Property Get FillCharacter() As Variant
    FillCharacter = mFillChar
End Property

''
' Sets the fill character used to pad empty space.
'
' @param RHS The character as either a String or Integer value.
' @remarks The new value can be either String, in which case the
' first characters is used, or it can be a number.
'
Public Property Let FillCharacter(ByVal RHS As Variant)
    Select Case VarType(RHS)
        Case vbLong, vbInteger, vbByte
            mFillChar = AscW(ChrW$(RHS))
        Case vbString
            mFillChar = AscW(RHS)
        Case Else
            Throw Cor.NewInvalidCastException("Fill Character can be either a string or number.")
    End Select
End Property

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
'
Public Function GetHashCode() As Long
    GetHashCode = ObjPtr(Me)
End Function

''
' Inserts a string into the current string value.
'
' @param index The index in the string value to start inserting the new string.
' @param value The string to be inserted.
' @param count The number of times to insert the string.
' @return This instance of StringBuilder
'
Public Function Insert(ByVal Index As Long, ByVal Value As String, Optional ByVal Count As Long = 1) As StringBuilder
    Dim l As Long
    
    If Index < 0 Or Index > mLength Then _
        Throw Cor.NewIndexOutOfRangeException("index must be non-negative and less than or equal to the length.")
    If Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Count must be non-negative.", "count", Count)
        
    l = Len(Value)
    If l > 0 And Count > 0 Then
        mLength = InsertSpace(Index, l * Count)
        l = l * 2
        Index = Index * 2
        Do While Count > 0
            CopyMemory ByVal mStringPtr + Index, ByVal StrPtr(Value), l
            Index = Index + l
            Count = Count - 1
        Loop
    End If
    Set Insert = Me
End Function

''
' Inserts an array of characters into the string value.
'
' @param index The index in the string value to start inserting.
' @param Chars The characters to be inserted.
' @param startindex The index in the Characters to start inserting from.
' @param count The number of characters to be inserted.
' @return This instance of StringBuilder.
'
Public Function InsertChars(ByVal Index As Long, ByRef Chars() As Integer, Optional ByVal StartIndex As Long, Optional ByRef Count As Variant) As StringBuilder
    Dim lb As Long
    Dim cnt As Long
    
    If cArray.IsNull(Chars) Then _
        Throw Cor.NewArgumentNullException("Array cannot be null", "chars")
    
    lb = LBound(Chars)
    If IsMissing(Count) Then
        cnt = cArray.GetLength(Chars) - (StartIndex - lb)
    Else
        cnt = Count
    End If
    
    If StartIndex < lb Then _
        Throw Cor.NewArgumentOutOfRangeException("startindex cannot be less than LowerBound.", "startindex", StartIndex)
    If StartIndex + cnt - 1 > UBound(Chars) Then _
        Throw Cor.NewArgumentException("Invalid count offset.", "Count")
    
    mLength = InsertSpace(Index, cnt)
    CopyMemory mString(Index), Chars(StartIndex), cnt * 2
    Set InsertChars = Me
End Function

''
' Returns the length of the current string value.
'
' @return The length of the string value.
'
Public Property Get Length() As Long
    Length = mLength
End Property

''
' Sets the length of the current string value.
'
' @param RHS The length of the string value.
' @remarks If the new value is longer than the current length, then the
' new space is padded with FillCharacter (default 32).
'
Public Property Let Length(ByVal RHS As Long)
    If RHS < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Length cannot be set less than 0.", "Length", Length)
    
    If RHS > mLength Then
        If RHS > mCapacity Then EnsureCapacity RHS
        Fill mLength, mFillChar, mCapacity - mLength
    End If
    mLength = RHS
End Property

''
' Removes a number of characters from the string value.
'
' @param startindex The index in the string value to start removing from.
' @param count The number of characters to remove from the string value.
' @return This instance of StringBuilder.
' @remarks startindex is zero-based.
'
Public Function Remove(ByVal StartIndex As Long, ByVal Count As Long) As StringBuilder
    If StartIndex < 0 Or Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Argument must be non-negative.", IIf(StartIndex < 0, "startindex", "count"))
    If StartIndex + Count > mLength Then _
        Throw Cor.NewArgumentException("Invalid count offset.", "count")
    
    CopyMemory ByVal mStringPtr + StartIndex * 2, ByVal mStringPtr + (StartIndex + Count) * 2, (mLength - (StartIndex + Count)) * 2
    mLength = mLength - Count
    Set Remove = Me
End Function

''
' Replaces a substring in the string value with a new string.
'
' @param oldValue The substring to be replaced.
' @param newValue The string to replace the old string with.
' @param index The index of the start of the substring in the new string value.
' @param count The number of characters in the new string to include in the replacement.
' @return This instance of StringBuilder.
' @remarks Use index and count to replace the old value with a substring of the new value. index is zero-based.
'
Public Function Replace(ByRef oldValue As String, ByVal newValue As String, Optional ByVal Index As Long, Optional ByRef Count As Variant) As StringBuilder
    Dim occurrences() As Long
    Dim i As Long
    Dim j As Long
    Dim occ As Long
    Dim foundIndex As Long
    Dim oldLen As Long
    Dim newLen As Long
    Dim NewPtr As Long
    
    oldLen = Len(oldValue)
    If oldLen = 0 Then Exit Function
    
    OldStrSA.pvData = StrPtr(oldValue)
    
    ReDim occurrences(mLength - 1)
    foundIndex = -1
    For i = 0 To mLength - oldLen + 1
        If mString(i) = OldStr(j) Then
            j = j + 1
            If foundIndex = -1 Then foundIndex = i
            If j = oldLen Then
                occurrences(occ) = foundIndex
                occ = occ + 1
                j = 0
                foundIndex = -1
            End If
        Else
            foundIndex = -1
            j = 0
        End If
    Next i
    
    If occ > 0 Then
        If Not IsMissing(Count) Then newValue = Mid$(newValue, Index + 1, Count)
        newLen = Len(newValue)
        NewPtr = StrPtr(newValue)
        mCapacity = mLength + newLen * occ
        BufStrSA.pvData = CoTaskMemAlloc(mCapacity * 2)
        j = 0: i = 0
        Dim k As Long
        Dim c As Long
        
        Do While c < occ
            k = occurrences(c) - i
            If k > 0 Then
                CopyMemory ByVal BufStrSA.pvData + j * 2, ByVal mStringSA.pvData + i * 2, k * 2
                j = j + k
                i = i + k
            End If
            If newLen > 0 Then
                CopyMemory ByVal BufStrSA.pvData + j * 2, ByVal NewPtr, newLen * 2
                j = j + newLen
            End If
            i = i + oldLen
            c = c + 1
        Loop
        If i < mLength Then
            CopyMemory ByVal BufStrSA.pvData + j * 2, ByVal mStringSA.pvData + i * 2, (mLength - i) * 2
        End If
        
        CoTaskMemFree mStringPtr
        mStringPtr = BufStrSA.pvData
        mStringSA.pvData = BufStrSA.pvData
        mLength = mLength + (newLen - oldLen) * occ
    End If

    Set Replace = Me
End Function

''
' Returns the current version of the string value.
'
' @param startindex The index to start the return substring from.
' @param length The number of characters to return in the string.
' @return A string or substring representing the internal string of the builder.
' @remarks startindex is zero-based.
'
Public Function ToString(Optional ByVal StartIndex As Long, Optional ByRef Length As Variant) As String
    Dim ln As Long
    
    If mLength = 0 Then
        ToString = ""
        Exit Function
    End If
    
    If StartIndex < 0 Or StartIndex >= mLength Then _
        Throw New IndexOutOfRangeException
    
    If IsMissing(Length) Then
        ln = mLength
    Else
        ln = Length
    End If
    
    If StartIndex + ln > mLength Then _
        Throw Cor.NewArgumentException("Invalid length offset.", "Length")

    ToString = SysAllocStringLen(mStringPtr + StartIndex * 2, ln)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef s As String, ByRef StartIndex As Variant, ByRef Count As Variant, ByVal Capacity As Long)
    Dim cnt As Long
    Dim idx As Long
    
    CoTaskMemFree mStringPtr
    
    If IsMissing(Count) Then
        cnt = Len(s)
    Else
        cnt = Count
    End If
    If IsMissing(StartIndex) Then
        idx = 0
    Else
        idx = StartIndex
    End If
    
    If idx < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "StartIndex", idx)
    If cnt < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "Count", cnt)
    If idx + cnt > Len(s) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "Count")
    
    If cnt > Capacity Then Capacity = cnt
    If Capacity < DEF_CAPACITY Then Capacity = DEF_CAPACITY
    
    mStringPtr = CoTaskMemAlloc(Capacity * 2)
    mCapacity = Capacity
    mLength = cnt
    CopyMemory ByVal mStringPtr, ByVal StrPtr(s) + idx * 2, cnt * 2
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mFillChar = DEF_FILLCHAR
    mStringPtr = CoTaskMemAlloc(DEF_CAPACITY * 2)
    mCapacity = DEF_CAPACITY
    If mStringPtr = 0 Then Throw New OutOfMemoryException
    
    With OldStrSA
        .cbElements = 2
        .cDims = 1
        .cElements = &H7FFFFFFF
    End With
    SAPtr(OldStr) = VarPtr(OldStrSA)
    BufStrSA = OldStrSA
    SAPtr(BufStr) = VarPtr(BufStrSA)
    mStringSA = OldStrSA
    SAPtr(mString) = VarPtr(mStringSA)
    mStringSA.pvData = mStringPtr
End Sub

Private Sub Class_Terminate()
    CoTaskMemFree mStringPtr
    SAPtr(OldStr) = 0
    SAPtr(BufStr) = 0
    SAPtr(mString) = 0
End Sub


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
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Fill(ByVal Index As Long, ByVal Char As Integer, ByVal Count As Long)
    Dim i As Long
    
    For i = Index To Index + Count - 1
        mString(i) = Char
    Next i
End Sub

Private Function FindCloseBrace(ByRef ch As Integer, ByRef pos As Long, ByVal Max As Long) As Long
    Dim i As Long
    i = pos
    Do
        If pos = Max Then FormatError
        ch = BufStr(pos)
        pos = pos + 1
        If ch = OPEN_BRACE Then
            If pos >= Max Then FormatError
            If BufStr(pos) = OPEN_BRACE Then pos = pos + 1
        ElseIf ch = CLOSE_BRACE Then
            If BufStr(pos) = CLOSE_BRACE Then
                pos = pos + 1
            Else
                pos = pos - 1
                Exit Do
            End If
        End If
        BufStr(i) = ch
        i = i + 1
    Loop
    FindCloseBrace = i
End Function

Private Function FindIndex(ByRef pos As Long, ByVal Max As Long) As Long
    Dim ch As Integer
    ch = BufStr(pos)
    If ch < 48 Or ch > 57 Then FormatError
    FindIndex = 0
    Do
        FindIndex = FindIndex * 10 + ch - 48
        pos = pos + 1
        If pos = Max Then FormatError
        ch = BufStr(pos)
    Loop While ch >= 48 And ch <= 57 And FindIndex < 1000000
End Function

Private Function FindOpenBrace(ByRef pos As Long, ByVal Max As Long) As Long
    Dim ch As Integer
    Dim i As Long
    i = pos
    Do While pos < Max
        ch = BufStr(pos)
        pos = pos + 1
        Select Case ch
        Case CLOSE_BRACE
            If pos >= Max Then FormatError
            If BufStr(pos) = CLOSE_BRACE Then pos = pos + 1
        Case OPEN_BRACE
            If pos >= Max Then
                pos = pos - 1
                Exit Do
            ElseIf BufStr(pos) = OPEN_BRACE Then
                pos = pos + 1
            Else
                pos = pos - 1
                Exit Do
            End If
        End Select
        BufStr(i) = ch
        i = i + 1
    Loop
    FindOpenBrace = i
End Function

Private Function FindWidth(ByVal ch As Integer, ByRef pos As Long, ByVal Max As Long) As Long
    Dim sign As Long
    Dim Width As Long
    sign = 1
    If ch = COMMA Then
        pos = pos + 1
        Do While pos < Max
            If BufStr(pos) <> 32 Then Exit Do
            pos = pos + 1
        Loop
        If pos = Max Then FormatError
        ch = BufStr(pos)
        If ch = MINUS Then
            sign = -1
            pos = pos + 1
            If pos = Max Then FormatError
            ch = BufStr(pos)
        End If
        If ch < 48 Or ch > 57 Then FormatError
        Do
            Width = Width * 10 + ch - 48
            pos = pos + 1
            If pos = Max Then FormatError
            ch = BufStr(pos)
        Loop While ch >= 48 And ch <= 57 And Width < 1000000
    End If
    FindWidth = Width * sign
End Function

Private Sub FormatError()
    Throw Cor.NewArgumentException("Invalid format.", "format")
End Sub

Private Function Formattable(ByRef arg As Variant) As IFormattable
    Set Formattable = arg
End Function

Private Function GetFormatInfo(ByRef ch As Integer, ByRef pos As Long, ByRef Format As String, ByVal Max As Long) As String
    Dim p As Long
    Dim i As Long
    If ch = COLON Then
        pos = pos + 1
        p = pos
        i = FindCloseBrace(ch, pos, Max)
        If i > p Then GetFormatInfo = Mid$(Format, p + 1, i - p)
    End If
End Function

Private Function GetNextNonWhiteSpace(ByRef pos As Long, ByVal Max As Long) As Integer
    Dim ch As Integer
    Do While pos < Max
        ch = BufStr(pos)
        If ch <> 32 Then Exit Do
        pos = pos + 1
    Loop
    GetNextNonWhiteSpace = ch
End Function

Private Function InsertSpace(ByVal Index As Long, ByVal Size As Long) As Long
    InsertSpace = mLength + Size
    EnsureCapacity InsertSpace
    If Index < mLength Then
        ' create space in the string.
        CopyMemory ByVal mStringPtr + (Index + Size) * 2, ByVal mStringPtr + Index * 2, (mLength - Index) * 2
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend helper methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub InternalAppendFormat(ByVal provider As IFormatProvider, ByVal Format As String, ByRef args() As Variant)
    
    Dim ch As Integer
    Dim pos As Long
    Dim Max As Long
    Dim i As Long
    Dim p As Long
    Dim Index As Long
    Dim argsLength As Long
    Dim Width As Long
    Dim leftJustify As Boolean
    Dim fmt As String
    Dim s As String
    Dim cf As ICustomFormatter
    Dim pad As Long
    
    argsLength = UBound(args) + 1
    Max = Len(Format)
    BufStrSA.pvData = StrPtr(Format)
    
    If Not provider Is Nothing Then Set cf = provider.GetFormat("ICustomFormatter")
    
    Do
        p = pos
        i = FindOpenBrace(pos, Max)
        If i > p Then AppendChars BufStr, p, i - p
        If pos = Max Then Exit Do
        pos = pos + 1
        If pos = Max Then FormatError
        
        Index = FindIndex(pos, Max)
        If Index >= argsLength Then _
            Throw Cor.NewFormatException("The argument index cannot be greater than the number of arguments.")
        
        ch = GetNextNonWhiteSpace(pos, Max)
        Width = FindWidth(ch, pos, Max)
        If Width < 0 Then
            leftJustify = True
            Width = -Width
        Else
            leftJustify = False
        End If
        
        ch = GetNextNonWhiteSpace(pos, Max)
        fmt = GetFormatInfo(ch, pos, Format, Max)
        If ch <> CLOSE_BRACE Then FormatError
        pos = pos + 1
        
        s = vbNullString
        If Not cf Is Nothing Then
            s = cf.Format(fmt, args(Index), provider)
        End If
        
        If StrPtr(s) = 0 Then s = Convert.ToString(args(Index), fmt, provider)
        
        pad = Width - Len(s)
        If Not leftJustify And (pad > 0) Then AppendChar 32, pad
        AppendString s
        If leftJustify And (pad > 0) Then AppendChar 32, pad
    Loop
End Sub


