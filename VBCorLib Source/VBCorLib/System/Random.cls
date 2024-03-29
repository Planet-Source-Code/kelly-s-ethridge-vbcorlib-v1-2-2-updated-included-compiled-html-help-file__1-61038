VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Random"
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
'    Module: Random
'

''
' Provides a sequence of pseudo-random generated numbers.
'
' @see Constructors
'
Option Explicit
Implements cObject

''
' Returns the next generated number in the sequence.
'
' @param maxValue The value that the generated number must be less than.
' @return The next generated number in the sequence.
' @remarks The generated number will be equal or greater than zero and less than
' maxValue. The number will never equal maxValue.
Public Function NextLong(Optional ByVal MaxValue As Long = &H7FFFFFFF) As Long
    If MaxValue < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "MaxValue", MaxValue)
    NextLong = NextNumber(0, MaxValue)
End Function

''
' Returns the next generated number in the sequence within a specified range.
'
' @param minValue The lowerbound of the number to be generated.
' @param maxValue The upperbound of the number to be generated. Must be equal to or greater than minValue.
' @return A number generated within the specified range.
' @remarks The number generated will always be equal to or greater than minValue, and less than maxValue.
' The number will never equal maxValue.
Public Function NextRange(ByVal MinValue As Long, ByVal MaxValue As Long) As Long
    If MaxValue < MinValue Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_MinMax), "maxValue", MaxValue)
    NextRange = NextNumber(MinValue, MaxValue)
End Function

''
' Fills an array of bytes with random numbers.
'
' @param bytes The array to be filled.
Public Sub NextBytes(ByRef Bytes() As Byte)
    Dim i As Long
    
    If cArray.IsNull(Bytes) Then _
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Array), "bytes")
        
    For i = LBound(Bytes) To UBound(Bytes)
        Bytes(i) = NextNumber(0, 256)    ' it won't reach 256, 255 is the max.
    Next i
End Sub

''
' Returns a random number between 0.0 and 1.0.
'
' @return The random number.
' @remarks The number will be equal to or greater than 0.0 and less than 1.0.
' The number will never equal 1.0.
Public Function NextDouble() As Double
    NextDouble = Rnd
End Function

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Public Function ToString() As String
    ToString = App.Title & "." & TypeName(Me)
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equality to.
' @return Boolean indicating equality.
Public Function Equals(ByRef Value As Variant) As Boolean
    If IsObject(Value) Then
        Equals = (Value Is Me)
    End If
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long
    GetHashCode = ObjPtr(Me)
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal seed As Long)
    If seed > 0 Then seed = -seed
    Rnd seed
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function NextNumber(ByVal MinValue As Long, ByVal MaxValue As Long) As Long
    NextNumber = Int(Rnd * (MaxValue - MinValue)) + MinValue
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Randomize
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   cObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function cObject_Equals(Value As Variant) As Boolean
    cObject_Equals = Equals(Value)
End Function

Private Function cObject_GetHashCode() As Long
    cObject_GetHashCode = GetHashCode
End Function

Private Function cObject_ToString() As String
    cObject_ToString = ToString
End Function
