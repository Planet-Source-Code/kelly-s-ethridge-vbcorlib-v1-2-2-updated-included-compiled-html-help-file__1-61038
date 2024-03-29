VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDateTimeStatic"
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
'    Module: cDateTimeStatic
'

''
' Represents methods to create cDateTime objects from various data sources.
'
' @remarks These methods are considered static. The <b>cDateTimeStatic</b> cannot
' be directly instantiated. Use these methods withe <b>cDateTime.*</b> syntax.
' <pre>
' Dim dt As cDateTime
' Set dt = cDateTime.FromOADate(#1/1/2001 8:30:00 AM#)
'
' ' A quicker way to create a <b>cDateTime</b> object is to use the <b>NewDate</b> method.
' Set dt = NewDate(#1/1/2001 8:30:00 AM#)
' </pre>
' This will create a new instance of <b>cDateTime</b> with the supplied date value.
' @see cDateTime
'
Option Explicit

' we put this here instead of cDateTime to prevent conflicts with
' the cDateTime.DayOfWeek property.

''
' This enumeration represents the days of the week as used by cDateTime.
'
' @param Sunday 0
' @param Monday 1
' @param Tuesday 2
' @param Wednesday 3
' @param Thursday 4
' @param Friday 5
' @param Saturday 6
' @remarks The days are 0 based, where as VB's vbDayOfWeek is 1 based.
Public Enum DayOfWeek
    Sunday
    Monday
    Tuesday
    Wednesday
    Thursday
    Friday
    Saturday
End Enum

''
' Used to specify the type of cDateTime object.
'
' @param UnspecifiedKind 0
' @param UtcKind 1
' @param LocalKind 2
'
Public Enum DateTimeKind
    UnspecifiedKind = 0
    UtcKind = 1
    LocalKind = 2
End Enum


Private mDaysInMonth() As Long
Private mDateAscDateTime As cDateTime
Private mMinValue As cDateTime
Private mMaxValue As cDateTime
Private mSystemTime As SystemTime


''
' Returns a boolean indicating if the specific year is a leap year.
'
' @param Year The year to test for leap year.
' @return Indication of leap year.
' @remarks If the year is evenly divisible by 4 or 400 and not by 100
' then it is a leap year.
Public Function IsLeapYear(ByVal Year As Long) As Boolean
    If (Year Mod 4) = 0 Then
        IsLeapYear = (Year Mod 100) <> 0 Or (Year Mod 400) = 0
    End If
End Function

''
' Returns the number of days in a gregorian calendar month.
'
' @param Year The year the month is in.
' @param Month The month to find the number of days for.
' @return The number of days in the specified month.
Public Function DaysInMonth(ByVal Year As Long, ByVal Month As Long) As Long
    If Year < 1 Or Year > 9999 Then _
        Throw Cor.NewArgumentOutOfRangeException("Year must be between 1 and 9999.")
    If Month < 1 Or Month > 12 Then _
        Throw Cor.NewArgumentOutOfRangeException("Month must be between 1 and 12.")
    If Month = 2 Then
        If IsLeapYear(Year) Then
            DaysInMonth = 29
            Exit Function
        End If
    End If
    DaysInMonth = mDaysInMonth(Month - 1)
End Function

''
' Returns a cDateTime object from a VB Date value.
'
' @param d The standard VB Date value.
' @param Kind Specifies whether the cDateTime is a UTC, Local, or unspecified time type.
' @return The cDateTime representation of the date.
Public Function FromOADate(ByVal d As Date, Optional ByVal Kind As DateTimeKind = UnspecifiedKind) As cDateTime
    Set FromOADate = New cDateTime
    FromOADate.InitFromDate d, Kind
End Function

''
' Returns a cDateTime object from a FileTime value.
'
' @param time The file time in UTC time.
' @return The cDateTime of the file time converted to local time from UTC time.
Public Function FromFileTime(ByVal Time As Currency) As cDateTime
    Set FromFileTime = FromFileTimeUtc(Time).ToLocalTime
End Function

''
' Returns a cDateTime object from a FileTime value.
'
' @param time The file time in UTC time.
' @return The cDateTime of the file time in UTC time.
Public Function FromFileTimeUtc(ByVal Time As Currency) As cDateTime
    Set FromFileTimeUtc = New cDateTime
    FromFileTimeUtc.InitFromFileTime Time, UtcKind
End Function

''
' Returns a cDateTime object from milliseconds.
'
' @param ms The number of milliseconds the time is equal to.
' @param Kind Specifies whether the cDateTime is a UTC, Local, or unspecified time type.
' @return The cDateTime object converted from milliseconds.
' @remarks Since a Currency can be passed in, the fractional portion
' can be used to represent the number ticks (1/10000 ms) in the date.
Public Function FromMilliseconds(ByVal ms As Currency, Optional ByVal Kind As DateTimeKind = UnspecifiedKind) As cDateTime
    Set FromMilliseconds = New cDateTime
    FromMilliseconds.InitFromMilliseconds ms, Kind
End Function

''
' Returns a cDateTime that represents the current time on the local machine.
'
' @return The cDateTime object representing the current time on the local machine.
' @remarks This method uses GetLocalTime which has a resolution down to the
' milliseconds. If the milliseconds is not wanted, then either the Millisecond
' must be subtracted for the instance, or use cDateTime.FromOADate(Now).
Public Function Now() As cDateTime
    GetLocalTime mSystemTime
    With mSystemTime
        Set Now = Cor.NewcDateTime(.wYear, .wMonth, .wDay, .wHour, .wMinute, .wSecond, .wMilliseconds)
    End With
End Function

''
' Returns a cDateTime that represents the current time with the UTC offset.
'
' @return The cDateTime object representing the current time with the UTC offset.
' @remarks This method uses GetSystemTime which has a resolution down to the
' milliseconds. There is no direct VB equivalent.
Public Function NowUtc() As cDateTime
    GetSystemTime mSystemTime
    With mSystemTime
        Set NowUtc = Cor.NewcDateTime(.wYear, .wMonth, .wDay, .wHour, .wMinute, .wSecond, .wMilliseconds)
    End With
End Function

''
' Returns a cDateTime that represents only the current date. The time
' portion is not present.
'
' @return The cDateTime representing the date.
' @remarks The time is still there, just is only 12:00AM.
Public Function Today() As cDateTime
    Set Today = FromOADate(Date)
End Function

''
' Returns the minimum cDateTime representation.
'
' @return The minimum cDateTime.
' @remarks The minimum cDateTime is 1/1/0001 12:00:00 AM.
Public Property Get MinValue() As cDateTime
    If mMinValue Is Nothing Then Set mMinValue = Cor.NewcDateTime(1, 1, 1)
    Set MinValue = mMinValue
End Property

''
' Returns the maximum cDateTime representation.
'
' @return The maximum cDateTime.
' @remarks The maximum cDateTime is 12/31/9999 11:59:59 PM."
Public Property Get MaxValue() As cDateTime
    If mMaxValue Is Nothing Then Set mMaxValue = Cor.NewcDateTime(9999, 12, 31, 23, 59, 59, 999)
    Set MaxValue = mMaxValue
End Property

''
' Parses a string representation of a date into a cDateTime object.
'
' @param s The date string to be parsed.
' @return A cDateTime object created from the date string.
' @remarks The parse method simply uses CDate to parse the string.
Public Function Parse(ByRef s As String) As cDateTime
    Set Parse = FromOADate(CDate(s))
End Function

''
' Creates a cDateTime object from a 64-bit currency value.
'
' @param Value The value to initialize the cDateTime with.
' @return A new cDateTime object.
' @remarks The 64-bit value contains both the time in 100-nanosecond
' increments and the DateTimeKind value used to determine if the
' value is a UTC, Local Time, or unspecified time.
' <p>The upper 2-bits represent the DateTimeKind value.</p>
'
Public Function FromBinary(ByVal Value As Currency) As cDateTime
    Set FromBinary = New cDateTime
    FromBinary.InitFromBinary Value
End Function

''
' Creates a new cDateTime object with the specified DateTimeKind.
'
' @param Value The cDateTime to convert to a specific kind.
' @param Kind Specifies whether the cDateTime is a UTC, Local, or unspecified time type.
' @return A new cDateTime with the original date and new kind.
'
Public Function SpecifyKind(ByVal Value As cDateTime, ByVal Kind As DateTimeKind) As cDateTime
    Set SpecifyKind = New cDateTime
    SpecifyKind.InitFromMilliseconds Value.TotalMilliseconds, Kind
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function AscDateTime(ByVal d As Date) As cDateTime
    Set AscDateTime = mDateAscDateTime
    AscDateTime.InitFromDate d, UnspecifiedKind
End Function

Friend Function GetcDateTime(ByRef d As Variant) As cDateTime
    Select Case VarType(d)
        Case vbObject
            If TypeOf d Is cDateTime Then
                Set GetcDateTime = d
            Else
                Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_DateRequired))
            End If
        Case vbDate
            Set GetcDateTime = FromOADate(d)
        Case Else
            Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_DateRequired))
    End Select
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mDaysInMonth = cArray.NewArray(ciLong, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
    InitcDateTimeHelpers
    Set mDateAscDateTime = New cDateTime
End Sub
