Attribute VB_Name = "modcDateTimeHelpers"
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
'    Module: modcDateTimeHelpers
'
Option Explicit

Public DaysToMonthLeapYear() As Long
Public DaysToMonth() As Long




Public Sub InitcDateTimeHelpers()
    DaysToMonth = cArray.NewArray(ciLong, 0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334, 365)
    DaysToMonthLeapYear = cArray.NewArray(ciLong, 0, 31, 60, 91, 121, 152, 182, 213, 244, 274, 305, 335, 366)
End Sub


