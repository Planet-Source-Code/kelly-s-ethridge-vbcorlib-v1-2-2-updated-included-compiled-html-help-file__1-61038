Attribute VB_Name = "modPublicFunctions"
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
'    Module: modPublicFunctions
'
'   This mirrors PublicFunctions.cls to allow the project access to the
'   same set of public functions.
Option Explicit
Public Declare Function GetCalendarInfo Lib "kernel32.dll" Alias "GetCalendarInfoA" (ByVal Locale As Long, ByVal Calendar As Long, ByVal CalType As Long, ByVal lpCalData As String, ByVal cchData As Long, ByRef lpValue As Any) As Long

Private Const LOCALE_RETURN_NUMBER As Long = &H20000000
Private Const CAL_RETURN_NUMBER As Long = LOCALE_RETURN_NUMBER

Private Const LOCALE_BUFFER_SIZE As Long = 1024

Public Cor As Constructors
Public cArray As cArray
Public cString As cString
Public comparer As ComparerStatic
Public Environment As Environment
Public Buffer As Buffer
Public NumberFormatInfo As NumberFormatInfoStatic
Public BitConverter As BitConverter
Public TimeSpan As TimeSpanStatic
Public cDateTime As cDateTimeStatic
Public DateTimeFormatInfo As DateTimeFormatInfoStatic
Public CultureTable As CultureTable
Public CultureInfo As CultureInfoStatic
Public Path As Path
Public Encoding As EncodingStatic
Public Directory As Directory
Public file As file

' create these only if they are used.
Private mConsole As Console
Private mCalendar As CalendarStatic
Private mGregorianCalendar As GregorianCalendarStatic
Private mJulianCalendar As JulianCalendarStatic
Private mHebrewCalendar As HebrewCalendarStatic
Private mKoreanCalendar As KoreanCalendarStatic
Private mThaiBuddhistCalendar As ThaiBuddhistCalendarStatic
Private mHijriCalendar As HijriCalendarStatic
Private mArrayList As ArrayListStatic
Private mVersion As VersionStatic
Private mBitArray As BitArrayStatic
Private mTimeZone As TimeZoneStatic
Private mStream As StreamStatic
Private mTextReader As TextReaderStatic
Private mRegistry As Registry
Private mRegistryKey As RegistryKeyStatic
Private mGuid As GuidStatic
Private mConvert As Convert
Private mResourceManager As ResourceManagerStatic
Private mDriveInfo As DriveInfoStatic



Public Powers(31) As Long
Private mLocaleBuffer As String



Public Sub InitPublicFunctions()
    mLocaleBuffer = String$(LOCALE_BUFFER_SIZE, 0)
    
    InitPowers
    
    Set comparer = New ComparerStatic
    Set Cor = New Constructors
    Set cArray = New cArray
    Set cString = New cString
    Set Environment = New Environment
    Set Buffer = New Buffer
    Set CultureTable = New CultureTable
    Set CultureInfo = New CultureInfoStatic
    Set NumberFormatInfo = New NumberFormatInfoStatic
    Set BitConverter = New BitConverter
    Set TimeSpan = New TimeSpanStatic
    Set cDateTime = New cDateTimeStatic
    Set DateTimeFormatInfo = New DateTimeFormatInfoStatic
    Set Path = New Path
    Set Encoding = New EncodingStatic
    Set Directory = New Directory
    Set file = New file
    
End Sub

Public Function FuncAddr(ByVal pfn As Long) As Long
    FuncAddr = pfn
End Function

Public Function CObj(ByRef Value As Variant) As Object
    Set CObj = Value
End Function

Public Function Modulus(ByVal x As Currency, ByVal y As Currency) As Currency
  Modulus = x - (y * Fix(x / y))
End Function

Public Function GetFileNameFromFindData(ByRef Data As WIN32_FIND_DATA) As String
    Select Case Data.cFileName(0)
        Case 46:    ' skip periods
        Case 0:     GetFileNameFromFindData = Environment.BytesToString(Data.cAlternate, 0, lstrlen(VarPtr(Data.cAlternate(0))))
        Case Else:  GetFileNameFromFindData = Environment.BytesToString(Data.cFileName, 0, lstrlen(VarPtr(Data.cFileName(0))))
    End Select
End Function

Public Function GetLocaleLong(ByVal LCID As Long, ByVal LCTYPE As Long) As Long
    Dim n As Long
    
    n = GetLocaleInfo(LCID, LCTYPE, mLocaleBuffer, LOCALE_BUFFER_SIZE)
    If n > 0 Then GetLocaleLong = CLng(left$(mLocaleBuffer, n - 1))
End Function

Public Function GetLocaleString(ByVal LCID As Long, ByVal LCTYPE As Long) As String
    Dim n As Long
    
    n = GetLocaleInfo(LCID, LCTYPE, mLocaleBuffer, LOCALE_BUFFER_SIZE)
    GetLocaleString = left$(mLocaleBuffer, n - 1)
End Function

Public Function GetCalendarLong(ByVal Cal As Long, ByVal CalType As Long) As Long
    GetCalendarInfo LOCALE_USER_DEFAULT, Cal, CalType Or CAL_RETURN_NUMBER, vbNullString, 0, GetCalendarLong
End Function



''
' Creates the static classes only if they are needed.
'
Public Function Console() As Console
    If mConsole Is Nothing Then Set mConsole = New Console
    Set Console = mConsole
End Function

Public Function Calendar() As CalendarStatic
    If mCalendar Is Nothing Then Set mCalendar = New CalendarStatic
    Set Calendar = mCalendar
End Function

Public Function GregorianCalendar() As GregorianCalendarStatic
    If mGregorianCalendar Is Nothing Then Set mGregorianCalendar = New GregorianCalendarStatic
    Set GregorianCalendar = mGregorianCalendar
End Function

Public Function JulianCalendar() As JulianCalendarStatic
    If mJulianCalendar Is Nothing Then Set mJulianCalendar = New JulianCalendarStatic
    Set JulianCalendar = mJulianCalendar
End Function

Public Function HebrewCalendar() As HebrewCalendarStatic
    If mHebrewCalendar Is Nothing Then Set mHebrewCalendar = New HebrewCalendarStatic
    Set HebrewCalendar = mHebrewCalendar
End Function

Public Function KoreanCalendar() As KoreanCalendarStatic
    If mKoreanCalendar Is Nothing Then Set mKoreanCalendar = New KoreanCalendarStatic
    Set KoreanCalendar = mKoreanCalendar
End Function

Public Function ThaiBuddhistCalendar() As ThaiBuddhistCalendarStatic
    If mThaiBuddhistCalendar Is Nothing Then Set mThaiBuddhistCalendar = New ThaiBuddhistCalendarStatic
    Set ThaiBuddhistCalendar = mThaiBuddhistCalendar
End Function

Public Function HijriCalendar() As HijriCalendarStatic
    If mHijriCalendar Is Nothing Then Set mHijriCalendar = New HijriCalendarStatic
    Set HijriCalendar = mHijriCalendar
End Function

Public Function ArrayList() As ArrayListStatic
    If mArrayList Is Nothing Then Set mArrayList = New ArrayListStatic
    Set ArrayList = mArrayList
End Function

Public Function Version() As VersionStatic
    If mVersion Is Nothing Then Set mVersion = New VersionStatic
    Set Version = mVersion
End Function

Public Function BitArray() As BitArrayStatic
    If mBitArray Is Nothing Then Set mBitArray = New BitArrayStatic
    Set BitArray = mBitArray
End Function

Public Function TimeZone() As TimeZoneStatic
    If mTimeZone Is Nothing Then Set mTimeZone = New TimeZoneStatic
    Set TimeZone = mTimeZone
End Function

Public Function Stream() As StreamStatic
    If mStream Is Nothing Then Set mStream = New StreamStatic
    Set Stream = mStream
End Function

Public Function TextReader() As TextReaderStatic
    If mTextReader Is Nothing Then Set mTextReader = New TextReaderStatic
    Set TextReader = mTextReader
End Function

Public Function Registry() As Registry
    If mRegistry Is Nothing Then Set mRegistry = New Registry
    Set Registry = mRegistry
End Function

Public Function RegistryKey() As RegistryKeyStatic
    If mRegistryKey Is Nothing Then Set mRegistryKey = New RegistryKeyStatic
    Set RegistryKey = mRegistryKey
End Function

Public Function Guid() As GuidStatic
    If mGuid Is Nothing Then Set mGuid = New GuidStatic
    Set Guid = mGuid
End Function

Public Function Convert() As Convert
    If mConvert Is Nothing Then Set mConvert = New Convert
    Set Convert = mConvert
End Function

Public Function ResourceManager() As ResourceManagerStatic
    If mResourceManager Is Nothing Then Set mResourceManager = New ResourceManagerStatic
    Set ResourceManager = mResourceManager
End Function

Public Function DriveInfo() As DriveInfoStatic
    If mDriveInfo Is Nothing Then Set mDriveInfo = New DriveInfoStatic
    Set DriveInfo = mDriveInfo
End Function



''
' Initializes an array for quick powers of 2 lookup.
'
Private Sub InitPowers()
    Dim i As Long
    For i = 0 To 30
        Powers(i) = 2 ^ i
    Next i
    Powers(31) = &H80000000
End Sub


