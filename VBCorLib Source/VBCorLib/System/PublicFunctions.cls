VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PublicFunctions"
Attribute VB_GlobalNameSpace = True
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
'    Module: PublicFunctions
'

''
' Provides direct access to static class methods.
'
' @remarks VB6 does not support static class methods. To replicate the syntax of
' accessing static methods, a function of the same name as the class that is
' supposed to contain the static methods is defined in this class. This class
' has instancing of <i>GlobalMultiUse</i>. This means that just the name of these
' functions can be typed directly without having to instantiate the <b>PublicFunctions</b>
' class. By just typing in any of these function names, the underlying class which
' contains the static methods is returned, and those methods are then available
' for calling. This gives the illusion of static methods being called on a class.
' <p>When typing in one of these function names, it should be followed immediately
' by a period to allow access to the static methods of that class.</p>
' <pre>
' '' Print the name of the current user
' Debug.Print Environment.UserName
' </pre>
' This small example simply refers to the <i>Environment</i> function in this
' class, then directly calls a method on the returned object from the
' <i>Environment</i> function.
'
Option Explicit

''
' Returns a object containing all of the parameterized constructors
' available within the VBCorLib library.
'
' @return A object containing all of the constructors in VBCorLib.
' @see constructors
'
Public Function Cor() As Constructors
    Set Cor = modPublicFunctions.Cor
End Function

''
' Returns an object used to access the <b>cArray</b> methods
' without instantiating the <b>cArray</b> class.
'
' @return A reference to the <b>cArray</b> object used to
' perform functions on arrays.
' @see cArray
'
Public Function cArray() As cArray
    Set cArray = modPublicFunctions.cArray
End Function

''
' Returns an object used to access the <b>cString</b> methods
' without instantiating the <b>cString</b> class.
'
' @return A reference to the <b>cString</b> object used to
' perform functions on Strings.
' @see cString
'
Public Function cString() As cString
    Set cString = modPublicFunctions.cString
End Function

''
' Returns an object used to access the <b>Comparer</b> static functions.
'
' @return A reference to the object used to access the <b>Comparer</b> static functions.
' @see Comparer
' @see ComparerStatic
'
Public Function comparer() As ComparerStatic
    Set comparer = modPublicFunctions.comparer
End Function

''
' Returns a reference used to access the methods of the <b>Environment</b> functions.
'
' @return A reference that can be used to access methods of the <b>Environment</b> class.
' @see Environment
'
Public Function Environment() As Environment
    Set Environment = modPublicFunctions.Environment
End Function

''
' Returns a reference used to access static methods of the <b>BitArray</b> class.
'
' @return A reference that can be used to access static methods of the <b>BitArray</b> class.
' @see BitArray
' @see BitArrayStatic
Public Function BitArray() As BitArrayStatic
    Set BitArray = modPublicFunctions.BitArray
End Function

''
' Returns a reference used to access static methods of the <b>Buffer</b> class.
'
' @return A reference that can be used to access static methods of the <b>Buffer</b> class.
' @see Buffer
'
Public Function Buffer() As Buffer
    Set Buffer = modPublicFunctions.Buffer
End Function

''
' Returns a reference used to access static methods of the <b>NumberFormatInfo</b> class.
'
' @return A reference that can be used to access static methods of the <b>NumberFormatInfo</b> class.
' @see NumberFormatInfo
' @see NumberFormatInfoStatic
'
Public Function NumberFormatInfo() As NumberFormatInfoStatic
    Set NumberFormatInfo = modPublicFunctions.NumberFormatInfo
End Function

''
' Returns a reference used to access static methods of the <b>BitConverter</b> class.
'
' @return A reference that can be used to access static methods of the <b>BitConverter</b> class.
' @see BitConverter
'
Public Function BitConverter() As BitConverter
    Set BitConverter = modPublicFunctions.BitConverter
End Function

''
' Returns a reference used to access static methods of the <b>Version</b> class.
'
' @return A reference that can be used to access static methods of the <b>Version</b> class.
' @see Version
' @see VersionStatic
'
Public Function Version() As VersionStatic
    Set Version = modPublicFunctions.Version
End Function

''
' Returns a reference used to access static methods of the <b>TimeSpan</b> class.
'
' @return A reference that can be used to access static methods of the <b>TimeSpan</b> class.
' @see TimeSpan
' @see TimeSpanStatic
'
Public Function TimeSpan() As TimeSpanStatic
    Set TimeSpan = modPublicFunctions.TimeSpan
End Function

''
' Returns a reference used to access static methods of the <b>cDateTime</b> class.
'
' @return A reference that can be used to access static methods of the <b>cDateTime</b> class.
' @see cDateTime
' @see cDateTimeStatic
'
Public Function cDateTime() As cDateTimeStatic
    Set cDateTime = modPublicFunctions.cDateTime
End Function

''
' Returns a reference used to access static methods of the <b>DateTimeFormatInfo</b> class.
'
' @return A reference that can be used to access static methods of the <b>DateTimeFormatInfo</b> class.
' @see DateTimeFormatInfo
' @see DateTimeFormatInfoStatic
'
Public Function DateTimeFormatInfo() As DateTimeFormatInfoStatic
    Set DateTimeFormatInfo = modPublicFunctions.DateTimeFormatInfo
End Function

''
' Returns a reference used to access static methods of the <b>CultureInfo</b> class.
'
' @return A reference that can be used to access static methods of the <b>CultureInfo</b> class.
' @see CultureInfo
' @see CultureInfoStatic
'
Public Function CultureInfo() As CultureInfoStatic
    Set CultureInfo = modPublicFunctions.CultureInfo
End Function

''
' Returns a reference used to access static methods of the <b>TimeZone</b> class.
'
' @return A reference that can be used to access static methods of the <b>TimeZone</b> class.
' @see TimeZone
' @see TimeZoneStatic
'
Public Function TimeZone() As TimeZoneStatic
    Set TimeZone = modPublicFunctions.TimeZone
End Function

''
' Returns a reference used to access static methods of the <b>Path</b> class.
'
' @return A reference that can be used to access static methods of the <b>Path</b> class.
' @see Path
'
Public Function Path() As Path
    Set Path = modPublicFunctions.Path
End Function

''
' Returns a reference used to access static methods of the <b>Encoding</b> class.
'
' @return A reference that can be used to access static methods of the <b>Encoding</b> class.
' @see Encoding
' @see EncodingStatic
'
Public Function Encoding() As EncodingStatic
    Set Encoding = modPublicFunctions.Encoding
End Function

''
' Returns a reference used to access static methods of the <b>TextReader</b> class.
'
' @return A reference that can be used to access static methods of the <b>TextReader</b> class.
' @see TextReader
' @see TextReaderStatic
' @see StreamReader
'
Public Function TextReader() As TextReaderStatic
    Set TextReader = modPublicFunctions.TextReader
End Function

''
' Returns a reference used to access static methods of the <b>Directory</b> class.
'
' @return A reference that can be used to access static methods of the <b>Directory</b> class.
' @see Directory
' @see DirectoryInfo
' @see File
' @see FileInfo
' @see DriveInfo
' @see DriveInfoStatic
'
Public Function Directory() As Directory
    Set Directory = modPublicFunctions.Directory
End Function

''
' Returns a reference used to access static methods of the <b>File</b> class.
'
' @return A reference that can be used to access static methods of the <b>File</b> class.
' @see File
' @see FileInfo
' @see Directory
' @see DirectoryInfo
' @see DriveInfo
' @see DriveInfoStatic
'
Public Function file() As file
    Set file = modPublicFunctions.file
End Function

''
' Returns a reference used to access static methods of the <b>Stream</b> class.
'
' @return A reference that can be used to access static methods of the <b>Stream</b> class.
' @see Stream
' @see StreamStatic
'
Public Function Stream() As StreamStatic
    Set Stream = modPublicFunctions.Stream
End Function

''
' Returns a reference used to access static methods of the <b>GregorianCalendar</b> class.
'
' @return A reference that can be used to access static methods of the <b>GregorianCalendar</b> class.
' @see GregorianCalendar
' @see GregorianCalendarStatic
'
Public Function GregorianCalendar() As GregorianCalendarStatic
    Set GregorianCalendar = modPublicFunctions.GregorianCalendar
End Function

''
' Returns a reference used to access static methods of the <b>JulianCalendar</b> class.
'
' @return A reference that can be used to access static methods of the <b>JulianCalendar</b> class.
' @see JulianCalendar
' @see JulianCalendarStatic
'
Public Function JulianCalendar() As JulianCalendarStatic
    Set JulianCalendar = modPublicFunctions.JulianCalendar
End Function

''
' Returns a reference used to access static methods of the <b>HebrewCalendar</b> class.
'
' @return A reference that can be used to access static methods of the <b>HebrewCalendar</b> class.
' @see HebrewCalendar
' @see HebrewCalendarStatic
'
Public Function HebrewCalendar() As HebrewCalendarStatic
    Set HebrewCalendar = modPublicFunctions.HebrewCalendar
End Function

''
' Returns a reference used to access static methods of the <b>KoreanCalendar</b> class.
'
' @return A reference that can be used to access static methods of the <b>KoreanCalendar</b> class.
' @see KoreanCalendar
' @see KoreanCalendarStatic
'
Public Function KoreanCalendar() As KoreanCalendarStatic
    Set KoreanCalendar = modPublicFunctions.KoreanCalendar
End Function

''
' Returns a reference used to access static methods of the <b>HijriCalendar</b> class.
'
' @return A reference that can be used to access static methods of the <b>HijriCalendar</b> class.
' @see HijriCalendar
' @see HijriCalendarStatic
'
Public Function HijriCalendar() As HijriCalendarStatic
    Set HijriCalendar = modPublicFunctions.HijriCalendar
End Function

''
' Returns a reference used to access static methods of the <b>Calendar</b> class.
'
' @return A reference that can be used to access static methods of the <b>Calendar</b> class.
' @see Calendar
' @see CalendarStatic
'
Public Function Calendar() As CalendarStatic
    Set Calendar = modPublicFunctions.Calendar
End Function

''
' Returns a reference used to access static methods of the <b>Console</b> class.
'
' @return A reference that can be used to access static methods of the <b>Console</b> class.
' @see Console
'
Public Function Console() As Console
    Set Console = modPublicFunctions.Console
End Function

''
' Returns a reference used to access static methods of the <b>ArrayList</b> class.
'
' @return A reference that can be used to access static methods of the <b>ArrayList</b> class.
' @see ArrayList
' @see ArrayListStatic
'
Public Function ArrayList() As ArrayListStatic
    Set ArrayList = modPublicFunctions.ArrayList
End Function

''
' Returns a reference used to access static methods of the <b>Registry</b> class.
'
' @return A reference that can be used to access static methods of the <b>Registry</b> class.
' @see Registry
' @see RegistryKey
'
Public Function Registry() As Registry
    Set Registry = modPublicFunctions.Registry
End Function

''
' Returns a reference used to access static methods of the <b>RegistryKey</b> class.
'
' @return A reference that can be used to access static methods of the <b>RegistryKey</b> class.
' @see RegistryKey
' @see RegistryKeyStatic
' @see Registry
'
Public Function RegistryKey() As RegistryKeyStatic
    Set RegistryKey = modPublicFunctions.RegistryKey
End Function

''
' Provides static methods for creating new <b>Guid</b> objects.
'
' @return A reference that can be used to access the static methods of the <b>Guid</b> class.
' @see GuidStatic
' @see Guid
'
Public Function Guid() As GuidStatic
    Set Guid = modPublicFunctions.Guid
End Function

''
' Provides methods for converting values.
'
' @return A reference to the method provider.
' @see Convert
'
Public Function Convert() As Convert
    Set Convert = modPublicFunctions.Convert
End Function

''
' Provides methods for creating custom resoure managers.
'
' @return A reference to the method provider.
' @see ResourceManagerStatic
' @see ResourceManager
' @see ResourceSet
' @see ResourceReader
'
Public Function ResourceManager() As ResourceManagerStatic
    Set ResourceManager = modPublicFunctions.ResourceManager
End Function

''
' Provides methods for retrieve information about all drives on the local machine.
'
' @return A reference to the method provider
' @see DriveInfo
' @see DirectoryInfo
' @see Directory
' @see FileInfo
' @see File
'
Public Function DriveInfo() As DriveInfoStatic
    Set DriveInfo = modPublicFunctions.DriveInfo
End Function


''
' Converts a Char value greater than 32768 to the corrisponding
' negative Integer value without overflow.
'
' @param Value The Long Ascii value to be converted.
' @return Integer version of the Long value.
' @remarks The Value can only range from -32767 to 65535, otherwise
' an <b>ArgumentOutOfRangeException</b> will be thrown.
'
Public Function LongToChar(ByVal Value As Long) As Integer
    If Value < -32767 Or Value > 65535 Then _
        Throw Cor.NewArgumentOutOfRangeException("Value must be between -32767 and 65535 inclusively.", "Value", Value)
    If Value > 32767 Then Value = Value - 65536
    LongToChar = Value
End Function

''
' Converts an Integer to the corrisponding Long Ascii char.
'
' @param Value An Integer character.
' @return A positive representation of the character.
'
Public Function CharToLong(ByVal Value As Integer) As Long
    AsWord(CharToLong) = Value
End Function

