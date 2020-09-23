VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CultureInfoStatic"
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
'    Module: CultureInfoStatic
'

''
' Provides a set of static methods retrieve CultureInfo objects.
'
' @see CultureInfo
'
Option Explicit

Private Declare Function GetUserDefaultUILanguage Lib "kernel32.dll" () As Long
Private Declare Function GetSystemDefaultUILanguage Lib "kernel32.dll" () As Long

Public Enum CultureTypes
    NeutralCultures = 1
    SpecificCultures = 2
    InstalledWin32Cultures = 4
    AllCultures = 7
End Enum

Private mCurrentCulture As CultureInfo
Private mInvariantCulture As CultureInfo
Private mCurrentUICulture As CultureInfo
Private mInstalledUICulture As CultureInfo



''
' Returns a ReadOnly wrapped CultureInfo object.
'
' @param Culture The culture to create a ReadOnly wrapper for.
' @return The wrapped ReadOnly culture.
'
Public Function ReadOnly(ByVal Culture As CultureInfo) As CultureInfo
    If Culture.IsReadOnly Then
        Set ReadOnly = Culture
    Else
        Set ReadOnly = Culture.Clone
        ReadOnly.IsReadOnly = True
    End If
End Function

''
' Returns a set of CultureInfo objects.
'
' @param Types The culture types that are to be retrieved.
' @return An array of CultureInfo objects based on Types specified.
'
Public Function GetCultures(ByVal Types As CultureTypes) As CultureInfo()
    GetCultures = CultureTable.GetCultures(Types)
End Function

''
' Returns the current culture associated with this machine.
'
' @return A ReadOnly version of CultureInfo specific for this machine.
'
Public Property Get CurrentCulture() As CultureInfo
    If mCurrentCulture Is Nothing Then
        Set mCurrentCulture = Cor.NewCultureInfo(GetSystemDefaultLCID)
        mCurrentCulture.IsReadOnly = True
    End If
    Set CurrentCulture = mCurrentCulture
End Property

''
' Returns a ReadOnly version of the invariant culture.
'
' @return A ReadOnly version of the invariant culture.
'
Public Property Get InvariantCulture() As CultureInfo
    If mInvariantCulture Is Nothing Then
        Set mInvariantCulture = Cor.NewCultureInfo(INVARIANT_LCID)
        mInvariantCulture.IsReadOnly = True
    End If
    Set InvariantCulture = mInvariantCulture
End Property

''
' Returns the culture for the current thread.
'
' @return A <b>CultureInfo</b> object for the current thread.
' @remarks The culture is determined by the language ID of the
' current thread by calling <b>GetUserDefaultUILanguage</b> API.
'
Public Property Get CurrentUICulture() As CultureInfo
    If mCurrentUICulture Is Nothing Then
        Set mCurrentUICulture = Cor.NewCultureInfo(GetUserDefaultUILanguage)
        mCurrentUICulture.IsReadOnly = True
    End If
    Set CurrentUICulture = mCurrentUICulture
End Property

''
' Returns the culture for the current systems language.
'
' @return A <b>CultureInfo</b> object for the current system language.
' @remarks The current system language is detemined by calling the
' <b>GetSystemDefaultUILanguage</b> API.
'
Public Property Get InstalledUICulture() As CultureInfo
    If mInstalledUICulture Is Nothing Then
        Set mInstalledUICulture = Cor.NewCultureInfo(GetSystemDefaultUILanguage)
        mInstalledUICulture.IsReadOnly = True
    End If
    Set InstalledUICulture = mInstalledUICulture
End Property


