VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WinResourceReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
'    Module: WinResourceReader
'

''
' A class that reads resources out of EXE and DLL files.
'
' @remarks Use this class to iterate the resources within an EXE or DLL file.
' <p>The only resource types recognized are Strings, Bitmaps, Icons and Cursors.
' All other resource types are maintained as byte arrays.</p>
' <p>Pass this object into the <b>ResourceManager</b> constructor to retrieve
' specific resources based on ID, Type, and Culture. Or, pass this object into
' a <b>ResourceSet</b> to retrieve resources for a specific culture.</p>
'
' @see Constructors
' @see ResourceManager
' @see ResourceSet
' @see ResourceReader
' @see IResourceReader
' @see Win32Resource
'
Option Explicit
Implements cObject
Implements IEnumerable
Implements IResourceReader


Private mBase As ResourceReaderBase

''
' Closes the reader
Public Sub CloseReader()
    ' do nothing
End Sub

''
' Returns an IDictionaryEnumerator object.
'
' @return An enumerator.
' @remarks The enumerator returns values as <b>DictionaryEntry</b>
' objects. The value property in the <b>DictionaryEntry</b> object
' returns a <b>Win32Resource</b> object which contains details about
' the specific resource found in the EXE or DLL file.
' The <i>Key</b> property returns the ID for the specific resource.
' The Keys may not be unique across all resources, os using the Key
' as a unique identifier is not recommended.
'
Public Function GetEnumerator() As IDictionaryEnumerator
    Dim ret As New ResourceEnumerator
    ret.Init Me, mBase.GetResources
    Set GetEnumerator = ret
End Function

''
' Returns an IDictionaryEnumerator object.
'
' @return An enumerator.
' @remarks The enumerator returns values as <b>DictionaryEntry</b>
' objects. The value property in the <b>DictionaryEntry</b> object
' returns a <b>Win32Resource</b> object which contains details about
' the specific resource found in the EXE or DLL file.
' The <i>Key</b> property returns the ID for the specific resource.
' The Keys may not be unique across all resources, os using the Key
' as a unique identifier is not recommended.
'
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = CreateEnumerator(GetEnumerator)
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
' @param value The value to compare equalit to.
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
Friend Sub Init(ByVal FileName As String)
    Dim h As Long
    
    h = LoadLibrary(FileName)
    If h = 0 Then _
        Throw Cor.NewIOException("Invalid Application file (" & FileName & ").")
    
    If EnumResourceTypes(h, AddressOf EnumResTypeProc, VarPtr(Me)) = BOOL_FALSE Then
        FreeLibrary h
        IOError Err.LastDllError, FileName
    End If
    
    FreeLibrary h
    
End Sub

Friend Sub AddResource(ByVal ResType As ResourceTypes, ByVal ResTypeName As String, ByVal ResOrdinal As Long, ByVal ResName As String, ByVal LangID As Long, ByRef Data() As Byte)
    mBase.AddResource ResType, ResTypeName, ResOrdinal, ResName, LangID, Data
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mBase = New ResourceReaderBase
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
'   IEnumerable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IEnumerable_GetEnumerator() As IEnumerator
    Set IEnumerable_GetEnumerator = GetEnumerator
End Function

Private Function IEnumerable_NewEnum() As stdole.IUnknown
    Set IEnumerable_NewEnum = NewEnum
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IResourceReader Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub IResourceReader_CloseReader()
    CloseReader
End Sub

Private Function IResourceReader_GetEnumerator() As IDictionaryEnumerator
    Set IResourceReader_GetEnumerator = GetEnumerator
End Function

Private Function IResourceReader_NewEnum() As stdole.IUnknown
    Set IResourceReader_NewEnum = NewEnum
End Function
