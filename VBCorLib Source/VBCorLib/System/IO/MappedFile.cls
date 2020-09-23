VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MappedFile"
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
'    Module: MappedFile
'

''
' Represents a simple file-to-memory mapping.
'
' @remarks Though this class is not part of .NET it is included
' to allow for easy access to mapping a file into memory.
'
Option Explicit
Implements cObject

Private mFileName As String
Private mFileHandle As Handle
Private mMapHandle As Handle
Private mViewHandle As Handle


''
' Returns the base address for the mapped view.
'
' @return Value of the base address.
' @remarks The BaseAddress is the starting memory location that
' the mapped file begins. Each byte in memory corrisponds to the
' same byte in the file.
'
Public Property Get BaseAddress() As Long
    If Not IsOpen Then _
        Throw Cor.NewInvalidOperationException("No file is mapped.")
    BaseAddress = mViewHandle
End Property

''
' Closes the currently mapped file.
'
Public Sub CloseFile()
    If IsOpen Then
        If mViewHandle <> 0 Then UnmapViewOfFile mViewHandle
        If mMapHandle <> 0 Then CloseHandle mMapHandle
        If mFileHandle <> 0 Then CloseHandle mFileHandle
        mFileHandle = 0
    End If
End Sub

''
' Returns value indicating if the mapped file is open.
'
' @return Status of the file.
'
Public Property Get IsOpen() As Boolean
    IsOpen = (mFileHandle <> 0)
End Property

''
' Returns the filename of the currently mapped file.
'
' @return The currently mapped file.
'
Public Property Get FileName() As String
    FileName = mFileName
End Property

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
'
Public Function Equals(ByRef Value As Variant) As Boolean
    If IsObject(Value) Then
        Equals = (Value Is Me)
    End If
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
'
Public Function GetHashCode() As Long
    GetHashCode = ObjPtr(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal FileName As String)
    mFileName = FileName
    mFileHandle = CreateFile(FileName, GENERIC_READ, 0, ByVal 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    
    If mFileHandle = INVALID_HANDLE Then IOError Err.LastDllError, FileName
    
    mMapHandle = CreateFileMapping(mFileHandle, 0, PAGE_READONLY, 0, 0, vbNullString)
    mViewHandle = MapViewOfFile(mMapHandle, FILE_MAP_READ, 0, 0, 0)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
    If IsOpen Then CloseFile
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
