VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Win32Resource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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
'    Module: Win32Resource
'

''
' Contains details about a resource.
'
' @remarks When iterating resources, a Win32Resource object is used to contain the
' information describing the resource found.
'
' @see Constructors
' @see ResourceReader
' @see WinResourceReader
'
Option Explicit
Implements cObject

''
' The type of resources acknowledged by Windows. These are used
' when a resource type is requested in function calls.
'
' @param CursorResource A Windows cursor resource type.
' @param BitmapResource A bitmap resource type.
' @param IconResource An icon resource type.
' @param MenuResource A menu set resource type.
' @param DialogBox A dialog box resource type.
' @param StringResource A String resource type.
' @param FontDirectory A font directory resource type.
' @param FontResource A font resource type.
' @param AcceleratorTable An accelerator table resource type.
' @param UserDefined A user-defined resource type.
' @param GroupCursor A group cursor resource type.
' @param GroupIcon A group icon resource type.
' @param VersionResource A version resource type.
' @param DialogInclude A dialog include resource type.
' @param PlugPlay A plug-n-play resource type.
' @param VXD A VXD resource type.
' @param AniCursor An animated cursor resource type.
' @param AniIcon An animated icon resource type.
' @param HTML An html resource type.
'
Public Enum ResourceTypes
    CursorResource = 1
    BitmapResource = 2
    IconResource = 3
    MenuResource = 4
    DialogBox = 5
    StringResource = 6
    FontDirectory = 7
    FontResource = 8
    AcceleratorTable = 9
    UserDefined = 10
    GroupCursor = 12
    GroupIcon = 14
    VersionResource = 16
    DialogInclude = 17
    PlugPlay = 19
    VXD = 20
    AniCursor = 21
    AniIcon = 22
    HTML = 23
End Enum

Private mType As ResourceTypes
Private mTypeName As String
Private mOrdinal As Long
Private mName As String
Private mLangID As Long
Private mValue As Variant


''
' Returns the language for the resource.
'
' @return The language ID, also called the Culture ID.
'
Public Property Get LanguageID() As Long
    LanguageID = mLangID
End Property

''
' Returns the name of the resource.
'
' @return The name of the resource.
' @remarks This value may be an empty string.
'
Public Property Get Name() As String
    Name = mName
End Property

''
' Returns the Ordinal or ID of the resource.
'
' @return The ordinal or ID.
'
Public Property Get Ordinal() As Long
    Ordinal = mOrdinal
End Property

''
' Returns the name of the resource type.
'
' @return The name of the resource type.
' @remarks This value may be an empty string.
'
Public Property Get ResourceTypeName() As String
    ResourceTypeName = mTypeName
End Property

''
' Returns the type of this resource.
'
' @return The type of resource.
'
Public Property Get ResourceType() As ResourceTypes
    ResourceType = mType
End Property

''
' Returns the resource value.
'
' @return The resource value.
'
Public Property Get Value() As Variant
    VariantCopyInd Value, mValue
End Property

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Public Function ToString() As String
    ToString = cString.Format("ResType: {0}, ResTypeName: {1}, ID: {2}, Name: {3}, LangID: {4}", mType, mTypeName, mOrdinal, mName, mLangID)
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
Friend Sub Init(ByVal ResType As ResourceTypes, ByVal ResTypeName As String, ByVal ResOrdinal As Long, ByVal ResName As String, ByVal LangID As Long, ByRef Value As Variant)
    mType = IIf(ResType = 0, ResourceTypes.UserDefined, ResType)
    mTypeName = ResTypeName
    mOrdinal = ResOrdinal
    mName = ResName
    mLangID = LangID
    VariantCopyInd mValue, Value
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
