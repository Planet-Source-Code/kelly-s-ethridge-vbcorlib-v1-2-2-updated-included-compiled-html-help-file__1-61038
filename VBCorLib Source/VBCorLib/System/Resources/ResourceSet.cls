VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResourceSet"
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
'    Module: ResourceSet
'

''
' Stores all the resources localized for one particular culture, ignoring all other cultures.
'
' @remarks This can be constructed using a .RES filename, a .RES formatted Stream, or
' an <b>IResourceReader</b> object.
' @see Constructors
' @see ResourceReader
' @see WinResourceReader
' @see ResourceManager
' @see Win32Resource
'
Option Explicit
Implements cObject
Implements IEnumerable

Private mResources As Hashtable
Private mReader As IResourceReader
Private mIsOpen As Boolean
Private mLangID As Long


''
' Returns the Culture ID for the specific set of resources.
'
' @return The culture ID.
'
Public Property Get LanguageID() As Long
    LanguageID = mLangID
End Property

''
' Returns the resource specified.
'
' @param ID The resource identifier.
' @param ResType The type of resource to search for.
' @return The resource value, or Empty if the resource was not found.
' @remarks The ResType can be either a String of the type of resource or
' a ResourceTypes enum value.
'
Public Function GetObject(ByVal ID As Variant, ByVal ResType As Variant) As Variant
    Dim ht As Hashtable
    Dim res As Win32Resource
    
    VerifyOpen
    
    ' A Hashtable considers Longs, Integers, and Bytes as different hashcode types
    ' so we merge them all into a Long type here.
    Select Case VarType(ID)
        Case vbLong, vbInteger, vbByte: VariantType(ID) = vbLong
    End Select
    
    Set ht = GetResource(ResType, False)
    If ht Is Nothing Then Exit Function
    
    If ht.Contains(ID) Then
        Set res = ht(ID)
        Helper.MoveVariant GetObject, res.Value
    End If
End Function

''
' Returns a String resource from the set.
'
' @param ID The string identifier.
' @return A resource string, or an empty string if not found.
'
Public Function GetString(ByVal ID As Variant) As String
    On Error Resume Next
    GetString = GetObject(ID, StringResource)
End Function

''
' Returns an enumerator to iterate through all resources
' within the resource set.
'
' @return An enumerator.
' @remarks The enumerator returns values as <b>DictionaryEntry</b>
' objects. The value property in the <b>DictionaryEntry</b> object
' returns a <b>Win32Resource</b> object which contains details about
' the specific resource found in the .RES file.
' The <i>Key</b> property returns the ID for the specific resource.
'
Public Function GetEnumerator() As IDictionaryEnumerator
    VerifyOpen
    
    Dim ret As New ResourceSetEnumerator
    ret.Init Me
    Set GetEnumerator = ret
End Function

''
' Returns an enumerator to iterate through all resources
' within the resource set.
'
' @return An enumerator.
' @remarks The enumerator returns values as <b>DictionaryEntry</b>
' objects. The value property in the <b>DictionaryEntry</b> object
' returns a <b>Win32Resource</b> object which contains details about
' the specific resource found in the .RES file.
' The <i>Key</b> property returns the ID for the specific resource.
'
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = CreateEnumerator(GetEnumerator)
End Function

''
' Returns the name of the type of resource reader to use with
' this specifici <b>ResourceSet</b>
'
' @return The name of a resource reader type.
'
Public Function GetDefaultReader() As String
    GetDefaultReader = "VBCorLib.ResourceReader"
End Function

''
' Returns the name of the type of resource writer to use with
' this specifici <b>ResourceSet</b>
'
' @return The name of a resource writer type.
'
Public Function GetDefaultWriter() As String
    ' Until I figure out how to create a .RES file this
    ' returns nothing.
End Function

''
' Closes the resource set.
'
Public Sub CloseResourceSet()
    If mIsOpen Then
        If Not mReader Is Nothing Then mReader.CloseReader
        mIsOpen = False
    End If
End Sub

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
Friend Sub Init(ByRef CultureID As Variant, ByRef Source As Variant)
    Dim res As Win32Resource
    Dim de As DictionaryEntry
    Dim LangID As Long
    
    If IsObject(Source) Then
        If TypeOf Source Is IResourceReader Then
            Set mReader = Source
        End If
    End If
    If mReader Is Nothing Then Set mReader = Cor.NewResourceReader(Source)
    
    LangID = GetLanguageID(CultureID)
    
    For Each de In mReader
        Set res = de.Value
        If res.LanguageID = LangID Then
            If Len(res.ResourceTypeName) > 0 Then
                GetResource(res.ResourceTypeName, True).Add res.Ordinal, res
            Else
                GetResource(res.ResourceType, True).Add res.Ordinal, res
            End If
        End If
    Next de
End Sub

Friend Sub AddResource(ByVal res As Win32Resource)
    Dim ht As Hashtable
    
    Set ht = GetResource(res.ResourceType, True)
    Set ht(res.Ordinal) = res
End Sub

Friend Function GetResourceSet() As Hashtable
    Set GetResourceSet = mResources
End Function

Friend Property Let LanguageID(ByVal RHS As Long)
    mLangID = RHS
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetResource(ByRef ResType As Variant, ByVal AddMissing As Boolean) As Hashtable
    VerifyResType ResType
    
    If mResources.Contains(ResType) Then
        Set GetResource = mResources(ResType)
    ElseIf AddMissing Then
        Set GetResource = New Hashtable
        mResources.Add ResType, GetResource
    End If
End Function

Private Sub VerifyResType(ByRef ResType As Variant)
    Select Case VarType(ResType)
        Case vbString: Exit Sub
        Case vbLong, vbInteger, vbByte
            Select Case ResType
                Case 1 To 10, 12, 14, 16, 17, 19 To 23
                Case Else
                    Throw Cor.NewArgumentOutOfRangeException("ResType is not a value ResourceType enum value", "ResType", ResType)
            End Select
        Case Else
            Throw Cor.NewArgumentException("Resource type must be a String or ResourceType enum.", "ResType")
    End Select
End Sub

Private Sub VerifyOpen()
    If Not mIsOpen Then Throw Cor.NewInvalidOperationException("ResourceSet is closed.")
End Sub

Private Function GetLanguageID(ByRef CultureID As Variant) As Long
    If IsMissing(CultureID) Then
        GetLanguageID = CultureInfo.CurrentCulture.LCID
    Else
        Select Case VarType(CultureID)
            Case vbObject
                If TypeOf CultureID Is CultureInfo Then
                    Dim ci As CultureInfo
                    Set ci = CultureID
                    GetLanguageID = ci.LCID
                Else
                    Throw Cor.NewArgumentException("CultureInfo object required.", "CultureID")
                End If
            Case vbLong, vbInteger, vbByte
                GetLanguageID = CultureID
            Case Else
                Throw Cor.NewArgumentException("CultureInfo object or Language ID required.")
        End Select
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mResources = Cor.NewHashtable(hcp:=New CaseInsensitiveHashCodePrvdr, comparer:=New CaseInsensitiveComparer)
    mIsOpen = True
End Sub

Private Sub Class_Terminate()
    CloseResourceSet
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
