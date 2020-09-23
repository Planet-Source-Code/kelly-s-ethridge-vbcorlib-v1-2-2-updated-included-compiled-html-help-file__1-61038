VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResourceManager"
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
'    Module: ResourceManager
'

''
' Provides convenient access to culture-specific resources at runtime.
'
' @remarks A ResourceManager contains all resources for all of the cultures found
' using an IResourceReader object or a .RES filename.
' <p>To load the resources from multiple .RES files, use <b>ResourceManager.CreateFileBasedResourceManager</b>.</p>
' @see Constructors
' @see ResourceManagerStatic
' @see ResourceReader
' @see WinResourceReader
' @see IResourceReader
' @see Win32Resource
'
Option Explicit
Implements cObject
Implements IEnumerable

Private mResourceSets As Hashtable


''
' Returns a resource from the specified culture.
'
' @param ID The resource identifier.
' @param ResType The type of resource to search for.
' @param CultureID A CultureInfo object or Culture ID. The current culture is used if this parameter is missing.
' @return The resource value, or Empty if the resourse was not found.
'
Public Function GetObject(ByVal ID As Variant, ByVal ResType As Variant, Optional ByVal CultureID As Variant) As Variant
    Dim LangID As Long
    Dim rs As ResourceSet
    Dim c As CultureInfo
    
    LangID = GetLanguageID(CultureID)
    Set rs = InternalGetResourceSet(LangID, False)
    If rs Is Nothing Then
        Set c = Cor.NewCultureInfo(LangID)
        Set c = c.Parent
        
        Do While (Not c.Equals(CultureInfo.InvariantCulture)) And (Not c.IsNeutralCulture)
            If mResourceSets.Contains(c.LCID) Then
                Set rs = mResourceSets(c.LCID)
                Exit Do
            End If
            Set c = c.Parent
        Loop
    End If
    
    If Not rs Is Nothing Then Helper.MoveVariant GetObject, rs.GetObject(ID, ResType)
End Function

''
' Returns a resource string for the specific culture.
'
' @param ID The resource ID.
' @param CultureID A CultureInfo object or Culture ID. The current culture is used if this parameter is missing.
' @return The resource as a String, or an empty string if the resource was not found.
'
Public Function GetString(ByVal ID As Variant, Optional ByVal CultureID As Variant) As String
    On Error Resume Next
    GetString = GetObject(ID, StringResource, CultureID)
End Function

''
' Returns a set of resources for a specific culture.
'
' @param CultureID A CultureInfo object or Culture ID. The current culture is used if this parameter is missing.
' @return A set of resources for a specific culture, or Nothing if not found.
'
Public Function GetResourceSet(ByVal CultureID As Variant) As ResourceSet
    Set GetResourceSet = InternalGetResourceSet(GetLanguageID(CultureID), False)
End Function

''
' Releases all resources.
'
Public Sub ReleaseAllResources()
    Set mResourceSets = Nothing
End Sub

''
' Returns an enumerator used to iterate through all resources for all cultures.
'
' @return An IDictionaryEnumerator object.
' @remarks The enumerator returns <b>DictionaryEntry</b> objects. The <i>Key</i>
' contains the Culture ID. The <i>Value</i> contains the <b>ResourceSet</b> for
' that specific culture.
'
Public Function GetEnumerator() As IDictionaryEnumerator
    Set GetEnumerator = mResourceSets.GetEnumerator
End Function

''
' Returns an enumerator used to iterate through all resources for all cultures.
'
' @return An IDictionaryEnumerator object.
' @remarks The enumerator returns <b>DictionaryEntry</b> objects. The <i>Key</i>
' contains the Culture ID. The <i>Value</i> contains the <b>ResourceSet</b> for
' that specific culture.
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
Friend Sub Init(ByRef Source As Variant)
    Dim de As DictionaryEntry
    Dim rs As ResourceSet
    Dim res As Win32Resource
    Dim r As IResourceReader
    
    If IsObject(Source) Then
        If TypeOf Source Is IResourceReader Then Set r = Source
    End If
    If r Is Nothing Then Set r = Cor.NewResourceReader(Source)
    
    For Each de In r
        Set res = de.Value
        Set rs = InternalGetResourceSet(res.LanguageID, True)
        rs.AddResource res
    Next de
End Sub

Friend Sub InitFromFolder(ByVal SearchPattern As String, ByVal ResourceDir As String)
    Dim Files() As String
    Dim i As Long
    
    Files = Directory.GetFiles(ResourceDir, SearchPattern)
    
    For i = 0 To UBound(Files)
        Init Path.Combine(ResourceDir, Files(i))
    Next i
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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

Private Function InternalGetResourceSet(ByVal LangID As Long, ByVal AddMissing As Boolean) As ResourceSet
    If mResourceSets.Contains(LangID) Then
        Set InternalGetResourceSet = mResourceSets(LangID)
    ElseIf AddMissing Then
        Set InternalGetResourceSet = New ResourceSet
        InternalGetResourceSet.LanguageID = LangID
        mResourceSets.Add LangID, InternalGetResourceSet
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mResourceSets = New Hashtable
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
