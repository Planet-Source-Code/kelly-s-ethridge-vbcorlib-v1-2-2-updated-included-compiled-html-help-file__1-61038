VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResourceReader"
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
'    Module: ResourceReader
'

''
' Provides a means to iterate through the resources in a .RES file.
'
' @remarks The only resource types recognized are Strings, Icons, Cursors, and Bitmaps.
' All other resource types are contained as a byte array.
' <p>Pass this object into the <b>ResourceManager</b> constructor to access specific
' resources based on ID, Type, and Culture. Or pass this object into a <b>ResourceSet</b>
' to retrieve the resources for a specific culture.</p>
'
' @see Constructors
' @see ResourceManager
' @see ResourceSet
' @see IResourceReader
' @see Win32Resource
'
Option Explicit
Implements cObject
Implements IEnumerable
Implements IResourceReader


Private mBase As ResourceReaderBase
Private mStream As Stream
Private mSB As StringBuilder


''
' Returns an IDictionaryEnumerator object.
'
' @return An enumerator.
' @remarks The enumerator returns values as <b>DictionaryEntry</b>
' objects. The value property in the <b>DictionaryEntry</b> object
' returns a <b>Win32Resource</b> object which contains details about
' the specific resource found in the .RES file.
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
' Returns a For..Each compatible enumerator.
'
' @return Enumerator object.
' @remarks The enumerator returns values as <b>DictionaryEntry</b>
' objects. The value property in the <b>DictionaryEntry</b> object
' returns a <b>Win32Resource</b> object which contains details about
' the specific resource found in the .RES file.
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
' Closes the internal Stream.
'
Public Sub CloseReader()
    If Not mStream Is Nothing Then
        mStream.CloseStream
        Set mStream = Nothing
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
Friend Sub Init(ByRef Source As Variant)
    Select Case VarType(Source)
        Case vbString
            Set mStream = Cor.NewFileStream(Source, FileMode.OpenExisting)
            mStream.Position = 32
        Case vbObject
            If TypeOf Source Is Stream Then
                Set mStream = Source
                If Not mStream.CanRead Then Throw Cor.NewArgumentException("The Stream must support reading.", "Source")
            Else
                Throw Cor.NewArgumentException("Source must implement Stream interface", "Source")
            End If
        Case Else
            Throw Cor.NewArgumentException("Source must be a file path or file stream.", "Source")
    End Select
    LoadResources
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadResources()
    Dim DataSize As Long
    Dim ResType As ResourceTypes
    Dim ResTypeName As String
    Dim ResOrdinal As Long
    Dim ResName As String
    Dim LangID As Integer
    Dim Data() As Byte
    Dim r As BinaryReader
    
    Set r = Cor.NewBinaryReader(mStream)
    
    Do
        On Error GoTo ErrorEndOfStream
        DataSize = r.ReadLong
        
        On Error GoTo ErrorInvalidResStream
        r.ReadLong  ' header size
        
        GetResOrdinalAndName r, ResType, ResTypeName
        GetResOrdinalAndName r, ResOrdinal, ResName
        DWordAlignStream mStream
        
        r.ReadLong    ' data version, usually 0
        r.ReadInteger ' memory flags
        
        LangID = r.ReadInteger
        
        r.ReadLong    ' version, usually 0
        r.ReadLong    ' characteristics, usually 0
        
        ReDim Data(0 To DataSize - 1)
        If r.ReadBuffer(Data, 0, DataSize) <> DataSize Then Throw New Exception    ' gets us to ErrorInvalidResStream
        
        DWordAlignStream mStream
        
        If LangID <> 0 Then mBase.AddResource ResType, ResTypeName, ResOrdinal, ResName, LangID, Data
    Loop
    
ErrorEndOfStream:
    ' normal way to finish the stream
    Dim ex As Exception
    If Catch(ex, Err) Then
        If TypeOf ex Is EndOfStreamException Then Exit Sub
    End If

ErrorInvalidResStream:
    On Error GoTo 0
    Throw Cor.NewArgumentException("Stream does not contain a valid .RES format.", "Source")
End Sub

Private Sub GetResOrdinalAndName(ByVal r As BinaryReader, ByRef Ordinal As Long, ByRef Name As String)
    Dim ch As Integer
    
    ch = r.ReadInteger
    If ch = -1 Then 'ordinal type
        Ordinal = r.ReadInteger
        Name = ""
    Else 'type name
        If mSB Is Nothing Then Set mSB = New StringBuilder
        mSB.Length = 0
        
        Do While ch <> 0
            mSB.AppendChar ch
            ch = r.ReadInteger
        Loop
        Name = mSB.ToString
        Ordinal = 0
    End If
End Sub

Private Sub DWordAlignStream(ByVal s As Stream)
    Dim i As Long
    i = s.Position
    If i Mod 4 <> 0 Then s.SeekPosition i Mod 4, SeekOrigin.FromCurrent
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mBase = New ResourceReaderBase
End Sub


Private Sub Class_Terminate()
    CloseReader
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
