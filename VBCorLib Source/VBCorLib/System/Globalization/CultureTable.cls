VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CultureTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
'    Module: CultureTable
'
Option Explicit

'   vbcultures.nlp structure
'
'''' Header ''''
'   Culture Count                           :   4 bytes
'   StringPool offset from base address     :   4 bytes
'   Culture Size                            :   4 bytes
'
'''' N Cultures ''''
'   Integer array
'    13 numeric values          :   2 bytes per value
'   Long array
'    44 pointers to string pool :   4 bytes per value
'
'''' StringPool ''''
'   Unknown number of strings.
'   Any strings that are to be
'   referenced as arrays of strings
'   has a 2 byte value indicating
'   the array size preceeding the
'   array elements.

Private Const CULTURE_FILENAME          As String = "vbcultures.nlp"
Private Const CULTURE_COUNT_OFFSET      As Long = 0
Private Const STRINGPOOL_OFFSET_OFFSET  As Long = 4
Private Const INVARIANT_LCID            As Long = 127
Private Const zh_CHS_LCID               As Long = 4
Private Const zh_CHT_LCID               As Long = 31748
Private Const sr_LCID                   As Long = 31770

' If the layout of a culture changes, then this
' structure needs to be updated to match.
Private Type CultureHeader
    Numbers(12) As Integer
    Strings(43) As Long
End Type


Private mFile As MappedFile
Private mCulturesByLCID As Hashtable
Private mNamesToLCID As Hashtable
Private mBaseAddress As Long
Private mStringPoolAddress As Long
Private mCulture() As CultureHeader
Private mCultureSA As SafeArray1d



Friend Property Get IsLoaded() As Boolean
    IsLoaded = (mBaseAddress <> 0)
End Property

Friend Function GetString(ByVal LCID As Integer, ByVal Index As Long) As String
    Dim pos As Long
    
    On Error GoTo errTrap
    mCultureSA.pvData = mCulturesByLCID(LCID)
    pos = mStringPoolAddress + mCulture(0).Strings(Index)
    GetString = SysAllocString(pos)
    Exit Function
    
errTrap:
    Throw Cor.NewArgumentException("Culture is not supported.", "LCID")
End Function

Friend Function GetStringArray(ByVal LCID As Integer, ByVal Index As Long) As String()
    Dim pos As Long
    Dim ret() As String
    Dim cnt As Long
    Dim i As Long
    
    On Error GoTo errTrap
    mCultureSA.pvData = mCulturesByLCID(LCID)
    pos = mStringPoolAddress + mCulture(0).Strings(Index)
    cnt = MemWord(pos)
    pos = pos + 2
    
    ReDim ret(0 To cnt - 1)
    For i = 0 To cnt - 1
        ret(i) = SysAllocString(pos)
        pos = pos + LenB(ret(i)) + 2
    Next i
    
    GetStringArray = ret
    Exit Function
    
errTrap:
    Throw Cor.NewArgumentException("Culture is not supported.", "LCID")
End Function

Friend Function GetNumber(ByVal LCID As Integer, ByVal Index As Long) As Integer
    On Error GoTo errTrap
    mCultureSA.pvData = mCulturesByLCID(LCID)
    GetNumber = mCulture(0).Numbers(Index)
    Exit Function
    
errTrap:
    Throw Cor.NewArgumentException("Culture is not supported.", "LCID")
End Function

Friend Function GetNumberArray(ByVal LCID As Integer, ByVal Index As Long) As Long()
    Dim pos As Long
    Dim ret() As Long
    Dim cnt As Long
    Dim i As Long
    Dim s As String
    
    On Error GoTo errTrap
    mCultureSA.pvData = mCulturesByLCID(LCID)
    pos = mStringPoolAddress + mCulture(0).Strings(Index)
    cnt = MemWord(pos)
    pos = pos + 2
    
    ReDim ret(0 To cnt - 1)
    For i = 0 To cnt - 1
        s = SysAllocString(pos) ' MemLong(pos) - 48
        ret(i) = s
        pos = pos + LenB(s) + 2
    Next i
    
    GetNumberArray = ret
    Exit Function
    
errTrap:
    Throw Cor.NewArgumentException("Culture is not supported.", "LCID")
End Function

Friend Function GetCultureID(ByVal Name As String) As Long
    Dim ID As Variant
    
    If mNamesToLCID Is Nothing Then
        Set mNamesToLCID = New Hashtable
        For Each ID In mCulturesByLCID.Keys
            mNamesToLCID.Add GetString(ID, SNAME), ID
        Next ID
    End If
    ID = mNamesToLCID(Name)
    If Not IsEmpty(ID) Then GetCultureID = ID
End Function

Friend Function IsNeutral(ByVal LCID As Long) As Boolean
    If GetSubLangID(LCID) = 0 Or LCID = sr_LCID Then
        IsNeutral = True
    Else
        Select Case LCID
            Case INVARIANT_LCID, zh_CHS_LCID, zh_CHT_LCID: IsNeutral = True
        End Select
    End If
End Function

Friend Function GetCultures(ByVal Types As CultureTypes) As CultureInfo()
    Dim ID As Variant
    Dim getNeutral As Boolean
    Dim getSpecific As Boolean
    Dim getInstalled As Boolean
    Dim list As ArrayList
    Dim ret() As CultureInfo
    
    getNeutral = CBool(Types And NeutralCultures)
    getSpecific = CBool(Types And SpecificCultures)
    getInstalled = CBool(Types And InstalledWin32Cultures)
    
    Set list = New ArrayList
    For Each ID In mCulturesByLCID.Keys
        If IsNeutral(ID) Then
            If getNeutral Then list.Add Cor.NewCultureInfo(ID)
        ElseIf IsInstalled(ID) And getInstalled Then
            list.Add Cor.NewCultureInfo(ID)
        ElseIf getSpecific Then
            list.Add Cor.NewCultureInfo(ID)
        End If
    Next ID
    ReDim ret(0 To list.Count - 1)
    list.CopyTo ret
    GetCultures = ret
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsInstalled(ByVal LCID As Long) As Boolean
    IsInstalled = IsValidLocale(LCID, LCID_INSTALLED)
End Function

Private Function GetSubLangID(ByVal LCID As Long) As Long
    GetSubLangID = Helper.ShiftRight(LCID, 10) And &H3F
End Function

Private Sub Load()
    Dim s As String
    Dim p As String
    
    On Error GoTo errTrap
    
    p = App.Path & "\"
    s = Dir$(p & CULTURE_FILENAME)
    If Len(s) = 0 Then
        p = Environment.GetFolderPath(Windows) & "\"
        s = Dir$(p & CULTURE_FILENAME)
    End If
    If Len(s) = 0 Then Exit Sub
    Set mFile = Cor.NewMappedFile(p & s)
    mBaseAddress = mFile.BaseAddress
    mStringPoolAddress = mBaseAddress + MemLong(mBaseAddress + STRINGPOOL_OFFSET_OFFSET)
    LoadLCIDLookup
errTrap:
End Sub

Private Sub LoadLCIDLookup()
    Dim i As Long
    Dim pos As Long
    Dim Size As Long
    
    Set mCulturesByLCID = New Hashtable
    i = MemLong(mBaseAddress + CULTURE_COUNT_OFFSET)
    Size = MemLong(mBaseAddress + 8)
    pos = mBaseAddress + 12
    mCultureSA.cbElements = Size
    
    Do While i > 0
        mCultureSA.pvData = pos
        mCulturesByLCID.Add mCulture(0).Numbers(ILCID), pos
        pos = pos + Size
        i = i - 1
    Loop
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    With mCultureSA
        .cDims = 1
        .cElements = 1
    End With
    SAPtr(mCulture) = VarPtr(mCultureSA)
    Load
End Sub

Private Sub Class_Terminate()
    SAPtr(mCulture) = 0
End Sub
