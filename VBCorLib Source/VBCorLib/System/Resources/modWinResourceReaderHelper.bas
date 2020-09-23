Attribute VB_Name = "modWinResourceReaderHelper"
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
'    Module: modWinResourceReader
'

Option Explicit

Public Function EnumResTypeProc(ByVal hModule As Long, ByVal lpszType As Long, ByRef Reader As WinResourceReader) As Long
    EnumResTypeProc = EnumResourceNames(hModule, lpszType, AddressOf EnumResNameProc, VarPtr(Reader))
End Function

Private Function EnumResNameProc(ByVal hModule As Long, ByVal lpszType As Long, ByVal lpszName As Long, ByRef Reader As WinResourceReader) As Long
    EnumResNameProc = EnumResourceLanguages(hModule, lpszType, lpszName, AddressOf EnumResLangProc, VarPtr(Reader))
End Function

Private Function EnumResLangProc(ByVal hModule As Long, ByVal lpszType As Long, ByVal lpszName As Long, ByVal wIDLanguage As Integer, ByRef Reader As WinResourceReader) As Long
    If wIDLanguage <> 0 Then
        Dim h As Long
        h = FindResourceEx(hModule, lpszType, lpszName, wIDLanguage)
        If h <> 0 Then
            Dim h2 As Long
            h2 = LoadResource(hModule, h)
            If h2 <> 0 Then
                Dim lpData As Long
                Dim l As Long
                Dim b() As Byte
                Dim ResType As Long
                Dim ResTypeName As String
                Dim ResName As String
                Dim ResOrdinal As Long
                
                lpData = LockResource(h2)
                l = SizeofResource(hModule, h)
                ReDim b(0 To l - 1)
                CopyMemory b(0), ByVal lpData, l
                
                If lpszType And &HFFFF0000 Then
                    ResTypeName = GetString(lpszType)
                Else
                    ResType = lpszType
                End If
                
                If lpszName And &HFFFF0000 Then
                    ResName = GetString(lpszName)
                Else
                    ResOrdinal = lpszName
                End If
                
                Reader.AddResource ResType, ResTypeName, ResOrdinal, ResName, wIDLanguage, b
                
            End If
        End If
    End If
    
    EnumResLangProc = BOOL_TRUE
End Function



Private Function GetString(ByVal lpsz As Long) As String
    Dim b() As Byte
    Dim l As Long
    
    l = lstrlen(lpsz)
    ReDim b(0 To l - 1)
    CopyMemory b(0), ByVal lpsz, l
    GetString = StrConv(b, vbUnicode)
End Function
