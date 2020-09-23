Attribute VB_Name = "modBufferHelper"
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
'    Module: modBufferHelper
'
Option Explicit

Public Type WordBuffer
    pVTable As Long
    this As IUnknown
    pRelease As Long
    data() As Integer
    SA As SafeArray1d
End Type

Private mpRelease As Long

Public Sub InitWordBuffer(ByRef Buffer As WordBuffer, ByVal pData As Long, ByVal Length As Long)
    If mpRelease = 0 Then mpRelease = FuncAddr(AddressOf WordBuffer_Release)
    With Buffer.SA
        .cbElements = 2
        .cDims = 1
        .cElements = Length
        .pvData = pData
    End With
    With Buffer
        .pVTable = VarPtr(.pVTable)
        .pRelease = mpRelease
        SAPtr(.data) = VarPtr(.SA)
        ObjectPtr(.this) = VarPtr(.pVTable)
    End With
End Sub

Private Function WordBuffer_Release(ByRef this As WordBuffer) As Long
    SAPtr(this.data) = 0
    this.SA.pvData = 0
End Function

