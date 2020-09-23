Attribute VB_Name = "modWeakReferenceHelpers"
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
'    Module: modWeakReferenceHelpers
'
Option Explicit

Private Type VTableType
    VTable(2) As Long
End Type
Public Type WeakRefHookType
    VTable As VTableType
    pOwner As Long
    pTarget As Long
    pOriginalVTable As Long
End Type

Private Type WeakSafeArray
    SA As SafeArray1d
    WeakRef() As WeakRefHookType
End Type

Private mWeak As WeakSafeArray
Private mRelease As Long

Public Sub InitWeakReference(ByRef ref As WeakRefHookType, ByRef Owner As WeakReference, ByVal Target As IUnknown)
    If mRelease = 0 Then
        mRelease = FuncAddr(AddressOf WeakReferenceRelease)
        With mWeak
            With .SA
                .cbElements = LenB(ref)
                .cDims = 1
                .cElements = 1
            End With
            SAPtr(.WeakRef) = VarPtr(.SA)
        End With
    End If

    Dim p As Long
    With ref
        p = VTablePtr(Target)
        With .VTable
            .VTable(0) = MemLong(p)
            .VTable(1) = MemLong(p + 4)
            .VTable(2) = mRelease
        End With
        .pOriginalVTable = p
        .pOwner = ObjectPtr(Owner)
        .pTarget = ObjectPtr(Target)
        p = MemLong(VarPtr(Target))
        Set Target = Nothing
        MemLong(p) = VarPtr(ref)
    End With
End Sub

Public Sub DisposeWeakReference(ByRef ref As WeakRefHookType)
    With ref
        If .pOriginalVTable <> 0 Then
            MemLong(.pTarget) = .pOriginalVTable
            .pOriginalVTable = 0
            .pTarget = 0
            .pOwner = 0
        End If
    End With
End Sub

Public Function WeakReferenceRelease(ByRef this As Long) As Long
    Dim tmpThis As Long
    Dim Target As IVBUnknown
    Dim Owner As WeakReference
    
    tmpThis = this
    mWeak.SA.pvData = this
    
    On Error GoTo errTrap
    With mWeak.WeakRef(0)
        this = .pOriginalVTable
        
        ObjectPtr(Target) = .pTarget
        WeakReferenceRelease = Target.Release
        ObjectPtr(Target) = 0
        
        ObjectPtr(Owner) = .pOwner
        Owner.Release WeakReferenceRelease
        ObjectPtr(Owner) = 0
        
        If WeakReferenceRelease > 0 Then
            this = tmpThis
        Else
            .pTarget = 0
            .pOwner = 0
        End If
    End With
    Exit Function
    
errTrap:
    ObjectPtr(Target) = 0
    ObjectPtr(Owner) = 0
    WeakReferenceRelease = 0
    ForceDispose this
End Function

' we use a separate sub to prevent WeakType from being
' allocated everytime WeakReferenceRelease is called.
' This is just to optimize memory allocation/deallocation.
Private Sub ForceDispose(ByVal this As Long)
    Dim WeakType As WeakRefHookType
        
    CopyMemory WeakType, ByVal this, LenB(WeakType)
    DisposeWeakReference WeakType
End Sub

