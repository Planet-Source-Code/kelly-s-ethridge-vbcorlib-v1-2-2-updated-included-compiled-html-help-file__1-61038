VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WeakReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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
'    Module: WeakReference
'

''
' Provides a reference to an object, but does not keep that object alive.
'
' @remarks
' This is to help with circular references to prevent unreachable objects from
' remaining in memory. The number of references is examined after each call to
' an object's IUnknown.Release method. Once it reaches zero, the WeakReference
' object clears the hook into the referenced object and sets the IsAlive status to false.
'
' @see Constructors
'
Option Explicit
Implements cObject

Private mHook As WeakRefHookType



''
' Returns a strong reference to the object a weak reference is being kept to.
'
' @return Strong referenced object.
Public Property Get Target() As Object
    If Not IsAlive Then _
        Throw Cor.NewInvalidOperationException("Object reference has already been terminated.")
    
    Dim ret As IVBUnknown
    ObjectPtr(ret) = mHook.pTarget
    Set Target = ret
    ObjectPtr(ret) = 0
End Property

''
' Sets the target object to keep a weak reference to.
'
' @param RHS Strong object to create a weak reference to.
Public Property Set Target(ByVal RHS As Object)
    DisposeWeakReference mHook
    Init RHS
End Property

''
' Returns the status of the weakly referenced object.
'
' @return Status of the weak referenced object.
Public Property Get IsAlive() As Boolean
    IsAlive = (mHook.pOriginalVTable <> 0)
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
Friend Sub Init(ByVal Target As Object)
    InitWeakReference mHook, Me, Target
End Sub

' callback used in modWeakReferenceHelpers
Friend Sub Release(ByVal cRefs As Long)
    If cRefs = 0 Then DisposeWeakReference mHook
End Sub

Private Sub Class_Terminate()
    DisposeWeakReference mHook
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
