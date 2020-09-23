VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Registry"
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
'    Module: Registry
'

''
' Supplies the base Registrykeys that access values and subkeys in the registry.
'
' @remarks These are the base sets of root registry keys used to access different
' classes of registry keys.
' <p>To access these methods, use the <b>Registry.*</b> syntax:
' <pre>
'     Dim rk As RegistryKey
'     Set rk = Registry.ClassesRoot
' </pres>
' @see RegistryKey
' @see RegistryKeyStatic
'
Option Explicit

Private mClassesRoot As RegistryKey
Private mCurrentConfig As RegistryKey
Private mCurrentUser As RegistryKey
Private mDynData As RegistryKey
Private mLocalMachine As RegistryKey
Private mPerformanceData As RegistryKey
Private mUsers As RegistryKey


''
' Returns a <b>RegistryKey</b> object that is then used to access registry keys and
' values for class types and their properties. The root section of the Registry
' accessed is HKEY_CLASSES_ROOT.
'
' @return Returns a <b>Registrykey</b> that is used to access the Class Root keys.
'
Public Property Get ClassesRoot() As RegistryKey
    Set ClassesRoot = mClassesRoot
End Property

''
' Returns a <b>RegistryKey</b> object that is then used to access registry keys and
' values for the current configuration of hardware. This is not user specific data.
' The root section of the Registry accessed is HKEY_CURRENT_CONFIG.
'
Public Property Get CurrentConfig() As RegistryKey
    Set CurrentConfig = mCurrentConfig
End Property

''
' Returns a <b>RegistryKey</b> object that is then used to access registry keys and
' values for information and preferences of the current user. The root section of
' the Registry accessed is HKEY_CURRENT_USER.
'
Public Property Get CurrentUser() As RegistryKey
    Set CurrentUser = mCurrentUser
End Property

''
' Contains dynamic registry data. This field reads the Windows registry base key HKEY_DYN_DATA.
'
' @remarks This root class is only available of Windows 95/98/ME. An exception
' is thrown for other windows platforms.
'
Public Property Get DynData() As RegistryKey
    Set DynData = mDynData
End Property

''
' Returns a <b>RegistryKey</b> object that can be used to access information and preferences
' about the local user machine. The root section of the Registry accessed is HKEY_LOCAL_MACHINE.
'
Public Property Get LocalMachine() As RegistryKey
    Set LocalMachine = mLocalMachine
End Property

''
' Contains performance information for software components. This field reads the Windows registry
' base key HKEY_PERFORMANCE_DATA.
'
Public Property Get PerformanceData() As RegistryKey
    Set PerformanceData = mPerformanceData
End Property

''
' Contains information about the default user configuration. This field reads the Windows
' registry base key HKEY_USERS.
'
Public Property Get Users() As RegistryKey
    Set Users = mUsers
End Property


Public Function GetValue(ByVal KeyName As String, ByVal ValueName As String, Optional ByRef DefaultValue As Variant) As Variant
    Dim rk As RegistryKey
    
    Set rk = GetRegistryKey(KeyName)
    Helper.MoveVariant GetValue, rk.GetValue(ValueName, DefaultValue)
End Function

Public Sub SetValue(ByVal KeyName As String, ByVal ValueName As String, ByRef Value As Variant)
    Dim rk As RegistryKey
    
    Set rk = GetRegistryKey(KeyName)
    If rk Is Nothing Then _
        Throw Cor.NewArgumentException("Could not save to registry key", "KeyName")
    
    rk.SetValue ValueName, Value
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetRegistryKey(ByVal KeyName As String) As RegistryKey
    Dim rk As RegistryKey
    Dim i As Long
    
    i = InStr(KeyName, "\")
    If i = 0 Then _
        Throw Cor.NewArgumentException("Invalid Key Format.", "KeyName")
    
    Select Case UCase$(Left$(KeyName, i - 1))
        Case "HKEY_CLASSES_ROOT":       Set rk = Me.ClassesRoot
        Case "HKEY_CURRENT_CONFIG":     Set rk = Me.CurrentConfig
        Case "HKEY_CURRENT_USER":       Set rk = Me.CurrentUser
        Case "HKEY_DYN_DATA":           Set rk = Me.DynData
        Case "HKEY_LOCAL_MACHINE":      Set rk = Me.LocalMachine
        Case "HKEY_PERFORMANCE_DATA":   Set rk = Me.PerformanceData
        Case "HKEY_USERS":              Set rk = Me.Users
        Case Else
            Throw Cor.NewArgumentException("Invalid Root Key.", "KeyName")
    End Select
    
    Set GetRegistryKey = rk.CreateSubKey(Mid$(KeyName, i))
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mClassesRoot = Cor.NewRegistryKey(HKEY_CLASSES_ROOT, "", 0, "", True)
    Set mCurrentConfig = Cor.NewRegistryKey(HKEY_CURRENT_CONFIG, "", 0, "", True)
    Set mCurrentUser = Cor.NewRegistryKey(HKEY_CURRENT_USER, "", 0, "", True)
    Set mDynData = Cor.NewRegistryKey(HKEY_DYN_DATA, "", 0, "", True)
    Set mLocalMachine = Cor.NewRegistryKey(HKEY_LOCAL_MACHINE, "", 0, "", True)
    Set mPerformanceData = Cor.NewRegistryKey(HKEY_PERFORMANCE_DATA, "", 0, "", True)
    Set mUsers = Cor.NewRegistryKey(HKEY_USERS, "", 0, "", True)
End Sub
