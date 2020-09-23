VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArgumentException"
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
'    Module: ArgumentException
'

''
' The exception that is thrown when one of the arguments provided to a method is not valid.
'
' @remarks
' <p><b>ArgumentException</b> is thrown when a method is invoked and at least one of the passed
' arguments does not meet the parameter specification of the called method. When throwing an
' <b>ArgumentException</b> a meaningful error message describing the invalid argument, the name
' of the argument, as well as the expected range of values for the argument should be defined.</p>
' <p>New exception classes derived from <b>ArgumentException</b> can be created starting with the
' <b>ArgumentExceptionBase</b> template class and modifying it accordingly.</p>
' <p><b>ArgumentException</b> uses the HResult COR_E_ARGUMENT, which is &H5.</p>
'
' @see Constructors
' @see ExceptionMethods
' @see Exception
'
Option Explicit
Implements cObject
Implements Exception
Implements SystemException

Private Const PROP_PARAMNAME         As String = "ParamName"
Private Const PROP_BASE              As String = "Base"
Private Const DEF_PARAMNAME          As String = ""
Private Const DEF_HRESULT            As Long = COR_E_ARGUMENT

Private mBase As Exception
Private mParamName As String



''
' Returns the parameter name that caused the exception.
'
' @return The name of the parameter that caused the exception.
' @remarks
' <p>Every <b>ArgumentException</b> should have a parameter name supplied.</p>
' <p>This property returns the same value as the value passed into the constructor.</p>
'
' @see Constructors
'
Public Property Get ParamName() As String
    ParamName = mParamName
End Property

''
' Returns the name of the help file associated with this exception.
'
' @return The help file assigned to this exception.
'
Public Property Get HelpLink() As String
    HelpLink = mBase.HelpLink
End Property

''
' Sets the name of the help file associated with this exception.
'
' @param RHS The help file to assign to this exception.
'
Public Property Let HelpLink(ByVal RHS As String)
    mBase.HelpLink = RHS
End Property

''
' Gets the HRESULT, a coded numerical value that is assigned to a specific exception.
'
' @return The value of the associated HResult.
' @remarks An HResult is associated with an error result code. This allows for VB specific
' error codes to be returned.
'
Public Property Get HResult() As Long
    HResult = mBase.HResult
End Property

''
' Sets the HRESULT, a coded numerical value that is assigned to a specific exception.
'
' @param RHS The value of the associated HResult.
' @remarks An HResult is associated with an error result code. This allows for VB specific
' error codes to be returned. This is the same as <b>Err.Number</b>.
'
Public Property Let HResult(ByVal RHS As Long)
    mBase.HResult = RHS
End Property

''
' Gets the name of the application or the object that causes the error.
'
' @return A string representation of the exception source.
'
Public Property Get Source() As String
    Source = mBase.Source
End Property

''
' Sets the name of the application or the object that causes the error.
'
' @param RHS A string representation of the exception source.
'
Public Property Let Source(ByVal RHS As String)
    mBase.Source = RHS
End Property

''
' Gets a message that describes the current exception.
'
' @return A string representation of the exception.
' @remarks The message is a human-readable representation of the exception.
' This is set in the <b>NewArgumentException</b> constructor.
'
Public Property Get Message() As String
    Message = mBase.Message
    If Len(mParamName) > 0 Then Message = Message & vbCrLf & vbCrLf & "Parameter Name: " & mParamName
End Property

''
' Gets the Exception instance that caused the current exception.
'
' @return The exception that caused this exception.
' @remarks This is set in the <b>NewArgumentException</b> constructor.
'
Public Property Get InnerException() As Exception
    Set InnerException = mBase.InnerException
End Property

''
' Returns the Exception that is the root cause of one or more subsequent exceptions.
'
' @return The base exception from which this chain of exceptions originated.
' @remarks The base exception is the first exception to be thrown in
' a chain of exceptions. When throwing a new exception, an existing exception
' can be supplied as an inner-exception, similar to having a stack trace.
'
Public Function GetBaseException() As Exception
    Set GetBaseException = mBase.GetBaseException
    If GetBaseException Is mBase Then
        Set GetBaseException = Me
    End If
End Function

''
' Returns a string representation of the current exception.
'
' @return A human-readable representation of the exception.
' @see cObject
'
Public Function ToString() As String
    ToString = App.Title & "." & TypeName(Me) & ": " & Message
End Function

''
' Determines whether two Object variables are the same instance.
'
' @param value The value to compare this instance to.
' @return Boolean value indicating equality.
' @see cObject
'
Public Function Equals(ByRef Value As Variant) As Boolean
    If IsObject(Value) Then
        Equals = (Value Is Me)
    End If
End Function

''
' Returns a pseudo-unique number for the current object instance.
'
' @return Value representitive of a pseudo-unique number.
' @see cObject
'
Public Function GetHashCode() As Long
    GetHashCode = ObjPtr(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef Message As String, ByRef ParamName As String, ByRef InnerException As Exception)
    Set mBase = Cor.NewException(Message, InnerException)
    mParamName = ParamName
    mBase.HResult = DEF_HRESULT
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_InitProperties()
    Init Environment.GetResourceString(Argument_Exception), DEF_PARAMNAME, Nothing
End Sub

' Serializes the object.
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Set mBase = .ReadProperty(PROP_BASE, Cor.NewException(Environment.GetResourceString(Argument_Exception), Nothing))
        mParamName = .ReadProperty(PROP_PARAMNAME, DEF_PARAMNAME)
    End With
End Sub

' Deserializes the object.
Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty PROP_BASE, mBase
        .WriteProperty PROP_PARAMNAME, mParamName, DEF_PARAMNAME
    End With
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
'   Exception Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Exception_Equals(Value As Variant) As Boolean
    Exception_Equals = Equals(Value)
End Function

Private Function Exception_GetBaseException() As Exception
    Set Exception_GetBaseException = GetBaseException
End Function

Private Function Exception_GetHashCode() As Long
    Exception_GetHashCode = GetHashCode
End Function

Private Property Let Exception_HelpLink(ByVal RHS As String)
    HelpLink = RHS
End Property

Private Property Get Exception_HelpLink() As String
    Exception_HelpLink = HelpLink
End Property

Private Property Let Exception_HResult(ByVal RHS As Long)
    HResult = RHS
End Property

Private Property Get Exception_HResult() As Long
    Exception_HResult = HResult
End Property

Private Property Get Exception_InnerException() As Exception
    Set Exception_InnerException = InnerException
End Property

Private Property Get Exception_Message() As String
    Exception_Message = Message
End Property

Private Property Let Exception_Source(ByVal RHS As String)
    Source = RHS
End Property

Private Property Get Exception_Source() As String
    Exception_Source = Source
End Property

Private Function Exception_ToString() As String
    Exception_ToString = ToString
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   SystemException Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SystemException_Equals(Value As Variant) As Boolean
    SystemException_Equals = Equals(Value)
End Function

Private Function SystemException_GetBaseException() As Exception
    Set SystemException_GetBaseException = GetBaseException
End Function

Private Function SystemException_GetHashCode() As Long
    SystemException_GetHashCode = GetHashCode
End Function

Private Property Let SystemException_HelpLink(ByVal RHS As String)
    HelpLink = RHS
End Property

Private Property Get SystemException_HelpLink() As String
    SystemException_HelpLink = HelpLink
End Property

Private Property Let SystemException_HResult(ByVal RHS As Long)
    HResult = RHS
End Property

Private Property Get SystemException_HResult() As Long
    SystemException_HResult = HResult
End Property

Private Property Get SystemException_InnerException() As Exception
    Set SystemException_InnerException = InnerException
End Property

Private Property Get SystemException_Message() As String
    SystemException_Message = Message
End Property

Private Property Let SystemException_Source(ByVal RHS As String)
    Source = RHS
End Property

Private Property Get SystemException_Source() As String
    SystemException_Source = Source
End Property

Private Function SystemException_ToString() As String
    SystemException_ToString = ToString
End Function


