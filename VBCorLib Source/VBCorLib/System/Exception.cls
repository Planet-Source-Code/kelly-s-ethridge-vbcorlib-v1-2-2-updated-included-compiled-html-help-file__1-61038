VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Exception"
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
'    Module: Exception
'

''
' This is the base class from which all exception derive their interface.
'
' @remarks This class is the base class for all exceptions. When an
' error occurs, either the system or the currently executing application reports
' it by throwing an exception containing information about the error. Once thrown,
' an exception is handled by the application or by the default exception handler.
'
' There are two methods of raising an error:
' Using the classical VB approach Err.Raise. This method does not place an
' exception object in holding to be caught. In order to catch the error using
' the Catch method, call Catch passing in an exception variable and also the
' Err object. This will create an Exception object with the Err information.
' Use this method if an error might be raised using standard Err.Raise methods.
'
' Using the VBCorLib approach of Throw <exception object>. This method does
' place the exception in holding to be caught. It then raises an error using
' the standard Err.Raise method giving the information within the exception
' object. This allows for the error to be caught by error traps not utilzing the
' Throw/Catch style. To catch the exception using Catch, an error trap is still
' required. Once in the trap call Catch passing in an exception variable that is
' set to the held exception object. Catch will return True if an exception was
' caught, False otherwise.
'
' Example:
'   This throws a specific exception type. When an exception is caught, an
'   If/Else can be used to determine if the exception was an expected type
'   or not.
'
' <pre>
'   Public Sub ThrowCatch()
'       On Error Goto CatchIt
'       Throw Cor.NewArgumentNullException("A null argument message")
'       Err.Raise 5,,"Something else."
'       Exit sub
'
'   CatchIt:
'       Dim ex as Exception
'       ' Err is passed because the potential that the error was raised
'       ' not using Throw, in which case we want to catch it the same.
'       If Catch(ex, Err) then
'           If TypeOf ex Is ArgumentNullException then
'               MsgBox "Your argument was null."
'           Else
'               MsgBox "Some other exception occurred: " & ex.Message
'           End If
'       End If
'       Resume Next
'   End Sub
' </pre>
'
Option Explicit
Implements cObject

Private Const DEF_MESSAGE            As String = "An exception of type VBCorLib.Exception has occured."
Private Const DEF_SOURCE             As String = ""
Private Const DEF_HELPLINK           As String = ""
Private Const PROP_MESSAGE           As String = "Message"
Private Const PROP_INNEREXCEPTION    As String = "InnerException"
Private Const PROP_HRESULT           As String = "HResult"
Private Const PROP_SOURCE            As String = "Source"
Private Const PROP_HELPLINK          As String = "HelpLink"


Private mMessage As String
Private mInnerException As Exception
Private mHResult As Long
Private mSource As String
Private mHelpLink As String




''
' Gets a link to the help file associated with this exception.
'
' @return The help file assigned to this exception.
Public Property Get HelpLink() As String
    HelpLink = mHelpLink
End Property

''
' Sets a link to the help file associated with this exception.
'
' @param RHS The help file to assign to this exception.
Public Property Let HelpLink(ByVal RHS As String)
    mHelpLink = RHS
End Property

''
' Gets the HRESULT, a coded numerical value that is assigned to a specific exception.
'
' @return The value of the associated HResult.
' @remarks An HResult is associated with an error result code. This allows for VB specific
' error codes to be returned.
Public Property Get HResult() As Long
    HResult = mHResult
End Property

''
' Sets the HRESULT, a coded numerical value that is assigned to a specific exception.
'
' @param RHS The value of the associated HResult.
' @remarks An HResult is associated with an error result code. This allows for VB specific
' error codes to be returned.
Public Property Let HResult(ByVal RHS As Long)
    mHResult = RHS
End Property

''
' Gets the name of the application or the object that causes the error.
'
' @return A string representation of the exception source.
Public Property Get Source() As String
    Source = mSource
End Property

''
' Sets the name of the application or the object that causes the error.
'
' @param RHS A string representation of the exception source.
Public Property Let Source(ByVal RHS As String)
    mSource = RHS
End Property

''
' Gets a message that describes the current exception.
'
' @return A string representation of the exception.
' @remarks The message is a human-readable representation of the exception.
Public Property Get Message() As String
    Message = mMessage
End Property

''
' Gets the Exception instance that caused the current exception.
'
' @return The exception that caused this exception.
Public Property Get InnerException() As Exception
    Set InnerException = mInnerException
End Property

''
' Returns the Exception that is the root cause of one or more subsequent exceptions.
'
' @return The base exception from which this chain of exceptions originated.
Public Function GetBaseException() As Exception
    If mInnerException Is Nothing Then
        Set GetBaseException = Me
    Else
        Set GetBaseException = mInnerException.GetBaseException
    End If
End Function

''
' Returns a string representation of the current exception.
'
' @return A human-readable representation of the exception.
Public Function ToString() As String
    ToString = App.Title & "." & TypeName(Me) & ": " & Message
End Function

''
' Determines whether two Object variables are the same instance.
'
' @param value The value to compare this instance to.
' @return Boolean value indicating equality.
Public Function Equals(ByRef Value As Variant) As Boolean
    If IsObject(Value) Then
        Equals = (Value Is Me)
    End If
End Function

''
' Returns a pseudo-unique number for the current object instance.
'
' @return Value representitive of a pseudo-unique number.
Public Function GetHashCode() As Long
    GetHashCode = ObjPtr(Me)
End Function

''
' Initializes an instance with user-defined information. Called by NewException function.
Friend Sub Init(ByRef Message As String, ByRef InnerException As Exception)
    mMessage = Message
    Set mInnerException = InnerException
End Sub

' Initializes the object to default values.
Private Sub Class_InitProperties()
    Init DEF_MESSAGE, Nothing
    mHResult = COR_E_EXCEPTION
End Sub

' Serializes the object.
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        mMessage = .ReadProperty(PROP_MESSAGE, DEF_MESSAGE)
        Set mInnerException = .ReadProperty(PROP_INNEREXCEPTION, Nothing)
        mHResult = .ReadProperty(PROP_HRESULT, COR_E_EXCEPTION)
        mSource = .ReadProperty(PROP_SOURCE, DEF_SOURCE)
        mHelpLink = .ReadProperty(PROP_HELPLINK, DEF_HELPLINK)
    End With
End Sub

' Deserializes the object.
Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty PROP_MESSAGE, mMessage, DEF_MESSAGE
        .WriteProperty PROP_INNEREXCEPTION, mInnerException, Nothing
        .WriteProperty PROP_HRESULT, mHResult, COR_E_EXCEPTION
        .WriteProperty PROP_SOURCE, mSource, DEF_SOURCE
        .WriteProperty PROP_HELPLINK, mHelpLink, DEF_HELPLINK
    End With
End Sub




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   cObject interface
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
