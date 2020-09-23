Attribute VB_Name = "modExceptionMethods"
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
'    Module: modExceptionMethods
'
Option Explicit

Private mException As Exception
Private mErrorMessage As String



Public Function Catch(ByRef ex As Exception, Optional ByVal Err As ErrObject) As Boolean
    If Not mException Is Nothing Then
        Set ex = mException
        Set mException = Nothing
        Catch = True
    ElseIf Not Err Is Nothing Then
        If Err.Number Then
            Set ex = Cor.NewException(Err.Description)
            ex.HResult = Err.Number
            ex.Source = Err.Source
            Err.Clear
            Catch = True
        End If
    End If
    VBA.Err.Clear
End Function

Public Sub Throw(ByVal ex As Exception)
    Set mException = ex
    Err.Raise ex.HResult, ex.Source, ex.Message
End Sub

Public Sub ClearException()
    Set mException = Nothing
End Sub

Public Function GetErrorMessage(ByVal MessageID As Long) As String
    If LenB(mErrorMessage) = 0 Then mErrorMessage = String$(1024, vbNullChar)
    GetErrorMessage = cString.TrimEnd(left$(mErrorMessage, FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, MessageID, 1033, mErrorMessage, Len(mErrorMessage), 0)))
End Function
