Attribute VB_Name = "modIOError"
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
'    Module: moIOError
'

''
' Throws the appropriate exception based on the windows IO error.
'
Option Explicit

Public Sub IOError(ByVal E As Long, Optional ByVal src As String)
    Dim ex As IOException
    
    Select Case E
        Case ERROR_PATH_NOT_FOUND
            Set ex = Cor.NewDirectoryNotFoundException("The directory '" & src & "' could not be found.")
        Case ERROR_FILE_NOT_FOUND
            Set ex = Cor.NewFileNotFoundException(, src)
        Case ERROR_ACCESS_DENIED
            Set ex = Cor.NewInvalidOperationException("Permission to the specified file is denied.")
        Case Else
            Set ex = Cor.NewIOException(GetErrorMessage(E))
    End Select
    ex.HResult = E
    ex.Source = src
    Throw ex
End Sub
