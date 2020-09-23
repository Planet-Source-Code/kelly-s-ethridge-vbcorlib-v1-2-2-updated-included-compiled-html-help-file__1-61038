VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Path"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
'    Module: Path
'

''
' Provides functions to locate and format path information.
'
' @remarks All members of Path are static. To access them, use the
' Path.* convention.
'
Option Explicit

Private Const ALT_DIRECTORY_SEPARATOR_CHAR      As Integer = 47
Private Const ALT_DIRECTORY_SEPARATOR_STRING    As String = "/"
Private Const DIRECTORY_SEPARATOR_CHAR          As Integer = 92
Private Const DIRECTORY_SEPARATOR_STRING        As String = "\"
Private Const PATH_SEPARATOR                    As Integer = 59
Private Const PATH_SEPARATOR_STRING             As String = ";"
Private Const VOLUME_SEPARATOR_CHAR             As Integer = 58
Private Const VOLUME_SEPARATOR_STRING           As String = ":"
Private Const FILE_EXTENSION_CHAR               As Integer = 46
Private Const MAX_PATH                          As Long = 260
Private Const DOUBLE_BACKSLASH                  As Long = &H5C005C
Private Const DOUBLE_FORSLASH                   As Long = &H2F002F
Private Const BACK_FORSLASH                     As Long = &H5C002F
Private Const FOR_BACKSLASH                     As Long = &H2F005C



Private mInvalidPathChars() As Integer
Private mBuffer As WordBuffer


''
' Returns the character code for the alternative directory separator character.
'
' @return The character code of the alternative directory separator (47).
' @remarks The alternate character is usually found on systems such as Unix.
'
Public Property Get AltDirectorySeparatorChar() As Integer
    AltDirectorySeparatorChar = ALT_DIRECTORY_SEPARATOR_CHAR
End Property

''
' Returns the string for the alternative directory separator character.
'
' @return The string of the alternative directory separator "/".
' @remarks The alternate string is usually found on systems such as Unix.
'
Public Property Get AltDirectorySeparatorString() As String
    AltDirectorySeparatorString = ALT_DIRECTORY_SEPARATOR_STRING
End Property

''
' Returns the directory separator character for Windows.
'
' @return The directory separator character for Windows (92).
'
Public Property Get DirectorySeparatorChar() As Integer
    DirectorySeparatorChar = DIRECTORY_SEPARATOR_CHAR
End Property

''
' Returns the directory separator string for Windows.
'
' @return The directory separator string for windows "\".
'
Public Property Get DirectorySeparatorString() As String
    DirectorySeparatorString = DIRECTORY_SEPARATOR_STRING
End Property

''
' Returns an array of characters that are invalid in a Windows path string.
'
' @return An Integer array representing the character code of the the invalid characters.
'
Public Property Get InvalidPathChars() As Integer()
    InvalidPathChars = mInvalidPathChars
End Property

''
' The character code used to separate strings in an environment variable.
'
' @return The environment variable string separator character code (59).
'
Public Property Get PathSeparator() As Integer
    PathSeparator = PATH_SEPARATOR
End Property

''
' The string used to seperate strings in an environment variable.
'
' @return The environment variable string separator ";".
'
Public Property Get PathSeparatorString() As String
    PathSeparatorString = PATH_SEPARATOR_STRING
End Property

''
' The character code used to separate volume labels from path information.
'
' @return The character code for the volume separator (58).
'
Public Property Get VolumeSeparatorChar() As Integer
    VolumeSeparatorChar = VOLUME_SEPARATOR_CHAR
End Property

''
' Returns the string version of the volume separator.
'
' @return The string version of the volume separator ":".
'
Public Property Get VolumeSeparatorString() As String
    VolumeSeparatorString = VOLUME_SEPARATOR_STRING
End Property

''
' Replaces the extension of a filename with a new extension.
'
' @param sPath The filename with an optional path.
' @param Extension The new extension to change the filename's to.
' @return The filename with the new extension, including a path if supplied.
'
Public Function ChangeExtension(ByRef sPath As String, ByVal Extension As Variant) As String
    Dim pos As Long
    Dim foundExtensionChar As Boolean
    Dim ext As String
    
    pos = Len(sPath) - 1
    If pos < 0 Then Exit Function
    VerifyPath sPath
    
    ' call treats vbNullString differently than an empty string.
    Select Case VarType(Extension)
        Case vbString
            If LenB(Extension) = 0 Then
                ext = "."
            ElseIf Asc(Extension) <> FILE_EXTENSION_CHAR Then
                ext = "."
            End If
        Case vbNull
            Extension = ""
        Case Else
            Throw Cor.NewArgumentException("Extension can be either a String or Null.", "Extension")
    End Select
    
    Do While pos >= 0
        Select Case mBuffer.Data(pos)
            Case DIRECTORY_SEPARATOR_CHAR, ALT_DIRECTORY_SEPARATOR_CHAR, VOLUME_SEPARATOR_CHAR
                Exit Do
            Case FILE_EXTENSION_CHAR
                foundExtensionChar = True
                Exit Do
        End Select
        pos = pos - 1
    Loop
    
    If Not foundExtensionChar Then pos = Len(sPath)
    ChangeExtension = left$(sPath, pos) & (ext & Extension)
End Function

''
' Returns an indication if the path is rooted or not.
'
' @param sPath The path to check if is rooted.
' @return The indication of the path being rooted.
' @remarks <p>A rooted path is one that starts with a volume name,
' UNC server name, or directory separator. Such paths as C:\MyDir\File,
' \\Server\Share\File, and \MyFile are rooted.</p>
' <p>This function does not verify that the path actually exists.</p>
'
Public Function IsPathRooted(ByRef sPath As String) As Boolean
    If Len(sPath) = 0 Then Exit Function
    
    VerifyPath sPath
    
    Select Case mBuffer.Data(0)
        Case DIRECTORY_SEPARATOR_CHAR, ALT_DIRECTORY_SEPARATOR_CHAR
            IsPathRooted = True
        Case Else
            IsPathRooted = (mBuffer.Data(1) = VOLUME_SEPARATOR_CHAR)
    End Select
End Function

''
' Returns if a filename has an extension.
'
' @param sPath A filename and optional path to check for an extension.
' @return An indication of the filename having an extension.
' @remarks <p>The path string is search for a '.' that separates the
' the extension from the filename itself. If any of the directory separator
' characters or volume separator is encountered before a '.', then the
' filename is considered not to have an extension.
'
Public Function HasExtension(ByRef sPath As String) As Boolean
    Dim pos As Long
    Dim Size As Long
    Dim foundExtensionChar As Boolean
    
    Size = Len(sPath)
    If Size = 0 Then Exit Function
    
    VerifyPath sPath
    
    pos = Size - 1
    Do While pos >= 0
        Select Case mBuffer.Data(pos)
            Case FILE_EXTENSION_CHAR
                foundExtensionChar = True
                Exit Do
            Case DIRECTORY_SEPARATOR_CHAR, ALT_DIRECTORY_SEPARATOR_CHAR, VOLUME_SEPARATOR_CHAR
                Exit Do
        End Select
        pos = pos - 1
    Loop
    
    If foundExtensionChar Then HasExtension = (pos < Size - 1)
End Function

''
' Combintes two paths into a single path.
'
' @param Path1 The left-hand side of the two paths to be joined.
' @param Path2 The right-hand side of the two paths to be joined.
' @return The joined paths.
' @remarks <p>If Path1 is an empty string, then Path2 is returned as is. If Path2
' is empty, but not Path1, then Path1 is returned as is.</p>
' <p>If Path2 is rooted (starts with a volume, UNC identifier or directory
' separator), then Path2 is returned, regardless of Path1.</p>
' <p>If Path1 does not end with a directory separator, one will be inserted.</p>
' <p>This function does not verify that the paths exist.</p>
'
Public Function Combine(ByRef Path1 As String, ByRef Path2 As String) As String
    Dim Size As Long
    
    VerifyPath Path1
    
    If IsPathRooted(Path2) Then
        Combine = Path2
        Exit Function
    End If
    
    Size = Len(Path1)
    If Size = 0 Then
        Combine = Path2
        Exit Function
    ElseIf LenB(Path2) = 0 Then
        Combine = Path1
        Exit Function
    End If
    
    mBuffer.SA.pvData = StrPtr(Path1)
    Select Case mBuffer.Data(Size - 1)
        Case DIRECTORY_SEPARATOR_CHAR, ALT_DIRECTORY_SEPARATOR_CHAR, VOLUME_SEPARATOR_CHAR
            Combine = Path1 & Path2
        Case Else
            Combine = Path1 & DIRECTORY_SEPARATOR_STRING & Path2
    End Select
End Function

''
' Returns the directory portion of a file path.
'
' @param sPath The path to parse the directory from.
' @return The directory parsed from the path.
' @remarks The function does not verify that the directory exists.
'
Public Function GetDirectoryName(ByRef sPath As String) As String
    Const UNC_PATH  As Long = 1
    Const VOLUME_PATH As Long = 2
    Const RELATIVE_PATH As Long = 4
    
    Dim pos As Long
    Dim outpos As Long
    Dim pth As String
    Dim maxindex As Long
    Dim ch As Integer
    Dim pathtype As Long
    Dim requiredslash As Boolean
    Dim prevwasslash As Boolean
    Dim slashcount As Long
    Dim lastslash As Long
    
    
    pth = Trim$(sPath)
    VerifyPath pth
    
    maxindex = Len(pth) - 1
    If maxindex = 0 Then _
        Throw Cor.NewArgumentException("The path is not in a valid format.")
    
                    
    Do While pos <= maxindex
        ch = mBuffer.Data(pos)
        If ch = ALT_DIRECTORY_SEPARATOR_CHAR Then
            ch = DIRECTORY_SEPARATOR_CHAR
            mBuffer.Data(pos) = DIRECTORY_SEPARATOR_CHAR
        End If
        Select Case ch
            Case DIRECTORY_SEPARATOR_CHAR
                lastslash = outpos
                If pos = 0 Then
                    pathtype = UNC_PATH Or RELATIVE_PATH
                ElseIf pos = 1 Then
                    If pathtype And UNC_PATH Then
                        requiredslash = True
                        pathtype = UNC_PATH
                        prevwasslash = False
                    End If
                ElseIf Not prevwasslash Then
                    slashcount = slashcount + 1
                End If
                If prevwasslash Then
                    outpos = outpos - 1
                Else
                    prevwasslash = True
                End If
            Case Else
                If pos = 0 Then
                    Select Case ch
                        Case LOWER_A_CHAR To LOWER_Z_CHAR, UPPER_A_CHAR To UPPER_Z_CHAR
                        Case Else: Throw Cor.NewArgumentException("Invalid path format.", "sPath")
                    End Select
                ElseIf pos = 1 Then
                    If ch = CHAR_COLON Then pathtype = VOLUME_PATH
                End If
                prevwasslash = False
        End Select
        mBuffer.Data(outpos) = ch
        pos = pos + 1
        outpos = outpos + 1
    Loop
    
    If pathtype = VOLUME_PATH Then
        If maxindex = 2 Then
            GetDirectoryName = vbNullString
            Exit Function
        ElseIf lastslash = 2 Then
            lastslash = 3
        End If
    ElseIf pathtype = UNC_PATH Then
        If slashcount < 2 Then
            GetDirectoryName = vbNullString
            Exit Function
        End If
    End If
    GetDirectoryName = left$(pth, lastslash)
End Function

''
' Returns the extension of a filename.
'
' @param sPath The filename to retrieve the extension of.
' @return The extension of the filename
'
Public Function GetExtension(ByRef sPath As String) As String
    Dim pos As Long
    Dim foundExtensionChar As Boolean
    Dim StartIndex As Long
    
    VerifyPath sPath
    
    StartIndex = Len(sPath) - 1
    pos = StartIndex
    If pos < 0 Then Exit Function
    
    Do While pos >= 0
        Select Case mBuffer.Data(pos)
            Case FILE_EXTENSION_CHAR
                foundExtensionChar = True
                Exit Do
            Case DIRECTORY_SEPARATOR_CHAR, ALT_DIRECTORY_SEPARATOR_CHAR, VOLUME_SEPARATOR_CHAR
                Exit Do
        End Select
        pos = pos - 1
    Loop
    If Not foundExtensionChar Then Exit Function
    If pos = StartIndex Then Exit Function
    GetExtension = Mid$(sPath, pos + 1)
End Function

''
' Returns the filename portion of the specified path.
'
' @param sPath The path to retrieve the filename from.
' @return The filetime portion of the specified path.
' @remarks This function does not verify that the file exists.
'
Public Function GetFileName(ByRef sPath As String) As String
    Dim pos As Long
    
    VerifyPath sPath
    
    pos = Len(sPath) - 1
    If pos < 0 Then Exit Function
    
    Do While pos >= 0
        Select Case mBuffer.Data(pos)
            Case DIRECTORY_SEPARATOR_CHAR, ALT_DIRECTORY_SEPARATOR_CHAR, VOLUME_SEPARATOR_CHAR
                Exit Do
        End Select
        pos = pos - 1
    Loop
    GetFileName = Mid$(sPath, pos + 2)
End Function

''
' Returns the filename portion of a path without the filename extension.
'
' @param sPath The path to parse the filename from.
' @return A filename without the extension present.
' @remarks This function does not verify the file exists.
'
Public Function GetFileNameWithoutExtension(ByRef sPath As String) As String
    Dim pos As Long
    Dim extPos As Long
    
    VerifyPath sPath
    pos = Len(sPath) - 1
    If pos < 0 Then Exit Function
    
    Do While pos >= 0
        Select Case mBuffer.Data(pos)
            Case FILE_EXTENSION_CHAR
                If extPos = 0 Then extPos = pos
            Case DIRECTORY_SEPARATOR_CHAR, ALT_DIRECTORY_SEPARATOR_CHAR, VOLUME_SEPARATOR_CHAR
                Exit Do
        End Select
        pos = pos - 1
    Loop
    If extPos > 0 Then
        GetFileNameWithoutExtension = Mid$(sPath, pos + 2, extPos - pos - 1)
    Else
        GetFileNameWithoutExtension = Mid$(sPath, pos + 2)
    End If
End Function

''
' Creates a full path to the specified filename.
'
' @param sPath The filename to prepend any path to.
' @return A fully qualified path to a filename.
' @remarks The filename is prepended with the current directory to
' give a full path. The file does not need to exists, nor is the
' existence verified.
'
Public Function GetFullPath(ByRef sPath As String) As String
    
    Dim pth As String
    Dim ret As String
    Dim file As String
    Dim Size As Long
    Dim ptr As Long
    
    pth = cString.Trim(sPath)
    If Len(pth) = 0 Then _
        Throw Cor.NewArgumentException("The path is not in a valid format.")
    
    VerifyPath pth, True
    
    Select Case LCase$(left$(pth, 5))
        Case "http:", "file:":  Throw Cor.NewNotSupportedException("URI paths are not supported.")
    End Select
    
    ptr = StrPtr(pth)
    
    ' quickly check the first two charcters for possible UNC path.
    ' convert first to characters into a long.
    Select Case MemLong(ptr)
        Case DOUBLE_BACKSLASH, DOUBLE_FORSLASH, BACK_FORSLASH, FOR_BACKSLASH
            VerifyUNCPathFormat ptr
    End Select
    
    ' convert first two characters into two integers.
    With MemDWord(ptr)
        ' check first character being ':'
        If .LoWord = VOLUME_SEPARATOR_CHAR Then _
            Throw Cor.NewNotSupportedException("Invalid path format.")
        
        ' check if absolute path ( a:, b:, ...) is being specified.
        If .HiWord = VOLUME_SEPARATOR_CHAR Then
            Select Case .LoWord
                Case LOWER_A_CHAR To LOWER_Z_CHAR, UPPER_A_CHAR To UPPER_Z_CHAR
                Case Else
                    Throw Cor.NewArgumentException("Incorrect absolute path information.")
            End Select
        End If
    End With
    
    ' make sure no more than 1 colon exists.
    If InStr(3, pth, VOLUME_SEPARATOR_STRING) > 0 Then _
        Throw Cor.NewNotSupportedException("Path format is not supported.")
    
    ret = SysAllocStringLen(0, MAX_PATH)
    Size = GetFullPathName(pth, MAX_PATH, ret, file)
    
    If Size = 0 Or Size > MAX_PATH Then _
        Throw Cor.NewPathTooLongException("The qualified path is longer than the maximum of 260 characters.")
    
    GetFullPath = left$(ret, Size)
End Function

''
' Returns the root directory of the specified path.
'
' @param sPath The path to retrieve the root directory from.
' @return The root directory of the specified path.
' @remarks <i>sPath</i> is not checked if it exists. The root is
' derived from what information is in <i>sPath</i>.
'
Public Function GetPathRoot(ByRef sPath As String) As String
    Dim ptr As Long
    
    VerifyPath sPath
    
    If Len(sPath) = 0 Then _
        Throw Cor.NewArgumentException("The path is not in a valid format.")
    
    ptr = StrPtr(sPath)
    With MemDWord(ptr)
        Select Case .LoWord
            Case DIRECTORY_SEPARATOR_CHAR, ALT_DIRECTORY_SEPARATOR_CHAR
                Select Case .HiWord
                    Case DIRECTORY_SEPARATOR_CHAR, ALT_DIRECTORY_SEPARATOR_CHAR
                        GetPathRoot = GetUNCPathRoot(ptr)
                        Exit Function
                    Case VOLUME_SEPARATOR_CHAR
                        Throw Cor.NewArgumentException("Invalid path format.")
                    Case Else
                        GetPathRoot = "\"
                        Exit Function
                End Select
        End Select
        If .HiWord = VOLUME_SEPARATOR_CHAR Then
            Select Case .LoWord
                Case LOWER_A_CHAR To LOWER_Z_CHAR, UPPER_A_CHAR To UPPER_Z_CHAR
                    GetPathRoot = left$(sPath, 2) & DIRECTORY_SEPARATOR_STRING
                    Exit Function
                Case Else
                    Throw Cor.NewArgumentException("Invalid path format.")
            End Select
        End If
    End With
End Function

''
' Returns the filename for a temporary file.
'
' @return The temporary file's filename.
' @remark This function creates a 0 byte temporary file of the returned name.
' The file is placed in the temporary path of the current user.
'
Public Function GetTempFileName() As String
    Dim tmpFileName As String
    
    tmpFileName = String$(MAX_PATH, 0)
    If GetTempFileNameA(GetTempPath, "tmp", 0, tmpFileName) <> BOOL_FALSE Then
        GetTempFileName = SysAllocString(StrPtr(tmpFileName))
    End If
End Function

''
' Returns the path to the temporary folder for the current user.
'
' @return The temporary path for the current user.
'
Public Function GetTempPath() As String
    Dim tmpPath As String
    Dim Size As Long
    
    tmpPath = String$(MAX_PATH, 0)
    Size = GetTempPathA(MAX_PATH, tmpPath)
    If Size = 0 Or Size > MAX_PATH Then _
        Throw Cor.NewSystemException("Could not retrieve the temporary path.")
    
    GetTempPath = left$(tmpPath, Size)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub VerifyPath(ByRef sPath As String, Optional ByVal FixSlashes As Boolean = False)
    Dim i As Long
    
    If Len(sPath) > MAX_PATH Then _
        Throw New PathTooLongException
    
    mBuffer.SA.pvData = StrPtr(sPath)
    
    For i = 0 To Len(sPath) - 1
        Select Case mBuffer.Data(i)
            Case 34, 60, 62, 124, 0, 8, 16, 17, 18, 20, 21, 22, 23, 24, 25
                Throw Cor.NewArgumentException("Invalid characters in path.", "sPath")
            Case 47
                If FixSlashes Then mBuffer.Data(i) = 92    ' change / to \
        End Select
    Next i
End Sub

Friend Function InternalGetFullPath(ByVal Path As String) As String
    Dim i As Long
    
    VerifyPath Path
    If Len(Path) = 0 Then Path = CurDir
    mBuffer.SA.pvData = StrPtr(Path)
    i = Len(Path) - 1
    Select Case mBuffer.Data(i)
        Case ALT_DIRECTORY_SEPARATOR_CHAR
            mBuffer.Data(i) = DIRECTORY_SEPARATOR_CHAR
        Case Is <> DIRECTORY_SEPARATOR_CHAR
            Path = Path & DIRECTORY_SEPARATOR_STRING
    End Select
    InternalGetFullPath = Path
End Function

Friend Function InternalGetFullRootedPath(ByVal Path As String) As String
    If Not IsPathRooted(Path) Then
        Path = GetFullPath(Path)
    End If
    InternalGetFullRootedPath = InternalGetFullPath(Path)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetUNCPathRoot(ByVal ptr As Long) As String
    Const DOUBLE_BACKSLASH As Long = &H5C005C
    
    Dim pos As Long
    Dim ch As Integer
    Dim foundSeparator As Boolean
    Dim ret As String
    
    mBuffer.SA.pvData = ptr
    pos = 2
    ch = mBuffer.Data(pos)
    Do While ch <> 0
        Select Case ch
            Case DIRECTORY_SEPARATOR_CHAR, ALT_DIRECTORY_SEPARATOR_CHAR
                If foundSeparator Then Exit Do
                foundSeparator = True
        End Select
        pos = pos + 1
        ch = mBuffer.Data(pos)
    Loop
    ret = SysAllocStringLen(ptr, pos)
    MemLong(StrPtr(ret)) = DOUBLE_BACKSLASH     ' replace first 2 characters with '\\'
    GetUNCPathRoot = ret
End Function

Private Sub VerifyUNCPathFormat(ByVal ptr As Long)
    Dim pos As Long
    Dim foundServer As Boolean
    Dim foundShare As Boolean
    Dim foundSeparator As Boolean
    
    Dim ch As Integer
    
    mBuffer.SA.pvData = ptr
    pos = 2
    ch = mBuffer.Data(pos)
    Do While ch <> 0
        Select Case ch
            Case DIRECTORY_SEPARATOR_CHAR, ALT_DIRECTORY_SEPARATOR_CHAR
                foundSeparator = True
            Case Else
                If Not foundServer Then
                    foundServer = True
                ElseIf Not foundShare And foundSeparator Then
                    foundShare = True
                    Exit Do
                End If
        End Select
        pos = pos + 1
        ch = mBuffer.Data(pos)
    Loop
    
    If Not foundServer Or Not foundShare Then _
        Throw Cor.NewArgumentException("UNC path must be in the format of \\server\share.")
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mInvalidPathChars = cArray.NewArray(ciInteger, 34, 60, 62, 124, 0, 8, 16, 17, 18, 20, 21, 22, 23, 24, 25)
    InitWordBuffer mBuffer, 0, &H7FFFFFFF
End Sub
