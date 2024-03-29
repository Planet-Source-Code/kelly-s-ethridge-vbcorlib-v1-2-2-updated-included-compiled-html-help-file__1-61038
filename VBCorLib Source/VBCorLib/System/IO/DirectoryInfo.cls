VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DirectoryInfo"
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
'    Module: DirectoryInfo
'

''
' Represents a directory to modify and retrieve information about.
'
' @remarks The DirectoryInfo object does not verify the existence
' of the directory until information has to be read or written to that directory.
'
' @see Constructors
'
Option Explicit
Implements cObject
Implements FileSystemInfo

Private mPath As String
Private mOriginalPath As String
Private mIsDirty As Boolean
Private mFileAttributes As Long
Private mCreationTime As cDateTime
Private mLastAccessTime As cDateTime
Private mLastWriteTime As cDateTime


''
' Returns the name of the directory this instance represents.
'
' @param The name of the directory.
'
Public Property Get Name() As String
    Name = Path.GetFileName(mPath)
    If Len(Name) = 0 Then Name = mPath
End Property

''
' Returns the name of the directory, including the full path to the directory.
'
' @param The full path to the directory.
'
Public Property Get FullName() As String
    FullName = mPath
End Property

''
' Returns the directory's attributes (ReadOnly, Hidden, Archive, ...)
'
' @return The directories attributes since the last time Refresh was called.
' @remarks The attributes are not kept current. If the they are changed since
' last checked, then Refresh must be called to retrieve the current settings.
'
Public Property Get Attributes() As FileAttributes
    If mIsDirty Then Refresh
    Attributes = mFileAttributes
End Property

''
' Sets the attributes for the directory.
'
' @param RHS The attributes to set for the directory.
' @remarks Setting just one attribute will clear all the other attributes. To
' preserve any attributes, then OR must be used with the attributes to
' be preserved.
'
Public Property Let Attributes(ByVal RHS As FileAttributes)
    If SetFileAttributes(mPath, RHS) = BOOL_FALSE Then IOError Err.LastDllError, mPath
    MarkDirty
End Property

''
' Returns the creation time for this directory in local time.
'
' @return A cDateTime object representing the creation time for this directory.
' @remarks The time returned is the UTC time with the timezone offset applied.
' For Pacific Coast, a -8 hours is added to the time before returning.
'
Public Property Get CreationTime() As Variant
    If mIsDirty Then Refresh
    Set CreationTime = mCreationTime
End Property

''
' Sets the creation time for the directory in local time.
'
' @param RHS A Date of the new time.
' @remarks The time has the UTC timezone offset applied to it before
' being written to the directory.
'
Public Property Let CreationTime(ByVal RHS As Variant)
    Set Me.CreationTime = cDateTime.GetcDateTime(RHS)
End Property

''
' Sets the creation time for the directory in local time.
'
' @param RHS A cDateTime object of the new time.
' @remarks The time has the UTC timezone offset applied to it before
' being written to the directory.
'
Public Property Set CreationTime(ByVal RHS As Variant)
    If Not TypeOf RHS Is cDateTime Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_DateRequired), "CreationTime")

    Directory.SetCreationTime mPath, mCreationTime
    MarkDirty
End Property

''
' Returns the creation time as UTC time.
'
' @return A cDateTime object of the UTC time.
' @remarks The time returned is the time on the directory without
' any offsets being applied.
'
Public Property Get CreationTimeUtc() As Variant
    If mIsDirty Then Refresh
    Set CreationTimeUtc = mCreationTime.ToUniversalTime
End Property

''
' Sets the creation time as UTC time.
'
' @param RHS A Date value as the new time.
' @remarks The time has the UTC timezone offset removed from it before
' being written to the directory.
'
Public Property Let CreationTimeUtc(ByVal RHS As Variant)
    Set Me.CreationTimeUtc = cDateTime.GetcDateTime(RHS)
End Property

''
' Sets the creation time as UTC time.
'
' @param RHS A cDateTime object of the new time.
' @remarks The time has the UTC timezone offset removed from it before
' being written to the directory.
'
Public Property Set CreationTimeUtc(ByVal RHS As Variant)
    If Not TypeOf RHS Is cDateTime Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_DateRequired), "CreationTime")

    Directory.SetCreationTimeUtc mPath, RHS
    MarkDirty
End Property

''
' Returns the last time the directory was accessed in local time.
'
' @return A cDateTime object of the time.
' @remarks The returned time is the UTC time with the UTC timezone offset applied.
'
Public Property Get LastAccessTime() As Variant
    If mIsDirty Then Refresh
    Set LastAccessTime = mLastAccessTime
End Property

''
' Sets the time the directory was last accessed in local time.
'
' @param RHS A Date or cDateTime object of the new time.
'
Public Property Let LastAccessTime(ByVal RHS As Variant)
    Set Me.LastAccessTime = cDateTime.GetcDateTime(RHS)
End Property

''
' Sets the time the directory was last accessed in local time.
'
' @param RHS A Date or cDateTime object of the new time.
'
Public Property Set LastAccessTime(ByVal RHS As Variant)
    If Not TypeOf RHS Is cDateTime Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_DateRequired), "LastAccessTime")
        
    Directory.SetLastAccessTime mPath, mLastAccessTime
    MarkDirty
End Property

''
' Returns the last time the directory was accessed in UTC time.
'
' @return a cDateTime object of the time.
'
Public Property Get LastAccessTimeUtc() As Variant
    If mIsDirty Then Refresh
    Set LastAccessTimeUtc = mLastAccessTime.ToUniversalTime
End Property

''
' Sets the last time the directory was accessed in UTC time.
'
' @param RHS A Date or cDateTime object of the new time.
'
Public Property Let LastAccessTimeUtc(ByVal RHS As Variant)
    Set Me.LastAccessTimeUtc = cDateTime.GetcDateTime(RHS)
End Property

''
' Sets the last time the directory was accessed in UTC time.
'
' @param RHS A Date or cDateTime object of the new time.
'
Public Property Set LastAccessTimeUtc(ByVal RHS As Variant)
    If Not TypeOf RHS Is cDateTime Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_DateRequired), "LastAccessTime")
    
    Directory.SetLastAccessTimeUtc mPath, RHS
    MarkDirty
End Property

''
' Returns the last time the directory was written to in local time.
'
' @return A cDateTime object of the time.
'
Public Property Get LastWriteTime() As Variant
    If mIsDirty Then Refresh
    Set LastWriteTime = mLastWriteTime
End Property

''
' Sets the last time the directory was written to in local time.
'
' @param RHS A Date or cDateTime object of the new time.
'
Public Property Let LastWriteTime(ByVal RHS As Variant)
    Set Me.LastAccessTime = cDateTime.GetcDateTime(RHS)
End Property

''
' Sets the last time the directory was written to in local time.
'
' @param RHS A Date or cDateTime object of the new time.
'
Public Property Set LastWriteTime(ByVal RHS As Variant)
    If Not TypeOf RHS Is cDateTime Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_DateRequired), "LastWriteTime")
    
    Directory.SetLastWriteTime mPath, mLastWriteTime
    MarkDirty
End Property

''
' Returns the last time the directory was written to in UTC time.
'
' @return A cDateTime object of the time.
'
Public Property Get LastWriteTimeUtc() As Variant
    If mIsDirty Then Refresh
    Set LastWriteTimeUtc = mLastWriteTime.ToUniversalTime
End Property

''
' Sets the last time the directory was written to in UTC time.
'
' @param RHS A Date or cDateTime object of the new time.
'
Public Property Let LastWriteTimeUtc(ByVal RHS As Variant)
    Set Me.LastWriteTimeUtc = cDateTime.GetcDateTime(RHS)
End Property

''
' Sets the last time the directory was written to in UTC time.
'
' @param RHS A Date or cDateTime object of the new time.
'
Public Property Set LastWriteTimeUtc(ByVal RHS As Variant)
    If Not TypeOf RHS Is cDateTime Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_DateRequired), "LastAccessTime")
    
    Directory.SetLastWriteTimeUtc mPath, RHS
    MarkDirty
End Property

''
' Reloads all of the attributes and timestamps for the directory.
'
Public Sub Refresh()
    Dim Data As WIN32_FILE_ATTRIBUTE_DATA
    Dim E As Long
    
    E = file.GetFileData(mPath, Data)
    If E <> NO_ERROR Then IOError E, mPath
    
    With Data
        mFileAttributes = .dwFileAttributes
        Set mCreationTime = cDateTime.FromFileTime(.ftCreationTime)
        Set mLastAccessTime = cDateTime.FromFileTime(.ftLastAccessTime)
        Set mLastWriteTime = cDateTime.FromFileTime(.ftLastWriteTime)
    End With
    
    mIsDirty = False
End Sub

''
' Checks if the directory actually exists.
'
' @return An indication of the directory existing.
'
Public Property Get Exists() As Boolean
    Exists = Directory.Exists(mPath)
End Property

''
' Returns an empty string for directories.
'
' @return An empty string.
'
Public Property Get Extension() As String
' do nothing
End Property

''
' Returns the parent directory as a DirectoryInfo object.
'
' @return A DirectroyInfo object of the parent directory, or Nothing if
' the original directory is a root, such as 'c:\'
Public Property Get Parent() As DirectoryInfo
    Dim p As String
    
    p = Path.GetDirectoryName(mPath)
    If Not cString.IsNull(p) Then Set Parent = Cor.NewDirectoryInfo(p)
End Property

''
' Returns a DirectoryInfo object of the root directory of this instance.
'
' @return A DirectoryInfo object of the root.
'
Public Property Get Root() As DirectoryInfo
    Set Root = Cor.NewDirectoryInfo(Path.GetPathRoot(mPath))
End Property

''
' Creates the path represented by this instance.
'
Public Sub Create()
    Directory.CreateDirectory mPath
End Sub

''
' Creates a directory within this instance, returning the new subdirectory.
'
' @param sPath The new subdirectory to be created.
' @return A DirectoryInfo object of the new subdirectory path.
'
Public Function CreateSubdirectory(ByVal sPath As String) As DirectoryInfo
    Dim fp As String
    
    fp = Path.Combine(mPath, sPath)
    Set CreateSubdirectory = Directory.CreateDirectory(fp)
End Function

''
' Deletes this directory if it is empty, or all of its contents and itself
' if Recursive is set to True.
'
' @param Recursive Indicates if all entries in the directory should be deleted.
'
Public Sub Delete(Optional ByVal Recursive As Boolean)
    Directory.Delete mPath, Recursive
End Sub

''
' Returns a list of directories within this directory.
'
' @param SearchPattern A pattern for all directories to match.
' @return A list of all matching directories.
' @remarks If no matching directories were found, then the root is returned.
'
Public Function GetDirectories(Optional ByVal SearchPattern As String = "*") As DirectoryInfo()
    Dim ret() As DirectoryInfo
    Dim s() As String
    Dim i As Long
    Dim l As Long
    
    s = Directory.GetDirectories(mPath, SearchPattern)
    l = cArray.GetLength(s)
    If l > 0 Then
        For i = 0 To l - 1
            s(i) = Path.Combine(mPath, s(i))
        Next i
        
        ReDim ret(0 To l - 1)
        For i = 0 To l - 1
            Set ret(i) = Cor.NewDirectoryInfo(s(i))
        Next i
    Else
        ReDim ret(0)
        Set ret(0) = Root
    End If
    GetDirectories = ret
End Function

''
' Moves this directory and its contents to the specified path.
'
' @param DestDirName The new path and name of the directory.
'
Public Sub MoveTo(ByVal DestDirName As String)
    Directory.Move mPath, DestDirName
    mPath = Path.GetFullPath(DestDirName)
    mOriginalPath = Path.GetDirectoryName(mPath)
End Sub

''
' Returns a list of file matching the pattern in this directory.
'
' @param SearchPattern The pattern the files must match.
' @return A list of files that matched the pattern.
' @remarks If no files matched the pattern, a zero-length array is returned.
'
Public Function GetFiles(Optional ByVal SearchPattern As String = "*") As FileInfo()
    Dim f() As String
    Dim l As Long
    Dim ret() As FileInfo
    Dim i As Long
    
    f = Directory.GetFiles(mPath, SearchPattern)
    l = cArray.GetLength(f)
    If l > 0 Then
        For i = 0 To l - 1
            f(i) = Path.Combine(mPath, f(i))
        Next i
        
        ReDim ret(0 To l - 1)
        For i = 0 To l - 1
            Set ret(i) = Cor.NewFileInfo(f(i))
        Next i
    Else
        ret = cArray.CreateInstance(ciObject)
    End If
    GetFiles = ret
End Function

''
' Returns a list of all entries in the directory that match the pattern.
'
' @param SearchPattern The pattern to match against all entries.
' @return A list of all entries that matched the pattern.
' @remarks If no entries matched the pattern, a zero-length array is returned.
'
Public Function GetFileSystemInfos(Optional ByVal SearchPattern As String = "*") As FileSystemInfo()
    Dim Files() As String
    Dim Folders() As String
    Dim ret() As FileSystemInfo
    Dim i As Long
    
    Files = Directory.GetFiles(mPath, SearchPattern)
    Folders = Directory.GetDirectories(mPath, SearchPattern)
    
    ReDim ret(0 To cArray.GetLength(Files) + cArray.GetLength(Folders) - 1)
    For i = 0 To UBound(Files)
        Set ret(i) = Cor.NewFileInfo(Path.Combine(mPath, Files(i)))
    Next i
    
    Dim j As Long
    j = UBound(Files) + 1
    For i = 0 To UBound(Folders)
        Set ret(j + i) = Cor.NewDirectoryInfo(Path.Combine(mPath, Folders(i)))
    Next i
    
    GetFileSystemInfos = ret

'    Dim f() As String
'    Dim l As Long
'    Dim ret() As FileSystemInfo
'    Dim i As Long
'
'    f = Directory.GetFileSystemEntries(mPath, SearchPattern)
'    l = cArray.GetLength(f)
'    If l > 0 Then
'        For i = 0 To l - 1
'            f(i) = Path.Combine(mPath, f(i))
'        Next i
'
'        ReDim ret(0 To l - 1)
'        For i = 0 To l - 1
'            If GetFileAttributes(f(i)) And vbDirectory Then
'                Set ret(i) = Cor.NewDirectoryInfo(f(i))
'            Else
'                Set ret(i) = Cor.NewFileInfo(f(i))
'            End If
'        Next i
'    Else
'        ret = cArray.CreateInstance(ciObject)
'    End If
'    GetFileSystemInfos = ret
End Function

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Public Function ToString() As String
    ToString = mOriginalPath
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
Friend Sub Init(ByVal sPath As String)
    Path.VerifyPath sPath
    mOriginalPath = sPath
    If Not Path.IsPathRooted(sPath) Then
        mPath = Path.GetFullPath(sPath)
    Else
        mPath = sPath
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub MarkDirty()
    mIsDirty = True
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    MarkDirty
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
'   FileSystemInfo Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get FileSystemInfo_Attributes() As FileAttributes
    FileSystemInfo_Attributes = Attributes
End Property

Private Property Let FileSystemInfo_Attriubtes(ByVal RHS As FileAttributes)
    Attributes = RHS
End Property

Private Property Set FileSystemInfo_CreationTime(ByVal RHS As Variant)
    Set CreationTime = RHS
End Property

Private Property Let FileSystemInfo_CreationTime(ByVal RHS As Variant)
    CreationTime = RHS
End Property

Private Property Get FileSystemInfo_CreationTime() As Variant
    Set FileSystemInfo_CreationTime = CreationTime
End Property

Private Property Set FileSystemInfo_CreationTimeUtc(ByVal RHS As Variant)
    Set CreationTimeUtc = RHS
End Property

Private Property Let FileSystemInfo_CreationTimeUtc(ByVal RHS As Variant)
    CreationTimeUtc = RHS
End Property

Private Property Get FileSystemInfo_CreationTimeUtc() As Variant
    Set FileSystemInfo_CreationTimeUtc = CreationTimeUtc
End Property

Private Sub FileSystemInfo_Delete()
    Delete
End Sub

Private Function FileSystemInfo_Equals(Value As Variant) As Boolean
    FileSystemInfo_Equals = Equals(Value)
End Function

Private Property Get FileSystemInfo_Exists() As Boolean
    FileSystemInfo_Exists = Exists
End Property

Private Property Get FileSystemInfo_Extension() As String
    FileSystemInfo_Extension = Extension
End Property

Private Property Get FileSystemInfo_FullName() As String
    FileSystemInfo_FullName = FullName
End Property

Private Function FileSystemInfo_GetHashCode() As Long
    FileSystemInfo_GetHashCode = GetHashCode
End Function

Private Property Set FileSystemInfo_LastAccessTime(ByVal RHS As Variant)
    Set LastAccessTime = RHS
End Property

Private Property Get FileSystemInfo_LastAccessTime() As Variant
    Set FileSystemInfo_LastAccessTime = LastAccessTime
End Property

Private Property Let FileSystemInfo_LastAccessTime(ByVal RHS As Variant)
    LastAccessTime = RHS
End Property

Private Property Set FileSystemInfo_LastAccessTimeUtc(ByVal RHS As Variant)
    Set LastAccessTimeUtc = RHS
End Property

Private Property Let FileSystemInfo_LastAccessTimeUtc(ByVal RHS As Variant)
    LastAccessTimeUtc = RHS
End Property

Private Property Get FileSystemInfo_LastAccessTimeUtc() As Variant
    Set FileSystemInfo_LastAccessTimeUtc = LastAccessTimeUtc
End Property

Private Property Set FileSystemInfo_LastWriteTime(ByVal RHS As Variant)
    Set LastWriteTime = RHS
End Property

Private Property Let FileSystemInfo_LastWriteTime(ByVal RHS As Variant)
    LastWriteTime = RHS
End Property

Private Property Get FileSystemInfo_LastWriteTime() As Variant
    Set FileSystemInfo_LastWriteTime = LastWriteTime
End Property

Private Property Set FileSystemInfo_LastWriteTimeUtc(ByVal RHS As Variant)
    Set LastWriteTimeUtc = RHS
End Property

Private Property Let FileSystemInfo_LastWriteTimeUtc(ByVal RHS As Variant)
    LastWriteTimeUtc = RHS
End Property

Private Property Get FileSystemInfo_LastWriteTimeUtc() As Variant
    Set FileSystemInfo_LastWriteTimeUtc = LastWriteTimeUtc
End Property

Private Property Get FileSystemInfo_Name() As String
    FileSystemInfo_Name = Name
End Property

Private Sub FileSystemInfo_Refresh()
    Refresh
End Sub

Private Function FileSystemInfo_ToString() As String
    FileSystemInfo_ToString = ToString
End Function

