VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StreamReader"
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
'    Module: StreamReader
'

''
' Reads characters from a byte array in a particular encoding.
'
' @remarks The <b>StreamReader</b> object cannot be created directly. In order to
' create a new <b>StreamReader</b>, use the <b>Cor.NewStreamReader</b> method.
' <p>Unlike the FileStream and MemoryStream classes, the StreamReader
' reads bytes that are of a particular encoding to be decoded. Instead of reading
' bytes, the StreamReader returns characters or strings of the decoded bytes.</p>
' <p>The default encoding is UTF8Encoding.</p>
'
' @see Constructors
' @see StreamWriter
' @see TextReader
' @include "..\..\Includes\StreamReader.txt"
Option Explicit
Implements cObject
Implements TextReader

Private Const MIN_BUFFERSIZE As Long = 128


Private mStream As Stream
Private mEncoding As Encoding
Private mDecoder As Decoder
Private mByteBuffer() As Byte
Private mByteIndex As Long
Private mByteLength As Long
Private mCharBuffer() As Integer
Private mCharIndex As Long
Private mCharLength As Long
Private mBufferSize As Long
Private mDetermineEncoding As Boolean
Private mIsOpen As Boolean
Private mSB As StringBuilder


''
' Returns the base stream this reader is reading from.
'
' @return The base stream.
'
Public Property Get BaseStream() As Stream
    Set BaseStream = mStream
End Property

''
' Returns the current encoding used by the reader.
'
' @return The encoding being used by the reader currently.
'
Public Property Get CurrentEncoding() As Encoding
    Set CurrentEncoding = mEncoding
End Property

''
' Closes the reader.
'
Public Sub CloseReader()
    If mIsOpen Then
        mIsOpen = False
        mStream.CloseStream
    End If
End Sub

''
' Discards the currently buffered data to allow reading from
' a new position in the underlying stream.
'
' @remarks The StreamReader buffers more data than is usually
' read at a single time. If the position of the underlying
' stream is moved, then the data buffered in the StreamReader
' needs to be discarded so the reading can begin at the new
' Stream position.
'
Public Sub DiscardBufferedData()
    mByteIndex = 0
    mByteLength = 0
    mCharIndex = 0
    mCharLength = 0
    Set mDecoder = mEncoding.GetDecoder
End Sub

''
' Returns the the next char from the reader without consuming it.
'
' @return The next character to be read, or -1 if the end of the
' stream has been reached.
' @remarks The character that is returned is not removed from the stream.
' Any characters above &H7FFF will be returned as a positive value.
'
Public Function Peek() As Long
    VerifyIsOpen
    
    If mCharIndex = mCharLength Then
        If ReadBuffer = 0 Then
            Peek = -1
            Exit Function
        End If
    End If
    AsWord(Peek) = mCharBuffer(mCharIndex)
End Function

''
' Returns the next char from the reader.
'
' @return The next character to be read, or -1 if the end of the stream was reached.
' @remarks Any characters above &H7FFF will be returned as a positive value.
'
Public Function Read() As Long
    VerifyIsOpen
    
    If mCharIndex = mCharLength Then
        If ReadBuffer = 0 Then
            Read = -1
            Exit Function
        End If
    End If
    AsWord(Read) = mCharBuffer(mCharIndex)
    mCharIndex = mCharIndex + 1
End Function

''
' Reads a block of characters from the the stream.
'
' @param Buffer The array to read the characters in to.
' @param Index The starting index in <i>Buffer</i> to begin reading to.
' @param Count The number of characters to be read.
' @return The actual number of characters read.
'
Public Function ReadBlock(ByRef Buffer() As Integer, ByVal Index As Long, ByVal Count As Long) As Long
    If cArray.IsNull(Buffer) Then _
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Array), "Buffer")
    If Index < LBound(Buffer) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "Index", Index)
    If Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "Count", Count)
    If Index + Count - 1 > UBound(Buffer) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "Count")
    
    VerifyIsOpen
    
    Dim n As Long
    Do While Count > 0
        If mCharIndex = mCharLength Then
            If ReadBuffer = 0 Then Exit Function
        End If
        n = mCharLength - mCharIndex
        If n > Count Then n = Count
        CopyMemory Buffer(Index), mCharBuffer(mCharIndex), n * 2
        Index = Index + n
        mCharIndex = mCharIndex + n
        Count = Count - n
        ReadBlock = ReadBlock + n
    Loop
End Function

''
' Returns a string of characters to the next new-line character.
'
' @return A string containing the characters up to the next new-line character.
' @remarks If there are no more bytes to in the stream, then a null string is returned.
' This can be checked for using the <b>cString.IsNull</b> function.
'
Public Function ReadLine() As String
    Dim s As String
    Dim StartIndex As Long
    Dim TermChar As Integer
    Dim useSB As Boolean
    
    
    VerifyIsOpen
    
    If mCharIndex = mCharLength Then
        If ReadBuffer = 0 Then Exit Function
    End If
    
    StartIndex = mCharIndex
    Do
        If mCharIndex = mCharLength Then
            If Not useSB Then
                If mSB Is Nothing Then Set mSB = New StringBuilder
                mSB.Length = 0
                useSB = True
            End If
            mSB.AppendChars mCharBuffer, StartIndex, mCharIndex - StartIndex
            StartIndex = 0
            If ReadBuffer = 0 Then Exit Do
        End If
        
        Select Case mCharBuffer(mCharIndex)
            Case 13, 10
                TermChar = mCharBuffer(mCharIndex)
                If Not useSB Then
                    s = cString.FromCharArray(mCharBuffer, StartIndex, mCharIndex - StartIndex)
                    mCharIndex = mCharIndex + 1
                    Exit Do
                Else
                    mSB.AppendChars mCharBuffer, StartIndex, mCharIndex - StartIndex
                    mCharIndex = mCharIndex + 1
                    Exit Do
                End If
        End Select
        mCharIndex = mCharIndex + 1
    Loop
    
    If Not useSB Then
        If Len(s) = 0 And mCharLength = 0 Then
            Exit Function
        Else
            ReadLine = s
        End If
    Else
        ReadLine = mSB.ToString
    End If
    
    If TermChar = 13 Then
        If mCharIndex < mCharLength Then
            If mCharBuffer(mCharIndex) = 10 Then mCharIndex = mCharIndex + 1
        Else
            If ReadBuffer > 0 Then
                If mCharBuffer(0) = 10 Then mCharIndex = mCharIndex + 1
            End If
        End If
    End If
End Function

''
' Returns a string of all remaining characters in the stream.
'
' @return A string containing the remaining characters.
'
Public Function ReadToEnd() As String
    VerifyIsOpen
    If mCharIndex = mCharLength Then If ReadBuffer = 0 Then Exit Function
    If mByteLength < mBufferSize Then
        ReadToEnd = cString.FromCharArray(mCharBuffer, mCharIndex, mCharLength)
    Else
        If mSB Is Nothing Then Set mSB = New StringBuilder
        mSB.Length = 0
        Do
            mSB.AppendChars mCharBuffer, mCharIndex, mCharLength - mCharIndex
        Loop While ReadBuffer > 0
        ReadToEnd = mSB.ToString
    End If
    mCharIndex = mCharLength
End Function

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
Friend Sub Init(ByRef Source As Variant, ByVal oEncoding As Encoding, ByVal determineEncodingFromByteOrderMarks As Boolean, ByVal BufferSize As Long)
    If BufferSize < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "BufferSize", BufferSize)
    
    Select Case VarType(Source)
        Case vbObject
            If TypeOf Source Is Stream Then
                Set mStream = Source
            Else
                Throw Cor.NewArgumentException("A Stream object is required.", "Source")
            End If
        Case vbString
            Set mStream = Cor.NewFileStream(Source, FileMode.OpenExisting, FileAccess.ReadAccess, FileShare.ReadShare)
        Case Else
            Throw Cor.NewArgumentException("A file path or Stream object is required.", "Source")
    End Select
    
    mDetermineEncoding = determineEncodingFromByteOrderMarks
    If oEncoding Is Nothing Then
        Set mEncoding = Encoding.UTF8
    Else
        Set mEncoding = oEncoding
    End If
    Set mDecoder = mEncoding.GetDecoder
    If BufferSize < MIN_BUFFERSIZE Then BufferSize = MIN_BUFFERSIZE
    mBufferSize = BufferSize
    mIsOpen = True
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ReadBuffer() As Long
    If cArray.IsNull(mByteBuffer) Then ReDim mByteBuffer(0 To mBufferSize - 1)
    mByteLength = mStream.ReadBlock(mByteBuffer, 0, mBufferSize)
    If mByteLength > 0 Then
        If mDetermineEncoding Then DetermineEncoding
        If cArray.IsNull(mCharBuffer) Then ReDim mCharBuffer(0 To mEncoding.GetMaxCharCount(mBufferSize) - 1)
        mCharLength = mDecoder.GetChars(mByteBuffer, mByteIndex, mByteLength, mCharBuffer, 0)
    Else
        mCharLength = 0
    End If
    mByteIndex = 0
    mCharIndex = 0
    ReadBuffer = mCharLength
End Function

Private Sub DetermineEncoding()
    If mByteLength < 2 Then Exit Sub
    Select Case AsWord(mByteBuffer(0))
        Case &HFEFF   ' Little-Endian UnicodeEncoding
            Set mEncoding = Cor.NewUnicodeEncoding(False, False)
            mByteIndex = 2
            mByteLength = mByteLength - 2
        Case &HFFFE   ' Big-Endian UnicodeEncoding
            Set mEncoding = Cor.NewUnicodeEncoding(True, False)
            mByteIndex = 2
            mByteLength = mByteLength - 2
        Case &HBBEF   ' possible UTF8 encoding
            If mByteLength < 3 Then Exit Sub
            If mByteBuffer(2) = &HBF Then Set mEncoding = Cor.NewUTF8Encoding(False, False)
            mByteIndex = 3
            mByteLength = mByteLength - 3
    End Select
    If Not mEncoding Is Nothing Then Set mDecoder = mEncoding.GetDecoder
    mDetermineEncoding = False
End Sub

Private Sub VerifyIsOpen()
    If Not mIsOpen Then Throw Cor.NewObjectDisposedException("StreamReader", "Cannot read from a closed StreamReader.")
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mEncoding = Encoding.UTF8
    Set mDecoder = mEncoding.GetDecoder
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
'   TextReader Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TextReader_CloseReader()
    CloseReader
End Sub

Private Function TextReader_Equals(Value As Variant) As Boolean
    TextReader_Equals = Equals(Value)
End Function

Private Function TextReader_GetHashCode() As Long
    TextReader_GetHashCode = GetHashCode
End Function

Private Function TextReader_Peek() As Long
    TextReader_Peek = Peek
End Function

Private Function TextReader_Read() As Long
    TextReader_Read = Read
End Function

Private Function TextReader_ReadBlock(Chars() As Integer, ByVal Index As Long, ByVal Count As Long) As Long
    TextReader_ReadBlock = ReadBlock(Chars, Index, Count)
End Function

Private Function TextReader_ReadLine() As String
    TextReader_ReadLine = ReadLine
End Function

Private Function TextReader_ReadToEnd() As String
    TextReader_ReadToEnd = ReadToEnd
End Function

Private Function TextReader_ToString() As String
    TextReader_ToString = ToString
End Function
