VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StreamWriter"
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
'    Module: StreamWriter
'

''
' Writes characters to a byte stream using the specific encoding.
'
' @remarks The default encoding used is UTF8Encoding. If the writer is
' not appending to the stream, then an encoding preamble is written to
' the stream.
'
' @see Constructors
' @see StreamReader
' @see TextWriter
'
Option Explicit
Implements cObject
Implements TextWriter

Private Const MIN_BUFFERSIZE As Long = 128


Private mStream As Stream
Private mAutoFlush As Boolean
Private mEncoding As Encoding
Private mEncoder As Encoder
Private mNewLine As String
Private mCharBuffer() As Integer
Private mByteBuffer() As Byte
Private mBufferSize As Long
Private mPosition As Long
Private mIsOpen As Boolean



''
' Returns if the writer autoflushes after each write.
'
' @return Indication of autoflush.
'
Public Property Get AutoFlush() As Boolean
    AutoFlush = mAutoFlush
End Property

''
' Sets if the writer autoflushes after each write.
'
' @param RHS The new value.
'
Public Property Let AutoFlush(ByVal RHS As Boolean)
    mAutoFlush = RHS
End Property

''
' Returns the stream this writer is using.
'
' @return The stream being used by this writer.
'
Public Property Get BaseStream() As Stream
    Set BaseStream = mStream
End Property

''
' Returns the encoding being used by this writer.
'
' @return The encoding being used by this writer.
'
Public Property Get Encoding() As Encoding
    Set Encoding = mEncoding
End Property

''
' Returns the format provider being used by this writer.
'
' @return The format provider used by this writer.
'
Public Property Get FormatProvider() As IFormatProvider
    Set FormatProvider = CultureInfo.CurrentCulture
End Property

''
' Returns the new-line string used by this writer.
'
' @return The new-line string.
'
Public Property Get NewLine() As String
    NewLine = mNewLine
End Property

''
' Sets the new-line string used by this writer.
'
' @param RHS The new new-line value.
'
Public Property Let NewLine(ByVal RHS As String)
    mNewLine = RHS
End Property

''
' Closes this writer, flushing any buffers needed.
'
Public Sub CloseWriter()
    If mIsOpen Then
        Flush
        mIsOpen = False
        mStream.CloseStream
    End If
End Sub

''
' Flushes the buffers to the underlying stream.
'
Public Sub Flush()
    VerifyIsOpen
    InternalFlush True, True
End Sub

''
' Writes the encoded string representation of the value to the stream.
'
' @param value The value to be written to the stream.
' @param args The arguments to be used in formatting the Value.
' @remarks The <i>args</i> are used when argument indexes are supplied
' in the value to be written. Indexes are indicate by enclosing the
' specific index in curly braces {}. The indexes start at 0.
'
Public Sub WriteValue(ByRef Value As Variant, ParamArray args() As Variant)
    Dim a() As Variant
    Helper.Swap4 ByVal ArrPtr(a), ByVal Helper.DerefEBP(16)
    InternalWriteValue Value, a
End Sub

''
' Writes the value, then appends the new-line string.
'
' @param value The value to be written.
' @remarks The first argument in <i>value</i> is the output to the
' string. Any additional arguments are used as formatting information.
' An argument can be refered to through formatting indexes. A
' formatting index is an integer enclosed in curly braces {}. The
' indexes start at 0.
'
Public Sub WriteLine(ParamArray Value() As Variant)
    Dim a() As Variant
    Helper.Swap4 ByVal ArrPtr(a), ByVal Helper.DerefEBP(12)
    InternalWriteLine a
End Sub

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
Friend Sub Init(ByVal Source As Variant, ByVal Encoding As Encoding, ByVal BufferSize As Long, ByVal Append As Boolean)
    Dim Mode As FileMode
    Dim emit As Boolean
    
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
            If Append Then
                Mode = FileMode.Append
            Else
                Mode = FileMode.Create
                emit = True
            End If
            Set mStream = Cor.NewFileStream(Source, Mode, FileAccess.WriteAccess, FileShare.ReadShare)
        Case Else
            Throw Cor.NewArgumentException("A file path or Stream object is required.", "Source")
    End Select
    
    If Encoding Is Nothing Then Set Encoding = Cor.NewUTF8Encoding(False, True)
    Set mEncoding = Encoding
    Set mEncoder = mEncoding.GetEncoder
    
    mNewLine = Environment.NewLine
    mIsOpen = True
    If BufferSize < MIN_BUFFERSIZE Then BufferSize = MIN_BUFFERSIZE
    ReDim mCharBuffer(0 To BufferSize - 1)
    ReDim mByteBuffer(0 To mEncoding.GetMaxByteCount(BufferSize) - 1)
    mBufferSize = BufferSize
    If emit Then WritePreamble
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InternalWrite(ByVal Value As String, ByRef args() As Variant, ByVal NewLine As Boolean)
    Dim Size As Long
    Dim ptr As Long
    Dim n As Long
    
    VerifyIsOpen
    
    If UBound(args) >= 0 Then Value = cString.FormatArray(Value, args)
    If NewLine Then Value = Value & mNewLine
    
    Size = Len(Value)
    ptr = StrPtr(Value)
    
    Do While Size > 0
        If mPosition + Size >= mBufferSize Then InternalFlush False, False
        If Size > mBufferSize Then
            n = mBufferSize
        Else
            n = Size
        End If
        CopyMemory mCharBuffer(mPosition), ByVal ptr, n * 2
        Size = Size - n
        ptr = ptr + n * 2
        mPosition = mPosition + n
    Loop
    If mAutoFlush Then InternalFlush True, False
End Sub

Private Sub InternalWriteLine(ByRef args() As Variant)
    Dim ub As Long
    Dim s As String
    
    ub = UBound(args)
    
    If ub >= 0 Then s = Convert.ToString(args(0))
    If ub > 0 Then
        Dim i As Long
        For i = 1 To ub
            Helper.MoveVariant args(i - 1), args(i)
        Next i
        ReDim Preserve args(ub - 1)
    End If
    
    InternalWrite s, args, True
End Sub

Private Sub InternalWriteValue(ByRef Value As Variant, ByRef args() As Variant)
    InternalWrite Convert.ToString(Value), args, False
End Sub

Private Sub VerifyIsOpen()
    If Not mIsOpen Then Throw Cor.NewObjectDisposedException("StreamWriter", "Cannot write to a close writer.")
End Sub

Private Sub WritePreamble()
    Dim b() As Byte
    
    If Not mStream.CanSeek Then Exit Sub
    If mStream.Length > 0 Then Exit Sub
    
    b = mEncoding.GetPreamble
    mStream.WriteBlock b, 0, cArray.GetLength(b)
End Sub

Private Sub InternalFlush(ByVal FlushStream As Boolean, ByVal FlushEncoder As Boolean)
    Dim Size As Long
    
    If mPosition > 0 Then
        Size = mEncoder.GetBytes(mCharBuffer, 0, mPosition, mByteBuffer, 0, FlushEncoder)
        mStream.WriteBlock mByteBuffer, 0, Size
        mPosition = 0
    End If
    
    If FlushStream Then mStream.Flush
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
    CloseWriter
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
'   TextWriter Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TextWriter_CloseWriter()
    CloseWriter
End Sub

Private Property Get TextWriter_Encoding() As Encoding
    Set TextWriter_Encoding = Encoding
End Property

Private Function TextWriter_Equals(Value As Variant) As Boolean
    TextWriter_Equals = Equals(Value)
End Function

Private Sub TextWriter_Flush()
    Flush
End Sub

Private Property Get TextWriter_FormatProvider() As IFormatProvider
    Set TextWriter_FormatProvider = FormatProvider
End Property

Private Function TextWriter_GetHashCode() As Long
    TextWriter_GetHashCode = GetHashCode
End Function

Private Property Let TextWriter_NewLine(ByVal RHS As String)
    NewLine = RHS
End Property

Private Property Get TextWriter_NewLine() As String
    TextWriter_NewLine = NewLine
End Property

Private Function TextWriter_ToString() As String
    TextWriter_ToString = ToString
End Function

Private Sub TextWriter_WriteLine(ParamArray Value() As Variant)
    Dim a() As Variant
    Helper.Swap4 ByVal ArrPtr(a), ByVal Helper.DerefEBP(12)
    InternalWriteLine a
End Sub

Private Sub TextWriter_WriteValue(ByRef Value As Variant, ParamArray args() As Variant)
    Dim a() As Variant
    Helper.Swap4 ByVal ArrPtr(a), ByVal Helper.DerefEBP(16)
    InternalWriteValue Value, a
End Sub

