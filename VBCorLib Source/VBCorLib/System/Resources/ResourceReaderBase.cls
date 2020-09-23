VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResourceReaderBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
'    Module: ResourceReaderBase
'

Option Explicit

Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (ByRef lpPictDesc As PictDesc, ByRef riid As VBGUID, ByVal fOwn As Long, ByRef lplpvObj As Any) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hDC As Long, ByRef pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateIconFromResource Lib "user32.dll" (ByRef presbits As Byte, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long) As Long

Private Type BITMAPINFOHEADER
     biSize As Long          'Size of this header
     biWidth As Long         'width of your image
     biHeight As Long        'height of your image
     biPlanes As Integer     'always 1
     byBitCount As Integer   'number of bits per pixel 1, 4, 8, or 24
     biCompression As Long   '0 data is not compressed
     biSizeImage As Long     'size of bitmap in bytes, typicaly 0 when uncompressed
     biXPelsPerMeter As Long 'preferred resolution in pixels per meter
     biYPelsPerMeter As Long 'preferred resolution in pixels per meter
     biClrUsed As Long       'number of colors that are actually used (can be 0)
     biClrImportant As Long  'which color is most important (0 means all of them)
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(255) As RGBQUAD
End Type

Private Type PictDesc
    cbSizeofStruct As Long
    picType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type

Private mResourceSets As Hashtable



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetResources() As Hashtable
    Set GetResources = mResourceSets
End Function

Friend Sub AddResource(ByVal ResType As ResourceTypes, ByVal ResTypeName As String, ByVal ResOrdinal As Long, ByVal ResName As String, ByVal LangID As Long, ByRef Data() As Byte)
    Select Case ResType
        Case ResourceTypes.StringResource
            AddStringTable LangID, ResOrdinal, Data

        Case ResourceTypes.BitmapResource
            AddBitmap LangID, ResOrdinal, Data
        
        Case ResourceTypes.CursorResource
            AddCursor LangID, ResOrdinal, Data
            
        Case ResourceTypes.IconResource
            AddIcon LangID, ResOrdinal, Data
            
        Case Else
            AddData LangID, ResType, ResTypeName, ResOrdinal, ResName, Data
    End Select
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddData(ByVal LangID As Long, ByVal ResType As ResourceTypes, ByVal ResTypeName As String, ByVal Ordinal As Long, ByVal ResName As String, ByRef Data() As Byte)
    Dim res As Hashtable
    
    Set res = GetResourceSet(LangID, ResType)
    Set res(Ordinal) = Cor.NewWin32Resource(LangID, ResType, ResTypeName, Ordinal, ResName, Data)
End Sub

Private Function PictureFromHandle(ByVal hPic As Long, ByVal picType As Long) As IPictureDisp
    Dim IUnk As VBGUID
    Dim h As Long
    Dim Pic As PictDesc
    Dim lpUnk As IUnknown

    With Pic
        .cbSizeofStruct = LenB(Pic)
        .picType = picType
        .hImage = hPic
    End With
    With IUnk
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    h = OleCreatePictureIndirect(Pic, IUnk, BOOL_TRUE, lpUnk)
    If h = 0 Then Set PictureFromHandle = lpUnk
End Function

Private Sub AddCursor(ByVal LangID As Long, ByVal Ordinal As Long, ByRef Data() As Byte)
    Dim h As Long
    Dim p As IPictureDisp
    Dim res As Hashtable
    
    h = CreateIconFromResource(Data(0), cArray.GetLength(Data), False, &H30000)
    If h <> 0 Then
        Set p = PictureFromHandle(h, vbPicTypeIcon)
        Set res = GetResourceSet(LangID, CursorResource)
        Set res(Ordinal) = Cor.NewWin32Resource(LangID, CursorResource, "", Ordinal, "", p)
    End If
End Sub

Private Sub AddIcon(ByVal LangID As Long, ByVal Ordinal As Long, ByRef Data() As Byte)
    Dim h As Long
    Dim p As IPictureDisp
    Dim res As Hashtable
    
    h = CreateIconFromResource(Data(0), cArray.GetLength(Data), True, &H30000)
    If h <> 0 Then
        Set p = PictureFromHandle(h, vbPicTypeIcon)
        Set res = GetResourceSet(LangID, IconResource)
        Set res(Ordinal) = Cor.NewWin32Resource(LangID, IconResource, "", Ordinal, "", p)
    End If
End Sub

Private Sub AddBitmap(ByVal LangID As Long, ByVal Ordinal As Long, ByRef Data() As Byte)
    Dim bm As BITMAPINFO
    Dim d As Long
    Dim h As Long
    Dim res As Hashtable
    Dim p As IPictureDisp
    
    With bm.bmiHeader
        .biWidth = BitConverter.ToLong(Data, 4)
        .biHeight = BitConverter.ToLong(Data, 8)
        .biPlanes = BitConverter.ToInteger(Data, 12)
        .byBitCount = BitConverter.ToInteger(Data, 14)
        .biSize = LenB(bm.bmiHeader)
        
        If .byBitCount < 24 Then
            .biClrUsed = Helper.ShiftLeft(1, .byBitCount)
            CopyMemory bm.bmiColors(0), Data(40), .biClrUsed * 4
        End If
        
        h = CreateDIBSection(0, bm, 0, d, 0, 0)
        If h <> 0 Then
            CopyMemory ByVal d, Data(40 + .biClrUsed * 4), UBound(Data) - (40 + .biClrUsed * 4)
            Set p = PictureFromHandle(h, vbPicTypeBitmap)
        End If
    End With
    
    Set res = GetResourceSet(LangID, BitmapResource)
    Set res(Ordinal) = Cor.NewWin32Resource(LangID, BitmapResource, "", Ordinal, "", p)
End Sub

Private Sub AddStringTable(ByVal LangID As Long, ByVal Group As Long, ByRef Data() As Byte)
    Dim ResStrings As Hashtable
    Dim i As Long
    Dim j As Long
    Dim l As Long
    Dim s As String
    
    Set ResStrings = GetResourceSet(LangID, StringResource)
    
    For i = 0 To 15
        l = BitConverter.ToInteger(Data, j)
        j = j + 2
        If l > 0 Then
            s = SysAllocStringLen(0, l)
            CopyMemory ByVal StrPtr(s), Data(j), l * 2
            Set ResStrings((Group - 1) * 16 + i) = Cor.NewWin32Resource(LangID, StringResource, "", (Group - 1) * 16 + i, "", s)
            j = j + l * 2
        End If
    Next i
End Sub

Private Function GetResourceSet(ByVal LangID As Long, ByVal ResType As ResourceTypes) As Hashtable
    Dim ResSet As Hashtable
    
    If mResourceSets.Contains(LangID) Then
        Set ResSet = mResourceSets(LangID)
    Else
        Set ResSet = New Hashtable
        mResourceSets.Add LangID, ResSet
    End If
    
    If ResSet.Contains(ResType) Then
        Set GetResourceSet = ResSet(ResType)
    Else
        Set GetResourceSet = New Hashtable
        ResSet.Add ResType, GetResourceSet
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mResourceSets = New Hashtable
End Sub



