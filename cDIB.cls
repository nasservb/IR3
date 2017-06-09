VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const BMP_MAGIC_COOKIE As Integer = 19778
Private Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type
Private Type BITMAPINFOHEADER
   biSize As Long
   biWidth As Long
   biHeight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type
Private Type RGBQUAD
    Red As Byte
    Green As Byte
    Blue As Byte
    Reserved As Byte
End Type
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Const BI_RGB  As Long = 0&
Private Declare Function GetProcessHeap Lib "kernel32.dll" () As Long
Private Declare Function HeapAlloc Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long 'Pointer to mem
Private Declare Function HeapFree Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long) As Long 'BOOL
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef dest As Any, ByRef src As Any, ByVal dwLen As Long)
Private Const HEAP_ZERO_MEMORY As Long = &H8
Private m_memBits() As Byte
Private m_memBitmapInfo() As Byte
Private m_bih As BITMAPINFOHEADER
Private m_bfh As BITMAPFILEHEADER
Public Function CreateFromFile(ByVal filename As String) As Boolean
    Dim hFile As Long
    If Not ExistFile(filename) Then
        MsgBox "File does not exist:" & vbCrLf & filename, vbCritical, App.title
        Exit Function
    End If
    hFile = FreeFile()
    On Error Resume Next
    Open filename For Binary Access Read As #hFile
    If Err Then
        If Err.Number = 70 Then
            MsgBox "File is locked - cannot access:" & vbCrLf & filename, vbCritical, App.title
        Else
            MsgBox Err.Description, vbInformation, App.title
        End If
        Exit Function
    End If
    On Error GoTo 0
    Get #hFile, , m_bfh
    If m_bfh.bfType <> BMP_MAGIC_COOKIE Then
        MsgBox "File is not a supported bitmap format:" & vbCrLf & filename, vbInformation, App.title
        Close #hFile
        Exit Function
    Else
        Get #hFile, Len(m_bfh) + 1, m_bih
        ReDim m_memBits(0 To m_bih.biSizeImage - 1)
        Get #hFile, m_bfh.bfOffBits + 1, m_memBits
        ReDim m_memBitmapInfo(0 To m_bfh.bfOffBits - 14)
        Get #hFile, Len(m_bfh) + 1, m_memBitmapInfo
        Close #hFile
    End If
    CreateFromFile = True
End Function
Public Function CreateFromPackedDIBPointer(ByRef pDIB As Long) As Boolean
        Debug.Assert pDIB <> 0
        Call CopyMemory(ByVal VarPtr(m_bih.biSize), ByVal pDIB, Len(m_bih))
        If m_bih.biBitCount < 16 Then
            Debug.Print "Error! DIB was less than 16 colors."
            Exit Function
        End If
        If m_bih.biSizeImage < 1 Then Exit Function
        ReDim m_memBits(0 To m_bih.biSizeImage - 1)
        Call CopyMemory(m_memBits(0), ByVal pDIB + 40, m_bih.biSizeImage)
        ReDim m_memBitmapInfo(0 To 39)
        Call CopyMemory(m_memBitmapInfo(0), m_bih, Len(m_bih))
        With m_bfh
            .bfType = BMP_MAGIC_COOKIE
            .bfSize = 55 + m_bih.biSizeImage
            .bfReserved1 = 0&
            .bfReserved2 = 0&
            .bfOffBits = 54
        End With
        CreateFromPackedDIBPointer = True
End Function
Public Function WriteToFile(ByVal filename As String) As Boolean
        Dim hFile As Integer
        On Error Resume Next
        hFile = FreeFile()
        Open filename For Binary As hFile
            Put hFile, 1, m_bfh
            Put hFile, Len(m_bfh) + 1, m_memBitmapInfo
            Put hFile, , m_memBits
        Close hFile
        WriteToFile = True
End Function
Private Function ExistFile(ByVal sSpec As String) As Boolean
    On Error Resume Next
    Call FileLen(sSpec)
    ExistFile = (Err = 0)
End Function
Public Property Get BitCount() As Long
    BitCount = m_bih.biBitCount
End Property
Public Property Get Height() As Long
    Height = m_bih.biHeight
End Property
Public Property Get Width() As Long
    Width = m_bih.biWidth
End Property
Public Property Get Compression() As Long
    Compression = m_bih.biCompression
End Property
Public Property Get SizeInfoHeader() As Long
    SizeInfoHeader = m_bih.biSize
End Property
Public Property Get SizeImage() As Long
    SizeImage = m_bih.biSizeImage
End Property
Public Property Get Planes() As Long
    Planes = m_bih.biPlanes
End Property
Public Property Get ClrImportant() As Long
    ClrImportant = m_bih.biClrImportant
End Property
Public Property Get ClrUsed() As Long
    ClrUsed = m_bih.biClrUsed
End Property
Public Property Get XPPM() As Long
    XPPM = m_bih.biXPelsPerMeter
End Property
Public Property Get YPPM() As Long
    YPPM = m_bih.biYPelsPerMeter
End Property
Public Property Get FileType() As Long
    FileType = m_bfh.bfType
End Property
Public Property Get SizeFileHeader() As Long
    SizeFileHeader = m_bfh.bfSize
End Property
Public Property Get BitOffset() As Long
    BitOffset = m_bfh.bfOffBits
End Property
Public Property Get PointerToBits() As Long
    PointerToBits = VarPtr(m_memBits(0))
End Property
Public Property Get PointerToBitmapInfo() As Long
    PointerToBitmapInfo = VarPtr(m_memBitmapInfo(0))
End Property
Public Property Get SizeBitmapInfo() As Long
    SizeBitmapInfo = UBound(m_memBitmapInfo()) + 1
End Property




