'Option Explicit
Private Declare Function GetSystemDirectory Lib "kernel32" Alias _
                        "GetSystemDirectoryA" (ByVal lpBuffer As String, _
                                               ByVal nSize As Long) As Long
Public Declare Function AVIFileInfo Lib "avifil32.dll" _
                    (ByVal pFile As Long, _
                    pfi As AVI_FILE_INFO, _
                    ByVal lSize As Long) As Long 'HRESULT
Public Declare Function AVIFileCreateStream Lib _
                                        "avifil32.dll" Alias "AVIFileCreateStreamA" _
                                        (ByVal pFile As Long, _
                                         ByRef ppavi As Long, _
                                         ByRef psi As AVI_STREAM_INFO) As Long
Public Declare Function mmioStringToFOURCC Lib "winmm.dll" Alias _
          "mmioStringToFOURCCA" (ByVal sz As String, ByVal uFlags As Long) As Long 'returns fourcc
Public Declare Function VideoForWindowsVersion Lib "msvfw32.dll" () As Long
Public Declare Function AVIFileOpen Lib "avifil32.dll" _
       (ByRef ppfile As Long, _
        ByVal szFile As String, _
        ByVal uMode As Long, _
        ByVal pclsidHandler As Long) As Long  'HRESULT
Public Declare Function AVISaveOptions Lib "avifil32.dll" (ByVal hwnd As Long, _
                        ByVal uiFlags As Long, _
                        ByVal nStreams As Long, _
                        ByRef ppavi As Long, _
                        ByRef ppOptions As Long) As Long 'TRUE if user pressed OK, False if cancel, or error if error
Public Declare Sub AVIFileInit Lib "avifil32.dll" ()
Public Declare Function AVISave Lib "avifil32.dll" Alias "AVISaveVA" (ByVal szFile As String, _
           ByVal pclsidHandler As Long, _
           ByVal lpfnCallback As Long, _
           ByVal nStreams As Long, _
           ByRef ppaviStream As Long, _
           ByRef ppCompOptions As Long) As Long
Public Declare Function AVISaveOptionsFree Lib "avifil32.dll" (ByVal nStreams As Long, _
                     ByRef ppOptions As Long) As Long
Public Declare Function AVIMakeCompressedStream Lib "avifil32.dll" (ByRef ppsCompressed As Long, _
                        ByVal psSource As Long, _
                        ByRef lpOptions As AVI_COMPRESS_OPTIONS, _
                        ByVal pclsidHandler As Long) As Long '
Public Declare Function AVIStreamWrite Lib "avifil32.dll" (ByVal pavi As Long, _
                 ByVal lStart As Long, _
                 ByVal lSamples As Long, _
                 ByVal lpBuffer As Long, _
                 ByVal cbBuffer As Long, _
                 ByVal dwFlags As Long, _
                 ByRef plSampWritten As Long, _
                 ByRef plBytesWritten As Long) As Long
Public Declare Function AVIStreamSetFormat Lib "avifil32.dll" (ByVal pavi As Long, _
          ByVal lPos As Long, _
          ByRef lpFormat As Any, _
          ByVal cbFormat As Long) As Long
Public Declare Function AVIStreamGetFrameOpen Lib "avifil32.dll" (ByVal pAVIStream As Long, _
                                   ByRef bih As Any) As Long
Public Declare Function AVIStreamReadFormat Lib "avifil32.dll" (ByVal pAVIStream As Long, _
                    ByVal lPos As Long, _
                    ByVal lpFormatBuf As Long, _
                    ByRef sizeBuf As Long) As Long
Public Declare Function AVIStreamRead Lib "avifil32.dll" (ByVal pAVIStream As Long, _
                                                            ByVal lStart As Long, _
                                                            ByVal lSamples As Long, _
                                                            ByVal lpBuffer As Long, _
                                                            ByVal cbBuffer As Long, _
                                                            ByRef pBytesWritten As Long, _
                                                            ByRef pSamplesWritten As Long) As Long
Public Declare Function AVIStreamGetFrameClose Lib "avifil32.dll" (ByVal pGetFrameObj As Long) As Long
Public Declare Function AVIPutFileOnClipboard Lib "avifil32.dll" (ByVal pAVIFile As Long) As Long
Public Declare Function AVIFileRelease Lib "avifil32.dll" (ByVal pFile As Long) As Long
Public Declare Function AVIFileGetStream Lib "avifil32.dll" _
                        (ByVal pFile As Long, _
                         ByRef ppaviStream As Long, _
                         ByVal fccType As Long, _
                         ByVal lParam As Long) As Long
Public Declare Function AVIMakeFileFromStreams Lib "avifil32.dll" _
          (ByRef ppfile As Long, _
           ByVal nStreams As Long, _
           ByVal pAVIStreamArray As Long) As Long
Public Declare Function AVIStreamInfo Lib "avifil32.dll" _
                                          (ByVal pAVIStream As Long, _
                                           ByRef psi As AVI_STREAM_INFO, _
                                           ByVal lSize As Long) As Long
Public Declare Function AVIStreamGetFrame Lib "avifil32.dll" (ByVal pGetFrameObj As Long, _
                                                              ByVal lPos As Long) As Long
Public Declare Function AVIStreamRelease Lib "avifil32.dll" (ByVal pavi As Long) As Long 'ULONG
Public Declare Function AVIStreamClose Lib "avifil32.dll" _
                                       Alias "AVIStreamRelease" _
                                      (ByVal pavi As Long) As Long 'ULONG
Public Declare Function AVIStreamLength Lib "avifil32.dll" (ByVal pavi As Long) As Long
Public Declare Function AVIFileClose Lib "avifil32.dll" Alias "AVIFileRelease" (ByVal pFile As Long) As Long
Public Declare Sub AVIFileExit Lib "avifil32.dll" ()
Public Declare Function AVIMakeStreamFromClipboard Lib "avifil32.dll" _
              (ByVal cfFormat As Long, _
               ByVal hGlobal As Long, _
               ByRef ppstream As Long) As Long
Public Declare Function AVIStreamStart Lib "avifil32.dll" (ByVal pavi As Long) As Long
Public Declare Function AVIGetFromClipboard Lib "avifil32.dll" (ByRef ppAVIFile As Long) As Long
Public Declare Function AVIClearClipboard Lib "avifil32.dll" () As Long
Public Const BMP_MAGIC_COOKIE As Integer = 19778
Public Type BITMAPFILEHEADER
        bfType        As Integer
        bfSize        As Long
        bfReserved1   As Integer
        bfReserved2   As Integer
        bfOffBits     As Long
End Type
Public Type BITMAPINFOHEADER
   biSize          As Long
   biWidth         As Long
   biHeight        As Long
   biPlanes        As Integer
   biBitCount      As Integer
   biCompression   As Long
   biSizeImage     As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed       As Long
   biClrImportant  As Long
End Type
Public Type BITMAPINFOHEADER_MJPEG
   biSize            As Long
   biWidth           As Long
   biHeight          As Long
   biPlanes          As Integer
   biBitCount        As Integer
   biCompression     As Long
   biSizeImage       As Long
   biXPelsPerMeter   As Long
   biYPelsPerMeter   As Long
   biClrUsed         As Long
   biClrImportant    As Long
   biExtDataOffset   As Long
   JPEGSize          As Long
   JPEGProcess       As Long
   JPEGColorSpaceID  As Long
   JPEGBitsPerSample As Long
   JPEGHSubSampling  As Long
   JPEGVSubSampling  As Long
End Type

 Public Type AVI_RECT
    left    As Long
    top     As Long
    right   As Long
    bottom  As Long
End Type
Public Type AVI_STREAM_INFO
    fccType               As Long
    fccHandler            As Long
    dwFlags               As Long
    dwCaps                As Long
    wPriority             As Integer
    wLanguage             As Integer
    dwScale               As Long
    dwRate                As Long
    dwStart               As Long
    dwLength              As Long
    dwInitialFrames       As Long
    dwSuggestedBufferSize As Long
    dwQuality             As Long
    dwSampleSize          As Long
    rcFrame               As AVI_RECT
    dwEditCount           As Long
    dwFormatChangeCount   As Long
    szName                As String * 64
End Type
Public Type AVI_FILE_INFO
    dwMaxBytesPerSecond   As Long
    dwFlags               As Long
    dwCaps                As Long
    dwStreams             As Long
    dwSuggestedBufferSize As Long
    dwWidth               As Long
    dwHeight              As Long
    dwScale               As Long
    dwRate                As Long
    dwLength              As Long
    dwEditCount           As Long
    szFileType            As String * 64
End Type
Public Type AVI_COMPRESS_OPTIONS
    fccType           As Long
    fccHandler        As Long
    dwKeyFrameEvery   As Long
    dwQuality         As Long
    dwBytesPerSecond  As Long
    dwFlags           As Long
    lpFormat          As Long
    cbFormat          As Long
    lpParms           As Long
    cbParms           As Long
    dwInterleaveEvery As Long
End Type
Public Type RenderInfo
    filename    As String
    Size_Height As Integer
    Size_Width  As Integer
    Codec       As Long
    FrameRate   As Integer
    Query       As Integer
    pOpts       As Long
    ps          As Long
    Opts        As AVI_COMPRESS_OPTIONS
    List        As ListBox
    pFile       As Long
    RenderMode  As Integer
End Type '--------------------------------
Global Const AVIERR_OK As Long = 0&
Private Const SEVERITY_ERROR    As Long = &H80000000
Private Const FACILITY_ITF      As Long = &H40000
Private Const AVIERR_BASE       As Long = &H4000
Global Const AVIERR_BADFLAGS    As Long = SEVERITY_ERROR Or FACILITY_ITF Or (AVIERR_BASE + 105) '-2147205015
Global Const AVIERR_BADPARAM    As Long = SEVERITY_ERROR Or FACILITY_ITF Or (AVIERR_BASE + 106) '-2147205014
Global Const AVIERR_BADSIZE     As Long = SEVERITY_ERROR Or FACILITY_ITF Or (AVIERR_BASE + 107) '-2147205013
Global Const AVIERR_USERABORT   As Long = SEVERITY_ERROR Or FACILITY_ITF Or (AVIERR_BASE + 198) '-2147204922
Global Const AVIFILEINFO_HASINDEX         As Long = &H10
Global Const AVIFILEINFO_MUSTUSEINDEX     As Long = &H20
Global Const AVIFILEINFO_ISINTERLEAVED    As Long = &H100
Global Const AVIFILEINFO_WASCAPTUREFILE   As Long = &H10000
Global Const AVIFILEINFO_COPYRIGHTED      As Long = &H20000
Global Const AVIFILECAPS_CANREAD          As Long = &H1
Global Const AVIFILECAPS_CANWRITE         As Long = &H2
Global Const AVIFILECAPS_ALLKEYFRAMES     As Long = &H10
Global Const AVIFILECAPS_NOCOMPRESSION    As Long = &H20
Global Const AVICOMPRESSF_INTERLEAVE     As Long = &H1           '// interleave
Global Const AVICOMPRESSF_DATARATE       As Long = &H2           '// use a data rate
Global Const AVICOMPRESSF_KEYFRAMES      As Long = &H4           '// use keyframes
Global Const AVICOMPRESSF_VALID          As Long = &H8           '// has valid data?
Global Const AVIGETFRAMEF_BESTDISPLAYFMT  As Long = 1
Global Const ICMF_CHOOSE_KEYFRAME           As Long = &H1     '// show KeyFrame Every box
Global Const ICMF_CHOOSE_DATARATE           As Long = &H2     '// show DataRate box
Global Const ICMF_CHOOSE_PREVIEW            As Long = &H4     '// allow expanded preview dialog
Global Const ICMF_CHOOSE_ALLCOMPRESSORS     As Long = &H8     '// don't only show those that
Global Const OF_READ             As Long = &H0
Global Const OF_WRITE            As Long = &H1
Global Const OF_SHARE_DENY_WRITE As Long = &H20
Global Const OF_CREATE           As Long = &H1000
Global Const streamtypeVIDEO       As Long = 1935960438 'equivalent to: mmioStringToFOURCC("vids", 0&)
Global Const streamtypeAUDIO       As Long = 1935963489 'equivalent to: mmioStringToFOURCC("auds", 0&)
Global Const streamtypeMIDI        As Long = 1935960429 'equivalent to: mmioStringToFOURCC("mids", 0&)
Global Const streamtypeTEXT        As Long = 1937012852
Global Const AVIIF_KEYFRAME  As Long = &H10
Global Const DIB_RGB_COLORS  As Long = 0
Global Const DIB_PAL_COLORS  As Long = 1
Global Const BI_RGB          As Long = 0
Global Const BI_RLE8         As Long = 1
Global Const BI_RLE4         As Long = 2
Global Const BI_BITFIELDS    As Long = 3
Public Declare Function GetProcessHeap Lib "kernel32.dll" () As Long 'handle
Public Declare Function SetRect Lib "user32.dll" _
             (ByRef lprc As AVI_RECT, _
              ByVal xLeft As Long, _
              ByVal yTop As Long, _
              ByVal xRight As Long, _
              ByVal yBottom As Long) As Long 'BOOL
Public Declare Function HeapFree Lib "kernel32.dll" _
                        (ByVal hHeap As Long, _
                         ByVal dwFlags As Long, _
                         ByVal lpMem As Long) As Long 'BOOL
Public Declare Function HeapAlloc Lib "kernel32.dll" _
        (ByVal hHeap As Long, _
         ByVal dwFlags As Long, _
         ByVal dwBytes As Long) As Long 'Pointer to mem
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" _
                                    (ByRef dest As Any, _
                                     ByRef src As Any, _
                                     ByVal dwLen As Long)
Public Const HEAP_ZERO_MEMORY As Long = &H8
Global gfAbort       As Boolean
Public Can           As Boolean
Public bmp           As cDIB
Public t6            As Integer, cunt As Integer
Public cunt2         As Integer
Public Saved         As Boolean
Private Templst As ListBox

'---------------------------------------
Public Sub Comperes(inf As RenderInfo)
On Error GoTo 4
        Dim i%, Value%
        Can = False
        frmRender.ProgressBar1.Max = inf.List.ListCount
        frmRender.ProgressBar2.Max = inf.List.ListCount * 2
        frmRender.Show
        frmRender.Picture1.Width = (inf.Size_Width + 4) * 15
        frmRender.Picture1.Height = (inf.Size_Height + 4) * 15
        frmRender.Picture3.Width = (inf.Size_Width + 4) * 15
        frmRender.Picture3.Height = (inf.Size_Height + 4) * 15
        frmRender.Picture2.Width = (inf.Size_Width + 4) * 15
        frmRender.Picture2.Height = (inf.Size_Height + 4) * 15
        If t6 = 20 And inf.RenderMode = 0 Then GoTo 5
        If t6 = 21 And inf.RenderMode = 1 Then GoTo 5
        frmRender.List1.Clear
        Set Templst = TimeLine.List1
        Templst.Clear
        If FSO.FolderExists(App.Path & "\temp") = True Then FSO.DeleteFolder (App.Path & "\temp")
        FSO.CreateFolder (App.Path & "\temp")
        For i = 0 To inf.List.ListCount - 1
            If Can = True Then Exit For
            With frmRender.Picture1
                .Picture = LoadPicture(inf.List.List(i))
                .PaintPicture .Picture, 0, 0, .Width, .Height
                Set .Picture = .Image
            End With
            If inf.RenderMode = 1 Then ZipIr3 inf, frmRender.Picture1
           SavePicture frmRender.Picture1.Picture, App.Path & "\temp\" & i & ".bmp"
           frmRender.List1.AddItem App.Path & "\temp\" & i & ".bmp"
           frmRender.List1.AddItem App.Path & "\temp\" & i & ".bmp"
           Templst.AddItem App.Path & "\temp\" & i & ".bmp"
           Templst.AddItem App.Path & "\temp\" & i & ".bmp"
           frmRender.ProgressBar1.Value = i
           frmRender.ProgressBar2.Value = i: frmRender.ProgressBar2.Refresh
           DoEvents
           frmRender.Label1.Caption = "File's::" & inf.List.List(i)
           frmRender.Caption = "Buffering ... " & left(Val((i / inf.List.ListCount) * 100), 5) & " %"
        Next i
        If Can = True Then
            Unload frmRender
            Main.Show
            Exit Sub
        End If
   Set inf.List = frmRender.List1
5  If t6 <> 0 Then Set inf.List = Templst
        If inf.RenderMode = 0 Then Render inf
        If inf.RenderMode = 1 Then Save inf
4 End Sub
Public Sub Render(inf As RenderInfo)
On Error GoTo 4
    Dim Dars             As String
    Dim FF               As String
    Dim psCompressed     As Long
    Dim strhdr           As AVI_STREAM_INFO
    Dim BI               As BITMAPINFOHEADER
    Dim i                As Long
    Dim a                As Long
    Dim res              As Long
    Dim Fil              As String * 255
    frmRender.Show
    res = inf.Codec
    t6 = 20
    If res <> 1 Then
        Call AVISaveOptionsFree(1, inf.pOpts)
        GoTo Error
    End If
    res = AVIMakeCompressedStream(psCompressed, inf.ps, inf.Opts, 0&)
    If res <> AVIERR_OK Then GoTo Error
    With BI
        .biBitCount = bmp.BitCount
        .biClrImportant = bmp.ClrImportant
        .biClrUsed = bmp.ClrUsed
        .biCompression = bmp.Compression
        .biHeight = bmp.Height
        .biWidth = bmp.Width
        .biPlanes = bmp.Planes
        .biSize = bmp.SizeInfoHeader
        .biSizeImage = bmp.SizeImage
        .biXPelsPerMeter = bmp.XPPM
        .biYPelsPerMeter = bmp.YPPM
    End With
    res = AVIStreamSetFormat(psCompressed, 0, ByVal bmp.PointerToBitmapInfo, bmp.SizeBitmapInfo)
    If (res <> AVIERR_OK) Then GoTo Error
    frmRender.ProgressBar1.Max = inf.List.ListCount
    frmRender.Label2.Caption = "Saveing to " & inf.filename
'------------------------Start Saving VideoFile---------------------------------------
    For i = 0 To inf.List.ListCount - 1
        If Can = True Then Exit For
        frmRender.Label1.Caption = "File's:" + inf.List.List(i)
        frmRender.ProgressBar1.Value = i
        If frmRender.ProgressBar2.Value < frmRender.ProgressBar2.Max And i Mod 2 = 0 _
                Then frmRender.ProgressBar2.Value = frmRender.ProgressBar2.Value + 1
        Dars = Mid(Str((i * 10) / Val(inf.List.ListCount / 10)), 2, 5) + " %"
        frmRender.Caption = "Rendering..." & Dars
        bmp.CreateFromFile (inf.List.List(i))
        res = AVIStreamWrite(psCompressed, i, 1, _
                                        bmp.PointerToBits, bmp.SizeImage, AVIIF_KEYFRAME, ByVal 0&, ByVal 0&)
        If res <> AVIERR_OK Then GoTo Error
        DoEvents
    Next
Error:
         Unload frmRender
         Main.Visible = True
         Set bmp = Nothing
         If (inf.ps <> 0) Then Call AVIStreamClose(inf.ps)
         If (psCompressed <> 0) Then Call AVIStreamClose(psCompressed)
         If (inf.pFile <> 0) Then Call AVIFileClose(inf.pFile)
         Call AVIFileExit
         If (res <> AVIERR_OK) Then MsgBox "There was an error writing the file.", vbInformation, "IranVideo5.3"
         '-------------------Deleting Temp File In Error-------------------------
        Command1 = inf.filename
        Player.Show
4 End Sub
'--------------------------------------------------------------

