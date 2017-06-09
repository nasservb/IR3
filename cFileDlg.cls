VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private Const MAX_PATH = 1024
Private Const MAX_FILE = 512
Public Enum SPECIAL_FOLDERS
        vbCSIDL_DESKTOP = &H0&
      vbCSIDL_CONTROLS = &H3&
       vbCSIDL_PRINTERS = &H4&
       vbCSIDL_PERSONAL = &H5&
      vbCSIDL_FAVORITES = &H6&
       vbCSIDL_STARTUP = &H7&
       vbCSIDL_RECENT = &H8&
      vbCSIDL_SENDTO = &H9&
       vbCSIDL_BITBUCKET = &HA&
       vbCSIDL_STARTMENU = &HB&
       vbCSIDL_DESKTOPDIRECTORY = &H10&
       vbCSIDL_DRIVES = &H11&
       vbCSIDL_NETWORK = &H12&
       vbCSIDL_NETHOOD = &H13&
      vbCSIDL_FONTS = &H14&
       vbCSIDL_TEMPLATES = &H15&
End Enum
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" _
                                                (ByVal hWndOwner As Long, _
                                                ByVal nFolder As SPECIAL_FOLDERS, _
                                                pidl As Long) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" _
   Alias "SHGetPathFromIDListA" _
  (ByVal pidl As Long, _
   ByVal pszPath As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Const NOERROR As Long = &H0
Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long
End Type
Private Declare Function GetOpenFileName Lib "COMDLG32" _
    Alias "GetOpenFileNameA" (filestruct As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "COMDLG32" _
    Alias "GetSaveFileNameA" (filestruct As OPENFILENAME) As Long
Private Declare Function GetFileTitle Lib "COMDLG32" _
    Alias "GetFileTitleA" (ByVal szFile As String, _
    ByVal szTitle As String, ByVal cbBuf As Integer) As Integer
Private Declare Function GetOpenFileNamePreview Lib "MSVFW32" _
    Alias "GetOpenFileNamePreviewA" (filestruct As OPENFILENAME) As Long
Private Declare Function GetSaveFileNamePreview Lib "MSVFW32" _
    Alias "GetSaveFileNamePreviewA" (filestruct As OPENFILENAME) As Long

Public Enum EOpenFile
    OFN_READONLY = &H1&
    OFN_OVERWRITEPROMPT = &H2&
    OFN_HIDEREADONLY = &H4&
    OFN_NOCHANGEDIR = &H8&
    OFN_SHOWHELP = &H10&
    OFN_ENABLEHOOK = &H20&
    OFN_ENABLETEMPLATE = &H40&
    OFN_ENABLETEMPLATEHANDLE = &H80&
    OFN_NOVALIDATE = &H100&
    OFN_ALLOWMULTISELECT = &H200&
    OFN_EXTENSIONDIFFERENT = &H400&
    OFN_PATHMUSTEXIST = &H800&
    OFN_FILEMUSTEXIST = &H1000&
    OFN_CREATEPROMPT = &H2000&
    OFN_SHAREAWARE = &H4000&
    OFN_NOREADONLYRETURN = &H8000&
    OFN_NOTESTFILECREATE = &H10000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOLONGNAMES = &H40000
    OFN_EXPLORER = &H80000
    OFN_NODEREFERENCELINKS = &H100000
    OFN_LONGNAMES = &H200000
End Enum
Private Declare Function CommDlgExtendedError Lib "COMDLG32" () As Long
Public Enum EDialogError
    CDERR_DIALOGFAILURE = &HFFFF&
    CDERR_GENERALCODES = &H0&
    CDERR_STRUCTSIZE = &H1&
    CDERR_INITIALIZATION = &H2&
    CDERR_NOTEMPLATE = &H3&
    CDERR_NOHINSTANCE = &H4&
    CDERR_LOADSTRFAILURE = &H5&
    CDERR_FINDRESFAILURE = &H6&
    CDERR_LOADRESFAILURE = &H7&
    CDERR_LOCKRESFAILURE = &H8&
    CDERR_MEMALLOCFAILURE = &H9&
    CDERR_MEMLOCKFAILURE = &HA&
    CDERR_NOHOOK = &HB&
    CDERR_REGISTERMSGFAIL = &HC&

    PDERR_PRINTERCODES = &H1000&
    PDERR_SETUPFAILURE = &H1001&
    PDERR_PARSEFAILURE = &H1002&
    PDERR_RETDEFFAILURE = &H1003&
    PDERR_LOADDRVFAILURE = &H1004&
    PDERR_GETDEVMODEFAIL = &H1005&
    PDERR_INITFAILURE = &H1006&
    PDERR_NODEVICES = &H1007&
    PDERR_NODEFAULTPRN = &H1008&
    PDERR_DNDMMISMATCH = &H1009&
    PDERR_CREATEICFAILURE = &H100A&
    PDERR_PRINTERNOTFOUND = &H100B&
    PDERR_DEFAULTDIFFERENT = &H100C&
    CFERR_CHOOSEFONTCODES = &H2000&
    CFERR_NOFONTS = &H2001&
    CFERR_MAXLESSTHANMIN = &H2002&
    FNERR_FILENAMECODES = &H3000&
    FNERR_SUBCLASSFAILURE = &H3001&
    FNERR_INVALIDFILENAME = &H3002&
    FNERR_BUFFERTOOSMALL = &H3003&
    CCERR_CHOOSECOLORCODES = &H5000&
End Enum
Private Const sEmpty As String = ""
Private m_FileMustExist As Boolean
Private m_MultiSelect As Boolean
Private m_ReadOnly As Boolean
Private m_HideReadOnly As Boolean
Private m_Filter As String
Private m_FilterIndex As Long
Private m_InitDir As String
Private m_DlgTitle As String
Private m_DefaultExt As String
Private m_Flags As Long
Private m_OverwritePrompt As Boolean
Private m_hWnd As Long
Private Sub Class_Initialize()
m_hWnd = -1&
m_FileMustExist = True
m_MultiSelect = False
m_HideReadOnly = True
m_DlgTitle = App.title
m_OverwritePrompt = True
m_InitDir = GetSpecialFolderLocation(vbCSIDL_PERSONAL) 'default to My Documents folder
End Sub
Property Get OwnerhWnd() As Long
    OwnerhWnd = m_hWnd
End Property
Property Let OwnerhWnd(ByVal vHwnd As Long)
    m_hWnd = vHwnd
End Property
Property Get FileMustExist() As Boolean
    FileMustExist = m_FileMustExist
End Property
Property Let FileMustExist(ByVal vNewValue As Boolean)
    m_FileMustExist = vNewValue
End Property
Property Get MultiSelect() As Boolean
    MultiSelect = m_MultiSelect
End Property
Property Let MultiSelect(ByVal vNewValue As Boolean)
    m_MultiSelect = vNewValue
End Property
Property Get ReadOnly() As Boolean
    ReadOnly = m_ReadOnly
End Property
Property Get HideReadOnly() As Boolean
    HideReadOnly = m_HideReadOnly
End Property
Property Let HideReadOnly(ByVal vNewValue As Boolean)
    m_HideReadOnly = vNewValue
End Property
Property Get Filter() As String
    Filter = m_Filter
End Property
Property Let Filter(ByVal vFilterString As String)
    m_Filter = vFilterString
End Property
Property Get FilterIndex() As Long
    FilterIndex = m_FilterIndex
End Property
Property Let FilterIndex(ByVal vIndex As Long)
    m_FilterIndex = vIndex
End Property
Property Get InitDirectory() As String
    InitDirectory = m_InitDir
End Property
Property Let InitDirectory(ByVal DirPath As String)
    m_InitDir = DirPath
End Property
Property Let InitDirectorySpecial(ByVal SpecialDirectory As SPECIAL_FOLDERS)
    m_InitDir = GetSpecialFolderLocation(SpecialDirectory)
End Property
Property Get DlgTitle() As String
    DlgTitle = m_DlgTitle
End Property
Property Let DlgTitle(ByVal title As String)
    m_DlgTitle = title
End Property
Property Get DefaultExt() As String
    DefaultExt = m_DefaultExt
End Property
Property Let DefaultExt(ByVal fileExt As String)
    m_DefaultExt = fileExt
End Property
Property Get Flags() As EOpenFile
    Flags = m_Flags
End Property
Property Let Flags(ByVal vFlags As EOpenFile)
    m_Flags = vFlags
End Property
Property Get OverwritePrompt() As Boolean
    OverwritePrompt = m_OverwritePrompt
End Property
Property Let OverwritePrompt(ByVal vShowPrompt As Boolean)
    m_OverwritePrompt = vShowPrompt
End Property
Public Function VBGetOpenFileName(filename As String, _
                           Optional FileTitle As String) As Boolean

    Dim opfile As OPENFILENAME
    Dim s As String
    Dim afFlags As Long
    Dim ch As String
    Dim i As Integer
    
With opfile
    .lStructSize = Len(opfile)
            .Flags = (-m_FileMustExist * OFN_FILEMUSTEXIST) Or _
             (-m_MultiSelect * OFN_ALLOWMULTISELECT) Or _
             (-m_ReadOnly * OFN_READONLY) Or _
             (-m_HideReadOnly * OFN_HIDEREADONLY) Or _
             (m_Flags And CLng(Not (OFN_ENABLEHOOK Or _
                                  OFN_ENABLETEMPLATE)))
       If m_hWnd <> -1 Then .hWndOwner = m_hWnd
       .lpstrInitialDir = m_InitDir
        .lpstrDefExt = m_DefaultExt
        .lpstrTitle = m_DlgTitle
          For i = 1 To Len(Filter)
        ch = Mid$(Filter, i, 1)
        If ch = "|" Or ch = ":" Then
            s = s & vbNullChar
        Else
            s = s & ch
        End If
    Next
       s = s & vbNullChar & vbNullChar
    .lpstrFilter = s
    .nFilterIndex = m_FilterIndex
    s = filename & String$(MAX_PATH - Len(filename), 0)
    .lpstrFile = s
    .nMaxFile = MAX_PATH
    s = FileTitle & String$(MAX_FILE - Len(FileTitle), 0)
    .lpstrFileTitle = s
    .nMaxFileTitle = MAX_FILE
        If GetOpenFileName(opfile) Then
        VBGetOpenFileName = True
        filename = Left$(.lpstrFile, InStr(.lpstrFile, vbNullChar) - 1)
        FileTitle = Left$(.lpstrFileTitle, InStr(.lpstrFileTitle, vbNullChar) - 1)
        m_Flags = .Flags
               m_FilterIndex = .nFilterIndex
               m_Filter = FilterLookup(.lpstrFilter, m_FilterIndex)
        If (.Flags And OFN_READONLY) Then m_ReadOnly = True
                m_InitDir = .lpstrFile
    Else
        VBGetOpenFileName = False
        filename = vbNullChar
        FileTitle = vbNullChar
        Flags = 0
        FilterIndex = -1
        Filter = vbNullChar
    End If
End With
End Function
Public Function VBGetOpenFileNamePreview(filename As String, _
                           Optional FileTitle As String) As Boolean

    Dim opfile As OPENFILENAME
    Dim s As String
    Dim afFlags As Long
    Dim ch As String
    Dim i As Integer
    With opfile
    .lStructSize = Len(opfile)
            .Flags = (-m_FileMustExist * OFN_FILEMUSTEXIST) Or _
             (-m_MultiSelect * OFN_ALLOWMULTISELECT) Or _
             (-m_ReadOnly * OFN_READONLY) Or _
             (-m_HideReadOnly * OFN_HIDEREADONLY) Or _
             (m_Flags And CLng(Not (OFN_ENABLEHOOK Or _
                                  OFN_ENABLETEMPLATE)))
      If m_hWnd <> -1 Then .hWndOwner = m_hWnd
       .lpstrInitialDir = m_InitDir
        .lpstrDefExt = m_DefaultExt
        .lpstrTitle = m_DlgTitle
        For i = 1 To Len(Filter)
        ch = Mid$(Filter, i, 1)
        If ch = "|" Or ch = ":" Then
            s = s & vbNullChar
        Else
            s = s & ch
        End If
    Next
        s = s & vbNullChar & vbNullChar
    .lpstrFilter = s
    .nFilterIndex = m_FilterIndex
        s = filename & String$(MAX_PATH - Len(filename), 0)
    .lpstrFile = s
    .nMaxFile = MAX_PATH
    s = FileTitle & String$(MAX_FILE - Len(FileTitle), 0)
    .lpstrFileTitle = s
    .nMaxFileTitle = MAX_FILE
    
    If GetOpenFileNamePreview(opfile) Then
        VBGetOpenFileNamePreview = True
        filename = Left$(.lpstrFile, InStr(.lpstrFile, vbNullChar) - 1)
        FileTitle = Left$(.lpstrFileTitle, InStr(.lpstrFileTitle, vbNullChar) - 1)
        m_Flags = .Flags
               m_FilterIndex = .nFilterIndex
               m_Filter = FilterLookup(.lpstrFilter, m_FilterIndex)
        If (.Flags And OFN_READONLY) Then m_ReadOnly = True
               m_InitDir = .lpstrFile
    Else
        VBGetOpenFileNamePreview = False
        filename = vbNullChar
        FileTitle = vbNullChar
        Flags = 0
        FilterIndex = -1
        Filter = vbNullChar
    End If
End With
End Function
Public Function VBGetSaveFileName(filename As String, _
                           Optional FileTitle As String) As Boolean
            
    Dim opfile As OPENFILENAME, s As String
With opfile
    .lStructSize = Len(opfile)
        .Flags = (-m_OverwritePrompt * OFN_OVERWRITEPROMPT) Or _
             OFN_HIDEREADONLY Or _
             (m_Flags And CLng(Not (OFN_ENABLEHOOK Or _
                                  OFN_ENABLETEMPLATE)))
        If m_hWnd <> -1 Then .hWndOwner = m_hWnd
       .lpstrInitialDir = m_InitDir
       .lpstrDefExt = m_DefaultExt
        .lpstrTitle = m_DlgTitle
            Dim ch As String, i As Integer
    For i = 1 To Len(Filter)
        ch = Mid$(Filter, i, 1)
        If ch = "|" Or ch = ":" Then
            s = s & vbNullChar
        Else
            s = s & ch
        End If
    Next
        s = s & vbNullChar & vbNullChar
    .lpstrFilter = s
    .nFilterIndex = m_FilterIndex
    s = filename & String$(MAX_PATH - Len(filename), 0)
    .lpstrFile = s
    .nMaxFile = MAX_PATH
    s = FileTitle & String$(MAX_FILE - Len(FileTitle), 0)
    .lpstrFileTitle = s
    .nMaxFileTitle = MAX_FILE
       If GetSaveFileName(opfile) Then
        VBGetSaveFileName = True
        filename = Left$(.lpstrFile, InStr(.lpstrFile, vbNullChar) - 1)
        FileTitle = Left$(.lpstrFileTitle, InStr(.lpstrFileTitle, vbNullChar) - 1)
        m_Flags = .Flags
               m_FilterIndex = .nFilterIndex
                m_Filter = FilterLookup(.lpstrFilter, FilterIndex)
                m_InitDir = .lpstrFile
    Else
        VBGetSaveFileName = False
        filename = vbNullChar
        FileTitle = vbNullChar
        m_Flags = 0
        m_FilterIndex = 0
        m_Filter = vbNullChar
    End If
End With
End Function
Private Function FilterLookup(ByVal sFilters As String, ByVal iCur As Long) As String
    Dim iStart As Long
    Dim iEnd As Long
    Dim s As String
        iStart = 1
    If sFilters = vbNullChar Then Exit Function
    Do
         iEnd = InStr(iStart, sFilters, vbNullChar)
        If iEnd = 0 Then Exit Function
        iEnd = InStr(iEnd + 1, sFilters, vbNullChar)
        If iEnd Then
            s = Mid$(sFilters, iStart, iEnd - iStart)
        Else
            s = Mid$(sFilters, iStart)
        End If
        iStart = iEnd + 1
        If iCur = 1 Then
            FilterLookup = s
            Exit Function
        End If
        iCur = iCur - 1
    Loop While iCur
End Function

Private Function StrZToStr(s As String) As String
    Dim TempString As String
    
    TempString = Left$(s, InStr(s, vbNullChar) - 1)
    If TempString = "" Then
                StrZToStr = s
    Else
        StrZToStr = TempString
    End If
End Function
Private Function GetSpecialFolderLocation(CSIDL As SPECIAL_FOLDERS) As String
   Dim sPath As String
   Dim pidl As Long
       If SHGetSpecialFolderLocation(m_hWnd, CSIDL, pidl) = NOERROR Then
          sPath = Space$(MAX_PATH)
      
      If SHGetPathFromIDList(ByVal pidl, ByVal sPath) Then
                 Call CoTaskMemFree(ByVal VarPtr(pidl))
         GetSpecialFolderLocation = Left(sPath, InStr(sPath, Chr$(0)) - 1)
         End If
        End If
   End Function
Public Function ExistFile(ByVal sSpec As String) As Boolean
    On Error Resume Next
    Call FileLen(sSpec)
    ExistFile = (Err = 0)
End Function
Public Function VBGetFileTitle(sFile As String) As String
    Dim sFileTitle As String, cFileTitle As Integer
    cFileTitle = MAX_PATH
    sFileTitle = String$(MAX_PATH, 0)
    cFileTitle = GetFileTitle(sFile, sFileTitle, MAX_PATH)
    If cFileTitle Then
        VBGetFileTitle = ""
    Else
        VBGetFileTitle = Left$(sFileTitle, InStr(sFileTitle, vbNullChar) - 1)
    End If
End Function

