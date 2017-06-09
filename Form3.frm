VERSION 5.00
Begin VB.Form Loader 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   6240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10305
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   6240
   ScaleWidth      =   10305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5280
      Top             =   5280
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "WWW.TCVB.TK"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   3
      Top             =   3720
      Width           =   3255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CopyRight:NasserNiazy,2008 "
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   3240
      Width           =   4335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "IRAN Video Player v 5.3"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   1
      Top             =   2760
      Width           =   4695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "WellCome"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      TabIndex        =   0
      Top             =   1800
      Width           =   4335
   End
End
Attribute VB_Name = "Loader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ASD As New FileSystemObject
Private Declare Function GetSystemDirectory Lib "kernel32" Alias _
                        "GetSystemDirectoryA" (ByVal lpBuffer As String, _
                                               ByVal nSize As Long) As Long
Private Enum HKEY_Type
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum
'----------------------------
Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Boolean
End Type
'Registry Entry Types
'------------------------------------------------------------------------
Public Enum Reg_Type
    REG_NONE = 0                  'No data type.
    REG_SZ = 1                    'A string terminated by a null character.
    REG_EXPAND_SZ = 2             'A null-terminated string which contains unexpanded environment variables.
    REG_BINARY = 3                'A non-text sequence of bytes.
    REG_DWORD = 4                 'Same as REG_DWORD_LITTLE_ENDIAN.
    REG_DWORD_LITTLE_ENDIAN = 4   'A 32-bit integer stored in little-endian format. This is the way Intel-based computers normally store numbers.
    REG_DWORD_BIG_ENDIAN = 5      'A 32-bit integer stored in big-endian format. This is the opposite of the way Intel-based computers normally store numbers -- the word order is reversed.
    REG_LINK = 6                  'A Unicode symbolic link.
    REG_MULTI_SZ = 7              'A series of strings, each separated by a null character and the entire set terminated by a two null characters.
    REG_RESOURCE_LIST = 8         'A list of resources in the resource map.
End Enum
'Secuirty Constants
'------------------------------------------------------------------------
Const KEY_ALL_ACCESS = &HF003F      'Permission for all types of access.
Const KEY_CREATE_LINK = &H20        'Permission to create symbolic links.
Const KEY_CREATE_SUB_KEY = &H4      'Permission to create subkeys.
Const KEY_ENUMERATE_SUB_KEYS = &H8  'Permission to enumerate subkeys.
Const KEY_EXECUTE = &H20019         'Same as KEY_READ.
Const KEY_NOTIFY = &H10             'Permission to give change notification.
Const KEY_QUERY_VALUE = &H1         'Permission to query subkey data.
Const KEY_READ = &H20019            'Permission for general read access.
Const KEY_SET_VALUE = &H2           'Permission to set subkey data.
Const KEY_WRITE = &H20006           'Permission for general write access.
'----------------------------------------------------------------------
'Error Numbers
'------------------------------------------------------------------------
Const REG_ERR_OK = 0                'No Problems
Const REG_ERR_NOT_EXIST = 1         'Key does not exist
Const REG_ERR_NOT_STRING = 2        'Value is not a string
Const REG_ERR_NOT_DWORD = 4         'Value not DWORD
'
Const ERROR_NONE = 0
Const ERROR_BADDB = 1
Const ERROR_BADKEY = 2
Const ERROR_CANTOPEN = 3
Const ERROR_CANTREAD = 4
Const ERROR_CANTWRITE = 5
Const ERROR_OUTOFMEMORY = 6
Const ERROR_ARENA_TRASHED = 7
Const ERROR_ACCESS_DENIED = 8
Const ERROR_INVALID_PARAMETERS = 87
Const ERROR_NO_MORE_ITEMS = 259
'--------------------------------------------------------------------------
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
                (ByVal hKey As Long, _
                 ByVal lpValueName As String, _
                 ByVal Reserved As Long, _
                 ByVal dwType As Long, _
                 lpData As Any, _
                 ByVal cbData As Long) As Long
Private Declare Function RegCreatekey Lib "advapi32.dll" Alias "RegCreateKeyA" _
                (ByVal hKey As Long, _
                 ByVal lpSubKey As String, _
                 phkResult As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" _
                (ByVal hKey As Long, _
                 ByVal lpSubKey As String, _
                 ByVal Reserved As Long, _
                 ByVal lpClass As String, _
                 ByVal dwOptions As Long, _
                 ByVal samDesired As Long, _
                 lpSecurityAttributes As SECURITY_ATTRIBUTES, _
                 phkResult As Long, _
                 lpdwDisposition As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Dim f%
Private Function CreateKey(Key As HKEY_Type, sSubKey As String) As Boolean
On Error GoTo 4
    Dim hKey As Long
    Dim retval As Long
    
    retval = RegCreatekey(Key, sSubKey, hKey)
    If retval <> ERROR_NONE Then
        CreateKey = False
    Else
        retval = RegCloseKey(hKey)
        CreateKey = True
    End If
4
End Function
'------------------------------
Private Function WriteString(Key As HKEY_Type, SubKey As String, sName As String, sData As String) As Boolean
On Error GoTo 4
    Dim hKey As Long
    Dim retval As Long
    Dim deposit As Long
    Dim secattr As SECURITY_ATTRIBUTES
    
    secattr.nLength = Len(secattr)
    secattr.lpSecurityDescriptor = 0
    secattr.bInheritHandle = 1
    
    retval = RegCreateKeyEx(Key, SubKey, 0, "", 0, KEY_WRITE, secattr, hKey, deposit)
    If retval <> ERROR_NONE Then
        WriteString = False
        Exit Function
    End If

    retval = RegSetValueEx(hKey, sName, 0, REG_SZ, ByVal sData, Len(sData))
    
    If retval <> ERROR_NONE Then
        WriteString = False
        Exit Function
    End If
    
    retval = RegCloseKey(hKey)
    WriteString = True
4 End Function


Private Sub Form_Load()
On Error GoTo 4
        Dim FG As String
        If App.PrevInstance = True Then Unload Me
        FG = App.Path & "\iranvideo5.exe"
        Call CreateKey(HKEY_CLASSES_ROOT, ".irp")
        Call WriteString(HKEY_CLASSES_ROOT, ".irp", "", "Nasservb.irp")
        Call WriteString(HKEY_CLASSES_ROOT, ".irp", "PerceivedType", "IR3Project")
        Call CreateKey(HKEY_CLASSES_ROOT, ".ir3")
        Call WriteString(HKEY_CLASSES_ROOT, ".ir3", "", "Nasservb.ir3")
        Call WriteString(HKEY_CLASSES_ROOT, ".ir3", "PerceivedType", "IRAN Video File")
        Call CreateKey(HKEY_CLASSES_ROOT, "Nasservb")
        Call CreateKey(HKEY_CLASSES_ROOT, "Nasservb.ir3")
        Call CreateKey(HKEY_CLASSES_ROOT, "Nasservb.irp")
        Call WriteString(HKEY_CLASSES_ROOT, "Nasservb.irp", "", "IR3Project")
        Call CreateKey(HKEY_CLASSES_ROOT, "Nasservb.irp\DefaultIcon")
        Call WriteString(HKEY_CLASSES_ROOT, "Nasservb.irp\DefaultIcon", "", FG & ",0")
        Call CreateKey(HKEY_CLASSES_ROOT, "Nasservb.irp\Shell")
        Call CreateKey(HKEY_CLASSES_ROOT, "Nasservb.irp\Shell\Open")
        Call CreateKey(HKEY_CLASSES_ROOT, "Nasservb.irp\Shell\Open\Command")
        Call WriteString(HKEY_CLASSES_ROOT, "Nasservb.irp\Shell\Open\Command", "", FG & ",%1")
        Call WriteString(HKEY_CURRENT_USER, "Softwar\Microsoft\Windows\CurrentVesion\Explorer\FileExts\.irp", "", FG & ",%1")
        Call WriteString(HKEY_CLASSES_ROOT, "Nasservb.ir3", "", "IRAN Video File")
        Call CreateKey(HKEY_CLASSES_ROOT, "Nasservb.ir3\DefaultIcon")
        Call WriteString(HKEY_CLASSES_ROOT, "Nasservb.ir3\DefaultIcon", "", FG & ",0")
        Call CreateKey(HKEY_CLASSES_ROOT, "Nasservb.ir3\Shell")
        Call CreateKey(HKEY_CLASSES_ROOT, "Nasservb.ir3\Shell\Open")
        Call CreateKey(HKEY_CLASSES_ROOT, "Nasservb.ir3\Shell\Open\Command")
        Call WriteString(HKEY_CLASSES_ROOT, "Nasservb.ir3\Shell\Open\Command", "", FG & ",%1")
        Call WriteString(HKEY_CURRENT_USER, "Softwar\Microsoft\Windows\CurrentVesion\Explorer\FileExts\.ir3", "", FG & ",%1")
4       Command1 = Command
        If Command1 = "" Then Exit Sub
        If LCase(Mid(Command1, Len(Command1) - 3, 3)) = "ir3" Then
            Player.Show
            gfAbort = True
            Unload Me
        ElseIf LCase(Mid(Command1, Len(Command1) - 3, 3)) = "irp" Then
            Command1 = Mid(Command1, 2, Len(Command1) - 2)
        End If
End Sub

Private Sub Timer1_Timer()
On Error GoTo 4
        f = f + 1
        If f > 2 Then
            If Command1 <> "" Then Main.IrpFile = Command1
            Main.Show
            Unload Me
        End If
4  End Sub
