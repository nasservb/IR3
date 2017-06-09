VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Explorer 
   BackColor       =   &H00808080&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Explorer"
   ClientHeight    =   5520
   ClientLeft      =   1035
   ClientTop       =   900
   ClientWidth     =   5730
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      BackColor       =   &H00808080&
      Caption         =   "----------------FX Seeting"
      ForeColor       =   &H00000000&
      Height          =   3375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   5535
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         Height          =   615
         Left            =   120
         ScaleHeight     =   37
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   61
         TabIndex        =   30
         Top             =   3000
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.PictureBox PicFX 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   1575
         Left            =   360
         ScaleHeight     =   101
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   125
         TabIndex        =   23
         Top             =   2880
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00808080&
         Caption         =   "Setting------- Preview"
         Height          =   2895
         Left            =   2160
         TabIndex        =   6
         Top             =   240
         Width           =   3255
         Begin VB.Frame Frame9 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            Caption         =   "Frame9"
            Height          =   495
            Left            =   120
            TabIndex        =   27
            Top             =   2280
            Visible         =   0   'False
            Width           =   3015
            Begin VB.Label Lblcolor 
               BackColor       =   &H80000008&
               Height          =   255
               Left            =   960
               TabIndex        =   29
               Top             =   240
               Width           =   2055
            End
            Begin VB.Label Label3 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Color:"
               Height          =   255
               Left            =   0
               TabIndex        =   28
               Top             =   0
               Width           =   975
            End
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            Caption         =   "Frame8"
            Height          =   495
            Left            =   120
            TabIndex        =   24
            Top             =   2280
            Visible         =   0   'False
            Width           =   3015
            Begin VB.HScrollBar HScroll2 
               Height          =   255
               Left            =   0
               Max             =   300
               TabIndex        =   25
               Top             =   240
               Value           =   150
               Width           =   3015
            End
            Begin VB.Label Label6 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Value"
               Height          =   255
               Left            =   0
               TabIndex        =   26
               Top             =   0
               Width           =   975
            End
         End
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   2400
            Visible         =   0   'False
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   0
            Scrolling       =   1
         End
         Begin Project1.Buttonl cmdCancel 
            Height          =   255
            Left            =   2400
            TabIndex        =   32
            Top             =   2400
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            Caption         =   "Cancel"
         End
         Begin VB.Image Image2 
            Height          =   1935
            Left            =   480
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.ListBox List2 
         BackColor       =   &H00FFFFFF&
         Height          =   2595
         ItemData        =   "browser.frx":0000
         Left            =   120
         List            =   "browser.frx":000D
         TabIndex        =   4
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808080&
         Caption         =   "Effect List:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Caption         =   "Frame7"
      Height          =   1695
      Left            =   120
      TabIndex        =   13
      Top             =   3720
      Width           =   5535
      Begin VB.Frame Frame1 
         BackColor       =   &H00808080&
         Caption         =   "Group Add"
         Height          =   855
         Left            =   2040
         TabIndex        =   19
         Top             =   0
         Width           =   3375
         Begin VB.TextBox Text1 
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   840
            TabIndex        =   20
            Top             =   360
            Width           =   975
         End
         Begin Project1.Buttonl CmdAdd 
            Height          =   375
            Left            =   1920
            TabIndex        =   21
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Caption         =   "Add"
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "Count"
            Height          =   255
            Left            =   360
            TabIndex        =   22
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00808080&
         Caption         =   "Preview"
         Height          =   1695
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   1935
         Begin VB.Image Image1 
            Height          =   1335
            Left            =   120
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00808080&
         Height          =   735
         Left            =   2040
         TabIndex        =   14
         Top             =   960
         Width           =   3375
         Begin Project1.Buttonl Cmd 
            Height          =   375
            Index           =   1
            Left            =   2280
            TabIndex        =   15
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            Caption         =   "Add"
         End
         Begin Project1.Buttonl Cmd 
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            Caption         =   "Set FX"
         End
         Begin Project1.Buttonl Cmd 
            Height          =   375
            Index           =   0
            Left            =   1200
            TabIndex        =   17
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            Caption         =   "Add All"
         End
      End
   End
   Begin MSComctlLib.ImageList imgMain 
      Left            =   4680
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "browser.frx":0033
            Key             =   "mycomputer"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "browser.frx":27E5
            Key             =   "genericfile"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "browser.frx":32AF
            Key             =   "removabledrive"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "browser.frx":5A61
            Key             =   "mydocs"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "browser.frx":8213
            Key             =   "cdrom"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "browser.frx":A9C5
            Key             =   "closedfolder"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "browser.frx":D177
            Key             =   "desktop"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "browser.frx":F929
            Key             =   "openfolder"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "browser.frx":120DB
            Key             =   "unknowndrive"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "browser.frx":1488D
            Key             =   "floppydrive"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "browser.frx":1703F
            Key             =   "harddrive"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "browser.frx":197F1
            Key             =   "netdrive"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picDefault 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   4200
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picBuff 
      AutoRedraw      =   -1  'True
      Height          =   540
      Left            =   3720
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   540
   End
   Begin MSComctlLib.ImageList Imgfiles 
      Left            =   3240
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   5040
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00808080&
      Caption         =   "FX"
      Height          =   255
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   615
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00808080&
      Caption         =   "Drive"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00808080&
      Caption         =   "Directory"
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5535
      Begin MSComctlLib.TreeView TvFolders 
         Height          =   3015
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   5318
         _Version        =   393217
         Style           =   7
         ImageList       =   "imgMain"
         Appearance      =   1
      End
      Begin MSComctlLib.ListView LvFiles 
         Height          =   3015
         Left            =   2760
         TabIndex        =   8
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   5318
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Label Lbltemp 
      Height          =   255
      Left            =   1920
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "Explorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociatedIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long


Private Sub Buttonl1_Click()

End Sub

Private Sub cmd_Click(Index As Integer)
On Error GoTo 4
        Dim s%
        Dim sPath As String, f$
        Select Case Index
            Case 0
                sPath = IIf(right(TvFolders.SelectedItem.Key, 1) = "\", TvFolders.SelectedItem.Key, TvFolders.SelectedItem.Key & "\")
                For s = 1 To LvFiles.ListItems.Count
                    f = LCase(right(LvFiles.ListItems(s).Text, 3))
                    If f = "bmp" Or f = "jpg" Then
                        TimeLine.Ruler1.AddPic sPath & LvFiles.ListItems(s).Text
                    End If
                Next
            Case 1 'Add Picture
                If Image1.Picture <> Empty Then
                    If Option2.Value = True Then
                        If FSO.FolderExists(App.Path & "\fx") = False Then FSO.CreateFolder (App.Path & "\fx")
                        cunt = cunt + 1
                        SavePicture PicFX.Picture, App.Path & "\fx\" & cunt & ".bmp"
                        TimeLine.Ruler1.AddPic App.Path & "\fx\" & cunt & ".bmp"
                    Else
                        TimeLine.Ruler1.AddPic Lbltemp.Caption
                    End If
                End If
            Case 2
                        Option2_Click
                        Option2.Value = True
        End Select
        If Index <> 2 Then
            Monitor.Timer1.Enabled = False
            Saved = True
            Main.Caption = "IranVideo 5.3*"
        End If
        
4  End Sub

Private Sub CmdAdd_Click()
        On Error GoTo 4
        If Val(Text1.Text) = 0 Then Exit Sub
        If Val(Text1.Text) > LvFiles.ListItems.Count Then Text1.Text = LvFiles.ListItems.Count
    Dim i%, f$
    Dim sPath As String
        Saved = True
        Main.Caption = "IranVideo 5.3*"
Monitor.Timer1.Enabled = False
    sPath = IIf(right(TvFolders.SelectedItem.Key, 1) = "\", TvFolders.SelectedItem.Key, TvFolders.SelectedItem.Key & "\")
    For i = 1 To Val(Text1.Text)
        f = LCase(right(LvFiles.ListItems(i).Text, 3))
        If f = "bmp" Or f = "jpg" Then
            TimeLine.Ruler1.AddPic sPath & LvFiles.ListItems(i).Text
        End If
    Next
4
End Sub

Private Sub cmdCancel_Click()
CanFX = True
End Sub

Private Sub File1_PathChange()
  On Error Resume Next
  Dim bRet As Boolean
  DoEvents
  LvFiles.ListItems.Clear
  Set LvFiles.Icons = Nothing
  Set LvFiles.SmallIcons = Nothing
  Imgfiles.ListImages.Clear
  MousePointer = vbHourglass
  'TalkHog "Retrieving files in folder: " & dirList.Path
  Dim sPath As String, W As Long
  sPath = IIf(right(File1.Path, 1) = "\", File1.Path, File1.Path & "\")
  Dim imgT As ListImage, i As Integer, hIcon
  For i = 0 To File1.ListCount - 1
    W = 1
    hIcon = ExtractAssociatedIcon(0, sPath & File1.List(i), W)
    If IsNull(hIcon) Then
                picBuff.Picture = picDefault.Picture
    Else
                Set picBuff.Picture = Nothing
                DoEvents
                DrawIcon picBuff.hDC, 0, 0, hIcon
                DoEvents
                picBuff.Picture = picBuff.Image
                DoEvents
    End If
    Set imgT = Imgfiles.ListImages.Add(, , picBuff.Picture)
  Next
  LvFiles.Icons = Imgfiles
  LvFiles.SmallIcons = Imgfiles
  For i = 0 To File1.ListCount - 1
        LvFiles.ListItems.Add , , File1.List(i), i + 1, i + 1
  Next
  '  TalkHog
  MousePointer = vbDefault
End Sub

Private Sub Form_Load()
        Me.left = 0
        Me.top = 0
        Me.Width = 5895
        LoadDrive
End Sub

Private Sub Form_Resize()
On Error GoTo 4
        Frame4.Width = Me.Width - 330
        TvFolders.Width = (Frame4.Width \ 2) - 150
        LvFiles.Width = (Frame4.Width \ 2) - 160
        LvFiles.left = (Frame4.Width \ 2) + 50
        Frame7.top = Me.Height - 2190
        Frame4.Height = Me.Height - 2535
        TvFolders.Height = Frame4.Height - 360
        LvFiles.Height = Frame4.Height - 360
        If Me.Width < 5655 Then Me.Width = 5655
        If Me.Height < 4365 Then Me.Height = 4365
4 End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        CmdAdd.Refrash
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 4
        Dim i%
        For i = 0 To Cmd.Count - 1
            Cmd(i).Refrash
        Next
4 End Sub


Private Sub Frame6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdCancel.Refrash
End Sub

Private Sub HScroll2_Change()
On Error GoTo 4
    Label6.Caption = "Value:" & HScroll2.Value & "%"
    Frame8.Visible = True
    Frame9.Visible = False
    Me.MousePointer = 11
    PicFX.Picture = Image1.Picture
    FX.Brightnes PicFX, HScroll2.Value
    Image2.Picture = PicFX.Picture
    Me.MousePointer = 0
4 End Sub
Private Sub HScroll2_Scroll()
        Label6.Caption = "Value:" & HScroll2.Value & "%"
End Sub

Private Sub Lblcolor_Click()
On Error GoTo 4
        Me.MousePointer = 11
        Main.CommonDialog1.ShowColor
        Lblcolor.BackColor = Main.CommonDialog1.Color
        PicFX.Picture = Image1.Picture
        ColorPalette Lblcolor.BackColor, PicFX
        Image2.Picture = PicFX.Picture
        Me.MousePointer = 0
4 End Sub

Private Sub List2_Click()
On Error GoTo 4
        Monitor.Timer1.Enabled = False
        cmdCancel.Visible = True
        ProgressBar1.Visible = True
         Select Case List2.ListIndex
                Case 0
            Frame8.Visible = True
            Frame9.Visible = False
            Image2.Picture = Image1.Picture
                Case 1
            Frame9.Visible = True
            Frame8.Visible = False
            Image2.Picture = Image1.Picture
                Case 2
            PicFX.Picture = Image1.Picture
            PicFX.PaintPicture PicFX.Picture, 0, 0, , , , , , , vbDstInvert
            Set PicFX.Picture = PicFX.Image
            Image2.Picture = PicFX.Picture
            Frame9.Visible = False
            Frame8.Visible = False
            cmdCancel.Visible = False
            ProgressBar1.Visible = False
        End Select
4 End Sub

Private Sub LvFiles_Click()
        On Error GoTo 4
        Dim s$
        s = IIf(Len(TvFolders.SelectedItem.Key) = 3, TvFolders.SelectedItem.Key, TvFolders.SelectedItem.Key & "\")
        s = s & LvFiles.SelectedItem.Text
        Image1.Picture = LoadPicture(s)
        Lbltemp.Caption = s
4
End Sub

Private Sub Option1_Click()
        Frame4.Visible = True
        Frame5.Visible = False
End Sub

Private Sub Option2_Click()
        Frame4.Visible = False
        Frame5.Visible = True
End Sub
Private Sub LoadDrive()
On Error GoTo 4
  Dim ADrive As Drive
  Dim Icon As String
  Dim Name As String
  Dim AFolder As Folder
  Dim DriveFolders As Folder

  ' Show the drives w/correct icons and names
  For Each ADrive In FSO.Drives
    If ADrive.DriveType = CDRom Then
      Icon = "cdrom"
      If ADrive.IsReady Then Name = ADrive.VolumeName Else Name = "CD-ROM Drive"
    ElseIf ADrive.DriveType = Fixed Then
      Icon = "harddrive"
      If ADrive.IsReady Then Name = ADrive.VolumeName Else Name = "Hard Drive"
    ElseIf ADrive.DriveType = Remote Then
      Icon = "netdrive"
      If ADrive.IsReady Then Name = ADrive.ShareName Else Name = "Network Drive"
    ElseIf ADrive.DriveType = Removable Then
      If ADrive.DriveLetter = "A" Or ADrive.DriveLetter = "B" Then Icon = "floppydrive" Else Icon = "removabledrive"
      If ADrive.IsReady Then
        Name = ADrive.VolumeName
      Else
        If ADrive.DriveLetter = "A" Or ADrive.DriveLetter = "B" Then Name = "Floppy Drive" Else Name = "Removable Drive"
      End If
    Else
      Icon = "unknowndrive"
      If ADrive.IsReady Then Name = ADrive.VolumeName Else Name = "Unknown"
    End If
    
    'Add the drives node to the root tree
    'The key is the drives path
    TvFolders.Nodes.Add , 0, ADrive.Path, Name & " (" & UCase(ADrive.DriveLetter) & ":)", Icon
    
    
    'If the drive is available grab the drives root directories
    'We do this before the user expands the drive the the plus-minus box shows up right.
    If ADrive.IsReady Then
      Set DriveFolders = FSO.GetFolder(ADrive.RootFolder)
      For Each AFolder In DriveFolders.SubFolders
        'Add the folder to the tree, with the drive as it's parent
        'The key is the full path to the folder
        TvFolders.Nodes.Add ADrive.Path, 4, AFolder.Path, AFolder.Name, "closedfolder"
      Next
    End If
  Next
4
End Sub

Private Sub ProgressBar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdCancel.Refrash
End Sub

Private Sub TvFolders_Collapse(ByVal Node As MSComctlLib.Node)
  If Node.Image = "openfolder" Then Node.Image = "closedfolder"
End Sub

Private Sub TvFolders_Expand(ByVal Node As MSComctlLib.Node)
 On Error Resume Next
  Dim SubSubFolder As Folder
  Dim SubFolder As Folder
  Dim AFolder As Folder
  If Node.Image = "closedfolder" Then Node.Image = "openfolder"
   Set AFolder = FSO.GetFolder(Node.Key & "\") 'Add the backslash :-)
  For Each SubFolder In AFolder.SubFolders
    For Each SubSubFolder In SubFolder.SubFolders
      'Add children to the expanded nodes children
      TvFolders.Nodes.Add SubFolder.Path, 4, SubSubFolder.Path, SubSubFolder.Name, "closedfolder"
    Next
  Next
End Sub

Private Sub TvFolders_NodeClick(ByVal Node As MSComctlLib.Node)
  On Error GoTo ErrorHandler
  Dim AFolder As Folder
  Dim AFile As File
  
   LvFiles.ListItems.Clear
  Set AFolder = FSO.GetFolder(Node.Key & "\") 'Add the backslash :-)
   File1.Path = IIf(Len(AFolder.Path) = 3, AFolder.Path, AFolder.Path & "\")

Exit Sub
ErrorHandler:
    If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, "Error Number: " & Err.Number

End Sub
