VERSION 5.00
Begin VB.Form export 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export Setting"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7050
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   255
      Left            =   1560
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   23
      Top             =   5640
      Visible         =   0   'False
      Width           =   375
   End
   Begin Project1.Buttonl cmdExport 
      Height          =   375
      Left            =   4440
      TabIndex        =   9
      Top             =   5520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Export"
   End
   Begin Project1.Buttonl cmdCancel 
      Height          =   375
      Left            =   5760
      TabIndex        =   10
      Top             =   5520
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Cancel"
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "Output File"
      Height          =   1455
      Left            =   120
      TabIndex        =   20
      Top             =   240
      Width           =   6735
      Begin Project1.Buttonl Brows 
         Height          =   315
         Left            =   5640
         TabIndex        =   2
         Top             =   840
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         Caption         =   "Browse"
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Text            =   "Video"
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000018&
         Caption         =   "A file name cannot using characters:  \/:*?""""<>|"
         Height          =   255
         Left            =   3000
         TabIndex        =   24
         Top             =   360
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.Label Label13 
         BackColor       =   &H00808080&
         Caption         =   "Directory:"
         Height          =   255
         Left            =   480
         TabIndex        =   22
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   840
         Width           =   4215
      End
      Begin VB.Label Label11 
         BackColor       =   &H00808080&
         Caption         =   "File Name:"
         Height          =   255
         Left            =   480
         TabIndex        =   21
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Output Format"
      Height          =   3495
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   6735
      Begin VB.ComboBox Combo5 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         ItemData        =   "export.frx":0000
         Left            =   1800
         List            =   "export.frx":0016
         TabIndex        =   8
         Text            =   "16"
         Top             =   2640
         Width           =   2655
      End
      Begin VB.ComboBox Combo4 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         ItemData        =   "export.frx":0037
         Left            =   1800
         List            =   "export.frx":0059
         TabIndex        =   7
         Text            =   "1"
         ToolTipText     =   "1-for picture Album"
         Top             =   2040
         Width           =   2655
      End
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         ItemData        =   "export.frx":008B
         Left            =   3960
         List            =   "export.frx":00B9
         TabIndex        =   6
         Text            =   "320"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         ItemData        =   "export.frx":0102
         Left            =   1800
         List            =   "export.frx":0130
         TabIndex        =   5
         Text            =   "480"
         Top             =   1440
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         ItemData        =   "export.frx":017B
         Left            =   1800
         List            =   "export.frx":0188
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   2535
      End
      Begin Project1.Buttonl Chose 
         Height          =   315
         Left            =   4440
         TabIndex        =   4
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "Chose Codec"
      End
      Begin VB.Label Label8 
         BackColor       =   &H00808080&
         Caption         =   "Video Query"
         Height          =   255
         Left            =   720
         TabIndex        =   19
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808080&
         Caption         =   "frame/sec"
         Height          =   255
         Left            =   4560
         TabIndex        =   18
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label6 
         BackColor       =   &H00808080&
         Caption         =   "Frame Rate"
         Height          =   255
         Left            =   720
         TabIndex        =   17
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H00808080&
         Caption         =   " Width"
         Height          =   255
         Left            =   1800
         TabIndex        =   16
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label4 
         BackColor       =   &H00808080&
         Caption         =   "Height"
         Height          =   255
         Left            =   3960
         TabIndex        =   15
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00808080&
         Caption         =   "Size"
         Height          =   255
         Left            =   1200
         TabIndex        =   14
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808080&
         Caption         =   "X"
         Height          =   255
         Left            =   3720
         TabIndex        =   13
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808080&
         Caption         =   "FileFormat:"
         Height          =   255
         Left            =   840
         TabIndex        =   12
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   960
      Stretch         =   -1  'True
      Top             =   5400
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "export"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ps As Long, res As Long
Dim Opts             As AVI_COMPRESS_OPTIONS
Dim pOpts            As Long
Dim szOutputAVIFile  As String
Dim pFile            As Long
Dim strhdr           As AVI_STREAM_INFO
Dim ppath            As String
Dim inform           As RenderInfo
Dim mc               As Boolean
Dim p$
Private Sub Brows_Click()
On Error GoTo 4
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    szTitle = "Directory:"
    With tBrowseInfo
        .hWndOwner = Me.hwnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Trim(sBuffer) 'left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        Label12.Caption = left(sBuffer, Len(sBuffer) - 1)
    End If
4 End Sub

Private Sub Chose_Click()
        On Error GoTo 4
        TimeLine.Ruler1.SetList inform.List
        If Val(Combo3.Text) = 0 Or Val(Combo2.Text) = 0 Or inform.List.ListCount = 0 Then Exit Sub
        If Combo1.ListIndex <> 0 Then Exit Sub
    ppath = IIf(Len(App.Path) = 3, App.Path, App.Path & "\")
    p = IIf(Len(Label12.Caption) = 3, Label12.Caption, Label12.Caption & "\")
    szOutputAVIFile = p & Text1.Text & Combo1.Text
    If FSO.FileExists(szOutputAVIFile) = True Then FSO.DeleteFile szOutputAVIFile, False
    res = AVIFileOpen(pFile, szOutputAVIFile, OF_WRITE Or OF_CREATE, 0&)
    If (res <> AVIERR_OK) Then Exit Sub
        Set bmp = New cDIB
        Picture1.Width = (Val(Combo2.Text) + 4) * 15
        Picture1.Height = (Val(Combo3.Text) + 4) * 15
    With Picture1
        .Picture = LoadPicture(TimeLine.Ruler1.PicList(0))
        .PaintPicture .Picture, 0, 0, .Width, .Height
        Set .Picture = .Image
    End With
    SavePicture Picture1.Picture, ppath & "7.bmp"
    If bmp.CreateFromFile(ppath & "7.bmp") <> True Then
        MsgBox "Could not load first bitmap file in list!", vbExclamation, App.title
        Kill (ppath & "7.bmp")
        Exit Sub
    End If
    Kill (ppath & "7.bmp")
    With strhdr
        .fccType = mmioStringToFOURCC("vids", 0&)
        .fccHandler = 0&
        .dwScale = 1
        .dwRate = Val(Combo4.Text) '=Secnd Stop On Pic
        .dwSuggestedBufferSize = bmp.SizeImage
        Call SetRect(.rcFrame, 0, 0, Val(Combo3.Text), Val(Combo2.Text)) ' bmp.Height
    End With
    If strhdr.dwRate < 1 Then strhdr.dwRate = 1
    If strhdr.dwRate > 30 Then strhdr.dwRate = 30
    res = AVIFileCreateStream(pFile, ps, strhdr)
    If (res <> AVIERR_OK) Then Exit Sub
    pOpts = VarPtr(Opts)
    res = AVISaveOptions(Me.hwnd, _
                        ICMF_CHOOSE_KEYFRAME Or ICMF_CHOOSE_DATARATE, _
                        1, ps, pOpts)
mc = True
4
End Sub

Private Sub cmdCancel_Click()
        Me.Hide
End Sub

Private Sub cmdExport_Click()
On Error GoTo 4
Dim Result%, Fname$
    If Combo1.ListIndex = 2 And TimeLine.Ruler1.CurentPic(Image1).Picture <> 0 Then
        Picture1.Width = (Val(Combo2.Text) + 4) * 15
        Picture1.Height = (Val(Combo3.Text) + 4) * 15
        Picture1.PaintPicture Image1.Picture, 0, 0, Picture1.Width, Picture1.Height
        Set Picture1.Picture = Picture1.Image
        p = IIf(Len(Label12.Caption) = 3, Label12.Caption, Label12.Caption & "\")
        If FSO.FileExists(p & Text1.Text & Combo1.Text) = True Then
            Result = MsgBox("Your selected file alreday exist's" & vbCrLf & _
            "Do you want to over write this file?", vbYesNoCancel + vbExclamation, Main.Caption)
            If Result = vbYes Then
                SavePicture Picture1.Picture, p & Text1.Text & Combo1.Text
            ElseIf Result = vbNo Then
                SavePicture Picture1.Picture, p & Text1.Text & "0001" & Combo1.Text
            End If
            If Result <> vbCancel Then Shell "explorer.exe" & " " & p, vbNormalFocus
        Else
            SavePicture Picture1.Picture, p & Text1.Text & Combo1.Text
            Shell "explorer.exe" & " " & p, vbNormalFocus
        End If
            Exit Sub
    End If
        TimeLine.Ruler1.SetList inform.List
        If Val(Combo2.Text) = 0 Or _
         Val(Combo3.Text) = 0 Or _
         Val(Combo4.Text) = 0 Or _
        (Combo1.Text) = "" Or _
         Text1.Text = "" Or _
         Label12.Caption = "" Or _
         inform.List.ListCount = 0 Then
2       MsgBox "Plase Complete Output Information or add image", vbInformation, Main.Caption
                Exit Sub
        End If
        If Combo1.ListIndex = 0 And mc = False Then GoTo 2
        Fname = p & Text1.Text & Combo1.Text
        If FSO.FileExists(p & Text1.Text & Combo1.Text) = True Then
            Result = MsgBox("Your selected file alreday exist's" & vbCrLf & _
            "Do you want to over write this file?", vbYesNoCancel + vbExclamation, Main.Caption)
            If Result = vbYes Then
                Fname = p & Text1.Text & Combo1.Text
            ElseIf Result = vbNo Then
                Fname = p & Text1.Text & "0001" & Combo1.Text
            End If
        If Result = vbCancel Then Exit Sub
        End If
        With inform
            .Codec = res
            p = IIf(Len(Label12.Caption) = 3, Label12.Caption, Label12.Caption & "\")
            .filename = Fname
            .FrameRate = Val(Combo4.Text)
            .pOpts = pOpts
            .ps = ps
            .Opts = Opts
            .Query = Val(Combo5.Text)
            .Size_Width = Val(Combo2.Text)
            .Size_Height = Val(Combo3.Text)
            .pFile = pFile
            .RenderMode = Combo1.ListIndex
        End With
        Me.Visible = False
        Main.Hide
        If Combo1.ListIndex = 0 Or Combo1.ListIndex = 1 Then Comperes inform
4
End Sub

Private Sub Combo1_Change()
        t6 = 0
End Sub

Private Sub Combo2_Change()
        t6 = 0
End Sub

Private Sub Combo3_Change()
        t6 = 0
End Sub

Private Sub Form_Load()
Monitor.Timer1.Enabled = False

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        cmdExport.Refrash
        cmdCancel.Refrash
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Chose.Refrash
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Brows.Refrash
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Label9.Visible = False
Select Case KeyAscii
       Case 124, 92, 47, 58, 42, 63, 34, 60, 62
            Label9.Visible = True
            KeyAscii = 0
       Case 24: KeyAscii = 0
End Select
End Sub
