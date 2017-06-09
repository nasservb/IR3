VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{51D6BC61-592D-4416-A732-87F39677E763}#1.0#0"; "Slider.ocx"
Begin VB.Form Player 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Player"
   ClientHeight    =   4560
   ClientLeft      =   4620
   ClientTop       =   1545
   ClientWidth     =   6210
   FillStyle       =   0  'Solid
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   6210
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   360
      Top             =   1920
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5295
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   4095
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5415
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   840
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "prview"
      Height          =   1335
      Left            =   5520
      TabIndex        =   0
      Top             =   2400
      Visible         =   0   'False
      Width           =   1095
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00404040&
         Height          =   960
         Left            =   0
         ScaleHeight     =   60
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   60
         TabIndex        =   1
         Top             =   240
         Width           =   960
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   620
      Left            =   0
      TabIndex        =   3
      Top             =   3960
      Width           =   5415
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   160
         Left            =   0
         ScaleHeight     =   165
         ScaleWidth      =   5415
         TabIndex        =   10
         Top             =   0
         Width           =   5415
         Begin Slider.cpvSlider HScroll1 
            Height          =   165
            Left            =   0
            Top             =   0
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   291
            BackColor       =   8421504
            SliderIcon      =   "Form5.frx":0ECA
            Orientation     =   0
            RailPicture     =   "Form5.frx":112C
            RailStyle       =   99
            Max             =   100
         End
      End
      Begin Project1.Buttonl Buttonl5 
         Height          =   495
         Left            =   0
         TabIndex        =   8
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         Caption         =   "Main"
      End
      Begin Project1.Buttonl Buttonl1 
         Height          =   495
         Left            =   2040
         TabIndex        =   7
         Top             =   120
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   873
         Caption         =   "Pause"
      End
      Begin Project1.Buttonl Buttonl2 
         Height          =   495
         Left            =   2760
         TabIndex        =   6
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         Caption         =   "Play"
      End
      Begin Project1.Buttonl Buttonl3 
         Height          =   495
         Left            =   3360
         TabIndex        =   5
         Top             =   120
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   873
         Caption         =   "Stop"
      End
      Begin Project1.Buttonl Buttonl4 
         Height          =   495
         Left            =   4680
         TabIndex        =   4
         Top             =   120
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   873
         Caption         =   "Load"
      End
      Begin VB.Image Image17 
         Height          =   255
         Left            =   1440
         MouseIcon       =   "Form5.frx":1148
         MousePointer    =   99  'Custom
         Picture         =   "Form5.frx":1452
         Stretch         =   -1  'True
         Top             =   240
         Width           =   255
      End
      Begin VB.Image Image2 
         Height          =   495
         Left            =   0
         Picture         =   "Form5.frx":1786
         Stretch         =   -1  'True
         Top             =   120
         Width           =   5415
      End
   End
   Begin VB.Image Image3 
      Height          =   195
      Left            =   5880
      Picture         =   "Form5.frx":1AD8
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   3960
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   5340
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   100
      mute            =   0   'False
      uiMode          =   "none"
      stretchToFit    =   -1  'True
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   0   'False
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   9419
      _cy             =   6985
   End
End
Attribute VB_Name = "Player"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i%
Dim g%, H%, kk&, FRm1&, Fileg$, avi As Boolean
'-------------------------
Private Sub Buttonl1_Click()
On Error GoTo 4
        Timer1.Enabled = False
        Model = Puase
        If avi = True Then WindowsMediaPlayer1.Controls.pause
        Me.Caption = "Puased Video"
4 End Sub

Private Sub Buttonl1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Buttonl2.Refrash
        Buttonl3.Refrash
End Sub

Private Sub Buttonl2_Click()
On Error GoTo 4
        Dim y6%
        If (Model = Readay) Or (Model = stopv) Or (Model = EndFilm) Then
            If avi = True Or WindowsMediaPlayer1.Visible = True Then
                WindowsMediaPlayer1.Controls.play
                Timer1.Enabled = True
                avi = True
                Me.Caption = "Playing...": Can = False: Model = play
                Exit Sub
            End If
            Me.Caption = "Playing...": Can = False: Model = play
            Close
            If CommonDialog1.filename = "" Then GoTo 9
            Open CommonDialog1.filename For Random Access Read As #2 Len = 2
                    Get #2, , y6
                    D1 = CInt(right(CStr(y6), 4))
                    Get #2, , y6
                    D2 = CInt(right(CStr(y6), 4))
                    Get #2, , y6
            HScroll1.Max = y6
            Get #2, , y6
            If y6 <> -7000 Then
                MsgBox "The File Format Is  Unknow!", vbInformation, "IranVideo"
                Me.Caption = "Player"
                Exit Sub
            End If
    
            Picture1.Width = (D1 * 15) + 4
            Picture1.Height = (D2 * 15) + 4
                WindowsMediaPlayer1.Visible = False
                Image1.Visible = True
                Player.HScroll1.Value = 0
                Timer1.Enabled = True
        ElseIf Model = Puase Then
            Me.Caption = "Playing..."
            Timer1.Enabled = True
            If avi = True Then WindowsMediaPlayer1.Controls.play
        ElseIf Model = NoFile Then
            MsgBox "The file is unknow." & vbCrLf & " Plase click on loadfile and select now file"
        ElseIf Model = NoFile Then
9           MsgBox "no file For Paly." & vbCrLf & " Plase click on loadfile and select a file"
        End If
4 End Sub

Private Sub Buttonl2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Buttonl1.Refrash
        Buttonl3.Refrash
End Sub

Private Sub Buttonl3_Click()
On Error GoTo 4
        Can = True
        If avi = True Then WindowsMediaPlayer1.Controls.stop
        Model = stopv
4 End Sub

Private Sub Buttonl3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Buttonl1.Refrash
        Buttonl2.Refrash
End Sub

Private Sub Buttonl4_Click()
On Error GoTo 4
        CommonDialog1.DialogTitle = "Select IranVideo File..."
        CommonDialog1.Filter = "all IranVideo File |*.ir3|Avi File's|*.avi"
        CommonDialog1.Action = 1
        If CommonDialog1.filename = "" Then Exit Sub
        Model = Readay
        If LCase(right(CommonDialog1.filename, 3)) = "avi" Then
            WindowsMediaPlayer1.Visible = True
            WindowsMediaPlayer1.URL = CommonDialog1.filename
            WindowsMediaPlayer1.Controls.play
            Me.Caption = "Playing..."
            Model = play
            avi = True
            Can = False
            Timer1.Enabled = True
            Image1.Visible = False
        Else
            Buttonl2_Click
        End If
4 End Sub

Private Sub Buttonl5_Click()
        gfAbort = False
        Main.Show
End Sub

Private Sub Form_Load()
On Error GoTo 4
        Dim bb$
        Model = NoFile
        bb = Command1
        If bb = "" Then Exit Sub
        If LCase(right(bb, 3)) = "ir3" Then
            CommonDialog1.filename = bb: Model = Readay
            Buttonl2_Click
            Exit Sub
        ElseIf LCase(right(bb, 3)) = "avi" Then
            WindowsMediaPlayer1.settings.autoStart = True
            WindowsMediaPlayer1.Visible = True
            WindowsMediaPlayer1.URL = bb
            WindowsMediaPlayer1.Controls.play
            HScroll1.Max = WindowsMediaPlayer1.Controls.currentItem.duration
            Me.Caption = "Playing..."
            Model = play
            avi = True
            Timer1.Enabled = True
        Else
            bb = right(bb, Len(bb) - 1)
            bb = left(bb, Len(bb) - 1)
            If bb <> "" And LCase(right(bb, 3)) <> "ir3" Then MsgBox "File Is Unknow!"
            If bb <> "" And LCase(right(bb, 3)) = "ir3" Then
                WindowsMediaPlayer1.Controls.stop
                WindowsMediaPlayer1.Visible = False
                CommonDialog1.filename = bb: Model = Readay
                Buttonl2_Click
            End If
        End If
4 End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Buttonl1.Refrash
        Buttonl2.Refrash
        Buttonl3.Refrash
        Buttonl4.Refrash
        Buttonl5.Refrash
End Sub
Private Sub Form_Resize()
On Error GoTo 4
        If Me.WindowState = 1 Then Exit Sub
        Frame3.Width = Me.Width
        Frame3.top = Me.Height - 1020
        Image2.Width = Me.Width
        Buttonl1.left = (Me.Width \ 2) - 950
        Buttonl3.left = (Me.Width \ 2) + 300
        Buttonl2.left = (Me.Width \ 2) - 300
        Buttonl4.left = Me.Width - 900
        Frame2.Width = Me.Width
        Image1.Width = Me.Width
        Image1.Height = Me.Height - 1020
        Frame2.Height = Abs(Me.Height - 1020)
        WindowsMediaPlayer1.Height = Me.Height - 1020
        WindowsMediaPlayer1.Width = Me.Width
        Picture2.Width = Me.Width - 100
        Picture2.PaintPicture Image3, 0, 0, Me.Width, 160
        Set Picture2.Picture = Picture2.Image
        Set HScroll1.RailPicture = Picture2.Picture
        If Me.Width < 3945 Then Me.Width = 3945
        If Me.Height < 3300 Then Me.Height = 3300
        
4 End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error GoTo 4
        If gfAbort = True Then End
        Buttonl3_Click
        Timer1_Timer
4 End Sub


Private Sub HScroll1_MouseDown(Shift As Integer)
        If avi = True Then WindowsMediaPlayer1.Controls.currentPosition = HScroll1.Value

End Sub

Private Sub HScroll1_MouseUp(Shift As Integer)
        If avi = True Then WindowsMediaPlayer1.Controls.currentPosition = HScroll1.Value
        If avi = False And right(CommonDialog1.filename, 3) = "ir3" And _
        HScroll1.Value < HScroll1.Max And HScroll1.Value > 0 Then
            Timer1.Enabled = False
            Me.MousePointer = 11
            SeekIR3 CommonDialog1.filename, HScroll1.Value
        End If

End Sub

Private Sub Image17_Click()
        About.Show 1
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Buttonl1.Refrash
        Buttonl2.Refrash
        Buttonl3.Refrash
        Buttonl4.Refrash
        Buttonl5.Refrash
End Sub


Private Sub Timer1_Timer()
On Error GoTo 4
        If Can = True Then
            If avi = True Then
                WindowsMediaPlayer1.Controls.stop
                avi = False
            End If
            Me.Caption = "Play Canceled!"
            Model = stopv
            Close
            Timer1.Enabled = False
        End If
        If avi = True Then
            Image1.Visible = False
            If Fix(WindowsMediaPlayer1.Controls.currentItem.duration) <> 0 Then _
                 HScroll1.Max = Fix(WindowsMediaPlayer1.Controls.currentItem.duration)
            HScroll1.Value = Fix(WindowsMediaPlayer1.Controls.currentPosition)
            Me.Caption = "Playing ... " & WindowsMediaPlayer1.Controls.currentPositionString _
                        & "/" & WindowsMediaPlayer1.currentMedia.durationString
            If Fix(WindowsMediaPlayer1.Controls.currentItem.duration) < 1 Then HScroll1.Max = 20
            If Abs(HScroll1.Value - HScroll1.Max) < 2 Then Can = True
        Else
            Image1.Visible = True
            LoadIr3
        End If
4 End Sub
