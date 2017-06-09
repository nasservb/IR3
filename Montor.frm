VERSION 5.00
Begin VB.Form Monitor 
   BackColor       =   &H00808080&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Monitor"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   6240
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2280
      Top             =   1680
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   4800
      Width           =   5655
      Begin Project1.Buttonl Btn 
         Height          =   375
         Index           =   1
         Left            =   3600
         TabIndex        =   2
         Top             =   0
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "Stop"
      End
      Begin Project1.Buttonl Btn 
         Height          =   375
         Index           =   2
         Left            =   2400
         TabIndex        =   3
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "Play"
      End
      Begin Project1.Buttonl Btn 
         Height          =   375
         Index           =   0
         Left            =   1320
         TabIndex        =   4
         Top             =   0
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "Pause"
      End
      Begin VB.Image Image2 
         Height          =   375
         Left            =   -600
         Picture         =   "Montor.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   6255
      End
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   0
      Picture         =   "Montor.frx":0352
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   6615
   End
   Begin VB.Image Image1 
      Height          =   4575
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "Monitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SS%, frem%, offs%
Public pus As Boolean
'-----------------------
Private Sub Btn_Click(Index As Integer)
On Error GoTo 4
        Select Case Index
                Case 0: Timer1.Enabled = False: pus = True
                Case 1
                    TimeLine.Ruler1.MoveFirst
                    TimeLine.Ruler1.CurentPic(Image1).Refresh
                    Timer1.Enabled = False
                    pus = False
                Case 2
                If pus = False Then
                    TimeLine.Ruler1.MoveFirst
                    TimeLine.Ruler1.CurentPic(Image1).Refresh
                End If
                pus = False
                Timer1.Enabled = True
        End Select
4 End Sub

Private Sub Form_Load()
        Me.Width = 6315
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 4
        Dim i%
        For i = 0 To Btn.Count - 1
            Btn(i).Refrash
        Next
4 End Sub

Private Sub Form_Resize()
        On Error GoTo 4
        Label1.Height = Me.Height - 780
        Image1.Height = Label1.Height
        Frame1.top = Me.Height - 780
        Image3.top = Frame1.top
        Image3.Width = Me.Width
        Frame1.left = (Me.Width \ 2) - (Frame1.Width \ 2)
        Label1.Width = Me.Width
        Image1.Width = Label1.Width
        If Me.Width < 3810 Then Me.Width = 3810
        If Me.Height < 2895 Then Me.Height = 2895
4
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Form_MouseMove Button, Shift, X, Y
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub Timer1_Timer()
On Error GoTo 4
If TimeLine.Ruler1.ListCount = 0 Then
    Timer1.Enabled = False
    GoTo 4
End If
offs = offs + 1
If offs = 200 Then
    frem = frem + 1
    offs = 0
    TimeLine.Ruler1.NextImg
    TimeLine.Ruler1.CurentPic(Image1).Refresh
End If
Me.Caption = "Monitor " & TTim(offs, frem * 2) & "/" & TimeLine.Ruler1.FilmTime
        If TimeLine.Ruler1.EndFilm = True Then
            TimeLine.Ruler1.MoveFirst
            frem = 0
            offs = 0
            TimeLine.Ruler1.CurentPic(Image1).Refresh
            Timer1.Enabled = False
        End If
4 End Sub

