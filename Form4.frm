VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form TimeLine 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "TimeLine"
   ClientHeight    =   2535
   ClientLeft      =   2760
   ClientTop       =   5505
   ClientWidth     =   10530
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   10530
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   9960
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   3615
      Begin Project1.Buttonl Cmd 
         Height          =   255
         Index           =   3
         Left            =   2520
         TabIndex        =   2
         Top             =   0
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         Caption         =   "<--"
      End
      Begin Project1.Buttonl Cmd 
         Height          =   255
         Index           =   4
         Left            =   3120
         TabIndex        =   3
         Top             =   0
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         Caption         =   "-->"
      End
      Begin Project1.Buttonl Cmd 
         Height          =   255
         Index           =   2
         Left            =   1680
         TabIndex        =   4
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         Caption         =   "Clear"
      End
      Begin Project1.Buttonl Cmd 
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   5
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         Caption         =   "Del"
      End
      Begin Project1.Buttonl Cmd 
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         Caption         =   "Add"
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   0
         Picture         =   "Form4.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3615
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9000
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Chose Picture"
      Filter          =   "All Picture File |*.bmp;*.jpg;*.jpeg"
   End
   Begin Project1.Ruler Ruler1 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   4471
   End
End
Attribute VB_Name = "TimeLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_Click(Index As Integer)
On Error GoTo 4
Select Case Index
    Case 0 'Add Picture
         CommonDialog1.ShowOpen
        If CommonDialog1.filename <> "" Then _
            Ruler1.AddPic CommonDialog1.filename
    Case 1: Ruler1.DelSelected
    Case 2 'claer time line
        Ruler1.Clear
        Saved = False
        Main.Caption = "IranVideo 5.3"
    Case 3: Ruler1.MoveBack
    Case 4: Ruler1.MoveNext
End Select
t6 = 0
Monitor.Timer1.Enabled = False
If Ruler1.ListCount > 0 Then
    Saved = True
    Main.Caption = "IranVideo 5.3*"
End If
4 End Sub

Private Sub Form_Resize()
        Ruler1.Width = Me.Width
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Ruler1_MuoseMove Button, Shift, X, Y
End Sub

Private Sub Ruler1_Click()
On Error GoTo 4
Monitor.Timer1.Enabled = False
        If Ruler1.CurentPic(Monitor.Image1).Picture <> Empty Then Exit Sub
4 End Sub

Private Sub Ruler1_ListChange()
Me.Caption = "TimeLine " & Ruler1.MoveTime & "/" & Ruler1.FilmTime
End Sub

Private Sub Ruler1_MuoseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 4
        Me.Caption = "TimeLine " & Ruler1.MoveTime & "/" & Ruler1.FilmTime
        Dim i%
        For i = 0 To Cmd.Count - 1
            Cmd(i).Refrash
        Next
4 End Sub
