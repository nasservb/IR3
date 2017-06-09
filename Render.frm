VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmRender 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Render"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8820
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Render.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Height          =   495
      Left            =   1080
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   37
      TabIndex        =   9
      Top             =   3240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   1800
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   8
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   255
      Left            =   3600
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   6
      Top             =   3360
      Visible         =   0   'False
      Width           =   255
   End
   Begin Project1.Buttonl Buttonl2 
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   3240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Background"
   End
   Begin Project1.Buttonl Buttonl1 
      Height          =   375
      Left            =   7200
      TabIndex        =   2
      Top             =   3240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Cancel"
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Progress"
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   8535
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   1800
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808080&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   7455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808080&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   7335
      End
   End
End
Attribute VB_Name = "frmRender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Buttonl1_Click()
        Can = True
End Sub

Private Sub Buttonl2_Click()
        Me.WindowState = 1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Buttonl1.Refrash
        Buttonl2.Refrash
End Sub

