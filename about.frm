VERSION 5.00
Begin VB.Form About 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   4620
   ClientLeft      =   405
   ClientTop       =   1890
   ClientWidth     =   5415
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.Buttonl Buttonl1 
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   3960
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   873
      Caption         =   "OK"
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Nasservb@Gmail.com"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2520
      MouseIcon       =   "about.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "www.tcvb.tk"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2520
      MouseIcon       =   "about.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "www.nasservb.blogfa.com"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2520
      MouseIcon       =   "about.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"about.frx":091E
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   2520
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   3855
      Left            =   0
      Picture         =   "about.frx":0A07
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetSystemDirectory Lib "kernel32" Alias _
                        "GetSystemDirectoryA" (ByVal lpBuffer As String, _
                                               ByVal nSize As Long) As Long
Dim f As String * 255
Private Sub Buttonl1_Click()
        Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo 4
        GetSystemDirectory f, 255
4 End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Buttonl1.Refrash
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Buttonl1.Refrash

End Sub

Private Sub Label2_Click()
On Error GoTo 4
        Shell left(f, 11) & "explorer.exe" & " http://www.nasservb.blogfa.com", vbNormalFocus
4 End Sub

Private Sub Label3_Click()
        Shell left(f, 11) & "explorer.exe" & " http://www.tcvb.tk", vbNormalFocus
End Sub

Private Sub Label4_Click()
On Error GoTo 4
        Shell left(f, 11) & "explorer.exe" & " mailto:nasservb@gmail.com", vbNormalFocus

4 End Sub
