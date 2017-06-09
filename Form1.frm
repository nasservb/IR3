VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7455
   ClientLeft      =   2805
   ClientTop       =   1890
   ClientWidth     =   9990
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   497
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   666
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin Project1.Buttonl Buttonl1 
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   5160
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "Player"
      End
      Begin VB.ListBox List1 
         BackColor       =   &H8000000A&
         Height          =   2790
         Left            =   6720
         TabIndex        =   18
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "......"
         Height          =   255
         Left            =   8040
         TabIndex        =   17
         Top             =   4920
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   6720
         TabIndex        =   16
         Top             =   4920
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Compile"
         Height          =   375
         Left            =   7800
         TabIndex        =   15
         Top             =   5160
         Width           =   975
      End
      Begin VB.DirListBox Dir1 
         Height          =   765
         Left            =   6720
         TabIndex        =   14
         Top             =   3840
         Width           =   1095
      End
      Begin VB.FileListBox File1 
         Height          =   870
         Left            =   7800
         Pattern         =   "*.bmp;*.jpg"
         TabIndex        =   13
         Top             =   3840
         Width           =   855
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   6720
         TabIndex        =   12
         Top             =   3480
         Width           =   1935
      End
      Begin VB.CommandButton Command4 
         Caption         =   "ReCompile"
         Height          =   375
         Left            =   6720
         TabIndex        =   11
         Top             =   5160
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   480
         TabIndex        =   10
         Text            =   "2"
         Top             =   4680
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form1.frx":0000
         Left            =   3360
         List            =   "Form1.frx":004F
         TabIndex        =   9
         Text            =   "64__64"
         Top             =   5160
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Caption         =   "prview"
         Height          =   1335
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1095
         Begin VB.PictureBox Picture1 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            Height          =   960
            Left            =   0
            ScaleHeight     =   60
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   60
            TabIndex        =   8
            Top             =   240
            Width           =   960
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   1215
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   1095
         Begin VB.PictureBox Picture2 
            AutoRedraw      =   -1  'True
            Height          =   855
            Left            =   0
            ScaleHeight     =   795
            ScaleWidth      =   915
            TabIndex        =   6
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   66
         Left            =   4320
         Top             =   4560
      End
      Begin VB.CommandButton Command3 
         Caption         =   "/\"
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   4680
         Width           =   375
      End
      Begin VB.CommandButton Command5 
         Caption         =   "\/"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   4680
         Width           =   375
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   7560
         TabIndex        =   2
         Text            =   "123"
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Add"
         Height          =   315
         Left            =   6720
         TabIndex        =   1
         Top             =   3120
         Width           =   735
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   8040
         Top             =   5520
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   4215
         Left            =   1320
         Stretch         =   -1  'True
         Top             =   240
         Width           =   5295
      End
      Begin VB.Label Label1 
         Caption         =   "Save Ofset"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   4320
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "out size"
         Height          =   255
         Left            =   3480
         TabIndex        =   21
         Top             =   4920
         Width           =   615
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4560
         TabIndex        =   20
         Top             =   5160
         Width           =   1215
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   19
         Top             =   5160
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BtMP%, Indx_ As String, Can As Boolean
Dim D1%, D2%, jj%, Fj() As Integer, t6%

Private Sub Buttonl1_Click()
	Form1.Show
End Sub

Private Sub Combo1_Change()
	Indx_ = Combo1.Text
    EditCustom
End Sub

Private Sub Command1_Click()
	On Error Resume Next
	CommonDialog1.Filter = "ir3|*.ir3|All file's |*.*"
	CommonDialog1.ShowSave
	Text1.Text = CommonDialog1.FileName
End Sub


Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
	If Button = 1 Then _
	Text2.Text = Val(Text2.Text) + 10

End Sub

Private Sub Command4_Click()
	If Text1.Text <> "" Then _
	Lode Text1.Text, Me
End Sub

Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then _
Text2.Text = Val(Text2.Text) - 10

End Sub

Private Sub Command6_Click()
	If File1.ListCount < Val(Text3.Text) Then Exit Sub
	List1.Clear
	Can = False: Command7.Visible = True
	For i = 0 To Val(Text3.Text)
		If Can = True Then Exit For
		List1.AddItem Dir1.Path + "\" + File1.List(i)
	Next
	Command7.Visible = False
End Sub

Private Sub Command7_Click()
	Can = True
	Command7.Visible = False
End Sub

Private Sub Dir1_Change()
	File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
	On Error Resume Next
	Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
	If Button = 2 Then
		If File1.List(File1.ListIndex) <> "" Then List1.AddItem Dir1.Path + "\" + File1.List(File1.ListIndex)
	ElseIf Button = 4 Then
		Can = False: Command7.Visible = True
		For i = 0 To File1.ListCount - 1
			If Can = True Then Exit For
			List1.AddItem Dir1.Path + "\" + File1.List(i)
		Next
		Command7.Visible = False
	End If
End Sub

Private Sub Form_Load()
	Dim n$
	Can = False
	Text1.Text = Command
	If Text1.Text <> "" And LCase(right(Text2.Text, 3)) = "ir3" Then Command2_Click
	If Text1.Text <> "" And LCase(right(Text2.Text, 3)) <> "ir3" Then MsgBox "File Is Unknow!"
End Sub


Private Sub Form_Unload(Cancel As Integer)
	Can = True
	Close
	Erase Fj
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Buttonl1.Refrash
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
	If Button = 1 Then
		Image1.Picture = LoadPicture(List1.Text)
	ElseIf Button = 2 Then
		List1.RemoveItem List1.ListIndex: t6 = 0
	ElseIf Button = 4 Then
		List1.Clear: t6 = 0
	End If
End Sub
