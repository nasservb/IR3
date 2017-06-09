VERSION 5.00
Begin VB.UserControl Buttonl 
   BackColor       =   &H00404040&
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1875
   ScaleHeight     =   510
   ScaleWidth      =   1875
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Button1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   0
      Picture         =   "UserControl11.ctx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1815
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   0
      Picture         =   "UserControl11.ctx":0352
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   0
      Picture         =   "UserControl11.ctx":0654
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "Buttonl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'FIXIT: Use Option Explicit to avoid implicitly creating variables of type Variant         FixIT90210ae-R383-H1984
Const DefultCap = "Button1"
Public Event Click()
Public Event MouseMove(Button%, Shift%, X!, Y!)
Public Event MouseDown(Button%, Shift%, X!, Y!)
Public Event MouseUp(Button%, Shift%, X!, Y!)
'-------------------------
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
	On Error GoTo 4
	Image1.Visible = False
	Image2.Visible = True
	RaiseEvent MouseDown(Button, Shift, X, Y)
4
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
	On Error GoTo 4
	Image2.Visible = False
	Image3.Visible = True
	RaiseEvent MouseUp(Button, Shift, X, Y)
4
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
	On Error GoTo 4
	Image2.Visible = False
	Image3.Visible = True
	RaiseEvent MouseUp(Button, Shift, X, Y)
4
End Sub
Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
	On Error GoTo 4
	Image3.Visible = False
	Image1.Visible = True
	RaiseEvent MouseMove(Button, Shift, X, Y)
4
End Sub

Private Sub Image12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
	On Error GoTo 4
	Call Image1_MouseDown(Button, Shift, X, Y)
4
End Sub
Private Sub Image12_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
	On Error GoTo 4
	Call Image1_MouseUp(Button, Shift, X, Y)
4
End Sub
'--------------------------
Private Sub Image11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
	On Error GoTo 4
	Call Image1_MouseDown(Button, Shift, X, Y)
4
End Sub
Private Sub Image11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
	On Error GoTo 4
	Call Image1_MouseUp(Button, Shift, X, Y)
4
End Sub
'--------------------------
Private Sub Image21_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
	On Error GoTo 4
	Call Image2_MouseUp(Button, Shift, X, Y)
4
End Sub
'--------------------------
Private Sub Image22_Mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)
	On Error GoTo 4
	Call Image2_MouseUp(Button, Shift, X, Y)
4
End Sub
'-------------------------------------------
Private Sub Image32_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
	On Error GoTo 4
	Call Image3_MouseMove(Button, Shift, X, Y)
4
End Sub
'---------------------------
Private Sub Image33_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
	On Error GoTo 4
	Call Image3_MouseMove(Button, Shift, X, Y)
4
End Sub

Private Sub Label1_Click()
	On Error GoTo 4
	RaiseEvent Click
4
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
	On Error GoTo 4
	Image1.Visible = False
	Image2.Visible = True
	RaiseEvent MouseDown(Button, Shift, X, Y)
4
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
	On Error GoTo 4
	Image3.Visible = False
	Image1.Visible = True
	RaiseEvent MouseMove(Button, Shift, X, Y)
4
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
	On Error GoTo 4
	Image1.Visible = False
	Image3.Visible = True
	UserControl.Refresh
	RaiseEvent MouseUp(Button, Shift, X, Y)
4
End Sub

Private Sub UserControl_Initialize()
	UserControl.Refresh
End Sub

Private Sub UserControl_Resize()
	On Error GoTo 4
	Dim W%, H%
	H = UserControl.Height: W = UserControl.Width
	Image1.Width = W: Image1.Height = H
	Image2.Width = W: Image2.Height = H
	Image3.Width = W: Image3.Height = H
	Label1.Width = W: Label1.top = (H \ 2) - (Label1.Height \ 2)
4
End Sub
Public Sub Refrash()
	On Error GoTo 4
	Image2.Visible = False
	Image1.Visible = False
	Image3.Visible = True
4
End Sub
Public Property Get Caption() As String
	On Error GoTo 4
	 Caption = Label1.Caption
4
End Property
'------------------------
Public Property Let Caption(Value$)
	On Error GoTo 4
	Label1.Caption = Value: PropertyChanged "Caption"
4
End Property

Private Sub UserControl_Terminate()
	UserControl.Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
	On Error GoTo 4
	Call PropBag.WriteProperty("Caption", Label1.Caption, DefultCap)
4
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
	On Error GoTo 4
	Label1.Caption = PropBag.ReadProperty("Caption", DefultCap)
4
End Sub
