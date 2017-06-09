VERSION 5.00
Object = "{51D6BC61-592D-4416-A732-87F39677E763}#1.0#0"; "Slider.ocx"
Begin VB.UserControl Ruler 
   BackColor       =   &H00404040&
   ClientHeight    =   2895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   ScaleHeight     =   2895
   ScaleWidth      =   9000
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   3840
      ScaleHeight     =   225
      ScaleWidth      =   5055
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   5055
      Begin Slider.cpvSlider HScroll1 
         Height          =   150
         Left            =   0
         Top             =   0
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   265
         BackColor       =   8421504
         SliderIcon      =   "Ruler.ctx":0000
         Orientation     =   0
         RailPicture     =   "Ruler.ctx":0315
         RailStyle       =   99
         ShowValueTip    =   0   'False
      End
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   5400
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Height          =   1335
      Left            =   5040
      ScaleHeight     =   1275
      ScaleWidth      =   1515
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   2955
      Left            =   0
      TabIndex        =   0
      Top             =   -140
      Width           =   4575
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   -120
         ScaleHeight     =   375
         ScaleWidth      =   4815
         TabIndex        =   1
         Top             =   2205
         Width           =   4815
         Begin VB.Line Line2 
            BorderColor     =   &H0080FF80&
            X1              =   3480
            X2              =   3480
            Y1              =   360
            Y2              =   0
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00808080&
         Caption         =   "Frame2"
         Height          =   735
         Left            =   -120
         MousePointer    =   5  'Size
         TabIndex        =   2
         Top             =   -360
         Width           =   5295
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   -120
         ScaleHeight     =   375
         ScaleWidth      =   4815
         TabIndex        =   3
         Top             =   345
         Width           =   4815
         Begin VB.Line Line1 
            BorderColor     =   &H0080FF80&
            X1              =   3360
            X2              =   3360
            Y1              =   360
            Y2              =   0
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00808080&
         Height          =   615
         Left            =   -240
         TabIndex        =   4
         Top             =   2520
         Width           =   5175
      End
      Begin VB.Image Img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1500
         Index           =   0
         Left            =   120
         Stretch         =   -1  'True
         Top             =   720
         Visible         =   0   'False
         Width           =   1650
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   1500
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin VB.Image Image3 
      Height          =   210
      Left            =   4800
      Picture         =   "Ruler.ctx":0331
      Top             =   2280
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image Image2 
      Height          =   405
      Left            =   6240
      Picture         =   "Ruler.ctx":04C3
      Top             =   2160
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image Image1 
      Height          =   405
      Left            =   6480
      Picture         =   "Ruler.ctx":07BA
      Top             =   240
      Visible         =   0   'False
      Width           =   195
   End
End
Attribute VB_Name = "Ruler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim PicCount As Integer, X1!, Selected%, PicLast As Boolean, Offst!, CourentIndex%
Public Event Click()
Public Event AddPic(sName As String)
Public Event MuoseMove(Button%, Shift%, X!, Y!)
Public Event Clear()
Public Event DelSelected()
Public Event ListChange()

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        RaiseEvent MuoseMove(Button, Shift, X, Y)
End Sub

Private Sub HScroll1_Scroll()
        HScroll1_ValueChanged
End Sub



Private Sub HScroll1_MouseDown(Shift As Integer)
        On Error GoTo 4
        Dim s As Long
        Rsize
        s = HScroll1.Value
        If (UserControl.Width / 1.5) < (-s + Frame4.Width) Then
            Frame4.left = -s
        Else
            Frame4.left = ((UserControl.Width / 1.5) - Frame4.Width)  ' + 50
        End If
        Exit Sub
4
        MsgBox "Picture is Invalid or list is full"

End Sub

Private Sub HScroll1_ValueChanged()
        On Error GoTo 4
        Dim s As Long, Z As Long
        Rsize
        s = HScroll1.Value
        If (UserControl.Width / 1.5) < (-s + Frame4.Width) Then
            Frame4.left = -s
        Else
            Frame4.left = ((UserControl.Width / 1.5) - Frame4.Width) ' + 50
        End If
        Exit Sub
4
        MsgBox "Picture is Invalid or list is full"
End Sub

Private Sub Img_Click(Index As Integer)
On Error GoTo 4
        Label1.left = Img(Index).left - 80
        Label1.Visible = True
        Selected = Index
        RaiseEvent Click
4 End Sub
Public Property Get MoveTime() As String
On Error GoTo 4
Dim ofset%
        If PicCount = 0 Then
            MoveTime = "00:00:00"
            GoTo 4
        End If
        ofset = IIf(Offst > 1650, Fix(Offst - Img(CourentIndex).left), Fix(Offst))
        MoveTime = TTim(ofset / 7.7, CourentIndex * 2)
4 End Property

Public Property Get PicList(Index As Integer) As String
On Error GoTo 4
        If PicCount <> 0 Then
            PicList = List1.List(Index)
        End If
4 End Property

Public Property Get FilmTime() As String
On Error GoTo 4
         FilmTime = ProsesTime(PicCount * 2)
4 End Property
Public Property Get CurentPic(ByRef ImgTemp As Image) As Image
On Error GoTo 4
        Set CurentPic = ImgTemp
        If PicCount <> 0 Then CurentPic.Picture = LoadPicture(List1.List(Selected))
4 End Property

Public Property Get ListCount() As Integer
On Error GoTo 4
        ListCount = List1.ListCount
4 End Property
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        On Error GoTo 4
        Disp X + Label1.left
        Offst = X
        CourentIndex = Selected
        RaiseEvent MuoseMove(Button, Shift, X, Y)
4
End Sub

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Disp X
        Offst = X
        RaiseEvent MuoseMove(Button, Shift, X, Y)
End Sub


Private Sub Img_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 4
        Disp Img(Index).left + X
        CourentIndex = Index
        Offst = X
        RaiseEvent MuoseMove(Button, Shift, X, Y)
4 End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Disp X
        RaiseEvent MuoseMove(Button, Shift, X, Y)
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Disp X
        RaiseEvent MuoseMove(Button, Shift, X, Y)
End Sub


Private Sub UserControl_Initialize()
        UserControl.Refresh
        Disp 0
        Rsize
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Disp X
        RaiseEvent MuoseMove(Button, Shift, X, Y)
End Sub
Private Function Disp(X!)
On Error GoTo 4
        Line1.X1 = X
        Line1.Y1 = 0
        Line1.X2 = X
        Line1.Y2 = Picture1.Height
        Line2.X1 = X
        Line2.Y1 = 0
        Line2.X2 = X
        Line2.Y2 = Picture2.Height
4 End Function
Private Sub UserControl_Resize()
        Rsize
End Sub
Public Sub SetList(ByRef Lst As ListBox)
On Error GoTo 4
        Set Lst = UserControl.List1
4  End Sub
Private Sub Rsize()
        On Error GoTo 4
        Dim V&
        Dim i&, W&, H&, t%
        DoEvents
        If PicCount = 0 Then
            Picture1.Width = UserControl.Width + 300
            Picture2.Width = UserControl.Width + 300
            Frame1.Width = UserControl.Width + 300
            Frame2.Width = UserControl.Width + 300
            Frame4.Width = UserControl.Width
        End If
        hssize UserControl.Width - 4080
        UserControl.Height = 2535
        Picture1.Width = Frame4.Width + 300
        Picture2.Width = Frame4.Width + 300
        Frame1.Width = Frame4.Width + 300
        Frame2.Width = Frame4.Width + 300
        Picture1.Cls: Picture2.Cls
        V = HScroll1.Value
        W = Picture1.Width
        H = Picture1.Height
        Picture1.PaintPicture Image1.Picture, V, 0, UserControl.Width + 50, H
        Picture2.PaintPicture Image2.Picture, V, 0, UserControl.Width + 50, H
        Set Picture1.Picture = Picture1.Image
        Set Picture2.Picture = Picture2.Image
        For i = 0 To W Step 60
            If i Mod 4 = 0 Then
                Picture1.Line (i, 0)-(i, H / 2)
                Picture2.Line (i, H / 2)-(i, H)
                        t = t + 1
            End If
            If t = 10 Then
                Picture1.Line (i, 0)-(i, H), RGB(150, 170, 15)
                Picture2.Line (i, 0)-(i, H), RGB(150, 170, 15)
                t = 0
            End If
        Next
4
End Sub
Public Sub AddPic(sName As String)
        On Error GoTo 4
        DoEvents
        If PicCount >= 136 Then Exit Sub
        Picture3.Picture = LoadPicture(sName)
        Picture3.PaintPicture Picture3.Picture, 0, 0, Picture3.Width, Picture3.Height
        Set Picture3.Picture = Picture3.Image
        Picture3.Refresh
        UserControl.Img(PicCount).Picture = Picture3.Picture
        UserControl.Img(PicCount).Visible = True
        List1.AddItem sName
        PicLast = False
        If PicCount <> 0 Then
                UserControl.Img(PicCount).left = UserControl.Img(PicCount - 1).left + 1800
                Frame4.Width = UserControl.Img(PicCount).left + 1800
        End If
        Rsize
        HsCall
        PicCount = PicCount + 1
        Load UserControl.Img(PicCount)
        RaiseEvent AddPic(sName)
        RaiseEvent ListChange
4
End Sub
Private Sub HsCall()
On Error GoTo 4
        If PicCount = 0 Then
            Frame4.Width = UserControl.Width
        End If
    Picture1.Width = Frame4.Width + 300
    Picture2.Width = Frame4.Width + 300
    Frame1.Width = Frame4.Width + 300
    Frame2.Width = Frame4.Width + 300
        If Frame4.Width > UserControl.Width Then
        HScroll1.Max = Frame4.Width
        HScroll1.Visible = True
        Picture4.Visible = True
        Else
        HScroll1.Visible = False
        Picture4.Visible = False
        Frame4.left = 0
        End If
4 End Sub
Private Sub UserControl_Terminate()
UserControl.Refresh
End Sub
Public Sub Clear()
On Error GoTo 4
        Dim n%
        If PicCount = 0 Then Exit Sub
        For n = PicCount To 1 Step -1
            Unload Img(n)
        Next
        PicCount = 0
        Img(0).Visible = False
        HScroll1.Visible = False
        Picture4.Visible = False
        HScroll1.Value = 0
        Label1.Visible = False
        Frame4.left = 0
        List1.Clear
        PicLast = False
        RaiseEvent Clear
        RaiseEvent ListChange
4 End Sub
Public Sub DelSelected()
On Error GoTo 4
        Dim V%
        If Label1.Visible = False Then Exit Sub
        For V = Selected To PicCount - 1
            Img(V).Picture = Img(V + 1).Picture
        Next
        List1.RemoveItem Selected
        Unload Img(PicCount)
        PicCount = PicCount - 1
        Frame4.Width = Frame4.Width - 1800
        Img(PicCount).Visible = False
         HScroll1_ValueChanged
        HsCall
        If Selected = PicCount Then Label1.Visible = False
        If PicCount = 0 Then
            HScroll1.Visible = False
            Picture4.Visible = False
            Label1.Visible = False
            Frame4.left = 0
        End If
        RaiseEvent DelSelected
        RaiseEvent ListChange
4 End Sub
Public Property Get EndFilm() As Boolean
        EndFilm = PicLast
End Property

Public Sub NextImg()
On Error GoTo 4
        If Selected = PicCount - 1 Then PicLast = True
        If Selected = PicCount - 1 Or PicCount = 0 Then Exit Sub
        Selected = Selected + 1
        Label1.left = Img(Selected).left - 80
4 End Sub
Public Sub MoveFirst()
On Error GoTo 4
        If PicCount = 0 Then Exit Sub
        Selected = 0
        HScroll1.Value = 0
        Frame4.left = 0
        Label1.left = Img(Selected).left - 80
        PicLast = False
4 End Sub
Public Sub MoveNext()
On Error GoTo 4
    If Selected = PicCount - 1 Or Label1.Visible = False Then Exit Sub
    Picture3.Picture = Img(Selected).Picture
    Img(Selected).Picture = Img(Selected + 1).Picture
    Frame4.Caption = List1.List(Selected)
    List1.List(Selected) = List1.List(Selected + 1)
    List1.List(Selected + 1) = Frame4.Caption
    Frame4.Caption = ""
    Img(Selected + 1).Picture = Picture3.Picture
    Selected = Selected + 1
    Label1.left = Img(Selected).left - 80
    PicLast = False
    RaiseEvent ListChange
4 End Sub
Public Sub MoveBack()
On Error GoTo 4
    If Selected = 0 Or Label1.Visible = False Then Exit Sub
    Picture3.Picture = Img(Selected).Picture
    Img(Selected).Picture = Img(Selected - 1).Picture
    Frame4.Caption = List1.List(Selected)
    List1.List(Selected) = List1.List(Selected - 1)
   List1.List(Selected - 1) = Frame4.Caption
    Frame4.Caption = ""
    Img(Selected - 1).Picture = Picture3.Picture
    Selected = Selected - 1
    Label1.left = Img(Selected).left - 80
    PicLast = False
    RaiseEvent ListChange
4 End Sub
'
Private Sub hssize(W As Integer)
 On Error GoTo 4
    If W < 100 Then GoTo 4
    Picture4.Width = W
    Picture4.Height = 240
    Picture4.PaintPicture Image3.Picture, 0, 0, W, 240
    Set Picture4.Picture = Picture4.Image
    Set HScroll1.RailPicture = Picture4.Picture
    HScroll1.Width = (W + 4) / 14
4 End Sub

