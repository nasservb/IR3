VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm Main 
   BackColor       =   &H00808080&
   Caption         =   "IranVideo 5.3"
   ClientHeight    =   8775
   ClientLeft      =   1230
   ClientTop       =   1320
   ClientWidth     =   12030
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2280
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "chose File"
      Filter          =   "IranVideo Projecr|*.irp|All Fille's|*.*"
   End
   Begin VB.Menu mFile 
      Caption         =   "File"
      Begin VB.Menu mNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mSaveas 
         Caption         =   "Save as.."
         Shortcut        =   {F12}
      End
      Begin VB.Menu nocod4 
         Caption         =   "-"
      End
      Begin VB.Menu mImport 
         Caption         =   "Import"
         Shortcut        =   ^I
      End
      Begin VB.Menu mExport 
         Caption         =   "Export"
         Shortcut        =   ^E
      End
      Begin VB.Menu nocode 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mEdit 
      Caption         =   "Edit"
      Begin VB.Menu mAdd 
         Caption         =   "Add"
         Shortcut        =   ^A
      End
      Begin VB.Menu mDel 
         Caption         =   "Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mClear 
         Caption         =   "Clear"
         Shortcut        =   ^X
      End
      Begin VB.Menu nocod5 
         Caption         =   "-"
      End
      Begin VB.Menu mMoveN 
         Caption         =   "MoveNext"
      End
      Begin VB.Menu mMoveB 
         Caption         =   "MoveBack"
      End
   End
   Begin VB.Menu mPlay 
      Caption         =   "Play"
      Begin VB.Menu msPlay 
         Caption         =   "Play"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mPause 
         Caption         =   "Pause"
      End
      Begin VB.Menu mStop 
         Caption         =   "Stop"
      End
   End
   Begin VB.Menu mWindows 
      Caption         =   "Windows"
      Begin VB.Menu mMonitor 
         Caption         =   "Monitor"
      End
      Begin VB.Menu mTimeLine 
         Caption         =   "TimeLine"
      End
      Begin VB.Menu mExplorer 
         Caption         =   "Explorer"
      End
      Begin VB.Menu mShowPlayer 
         Caption         =   "ShowPlayer"
      End
   End
   Begin VB.Menu mHelp 
      Caption         =   "Help"
      Begin VB.Menu mHelpm 
         Caption         =   "Online Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mSuport 
         Caption         =   "Suport Page"
      End
      Begin VB.Menu mAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public IrpFile$
Private Sub mAbout_Click()
Monitor.Timer1.Enabled = False
About.Show 1
End Sub

Private Sub mAdd_Click()
        Saved = True
        Main.Caption = "IranVideo 5.3*"
Monitor.Timer1.Enabled = False
    t6 = 0
    TimeLine.CommonDialog1.ShowOpen
    If TimeLine.CommonDialog1.filename <> "" Then _
        TimeLine.Ruler1.AddPic TimeLine.CommonDialog1.filename
End Sub

Private Sub mClear_Click()
        Saved = True
        Main.Caption = "IranVideo 5.3*"
Monitor.Timer1.Enabled = False
        TimeLine.Ruler1.Clear
End Sub

Private Sub mDel_Click()
        Saved = True
        Main.Caption = "IranVideo 5.3*"
Monitor.Timer1.Enabled = False
        TimeLine.Ruler1.DelSelected
End Sub

Private Sub MDIForm_Load()
        Saved = False
        Monitor.Show
        Monitor.Width = 6330
        MDIForm_Resize
        TimeLine.Show
        Explorer.Show
        If IrpFile <> "" Then irpOpen Me.IrpFile
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo 4
Dim f As Integer
        If Saved = True Then
            f = MsgBox("Do you want to save project!", vbInformation + vbYesNoCancel, Me.Caption)
            If f = vbYes Then
                mSave_Click
            ElseIf f = vbCancel Then
                Cancel = True
                Exit Sub
            End If
        End If
        If FSO.FolderExists(App.Path & "\temp") = True Then FSO.DeleteFolder (App.Path & "\temp")
        If FSO.FolderExists(App.Path & "\fx") = True Then FSO.DeleteFolder (App.Path & "\fx")
        If FSO.FolderExists(App.Path & "\import") = True Then FSO.DeleteFolder (App.Path & "\import")
        Unload export
4
End Sub

Private Sub MDIForm_Resize()
        Dim B%
        On Error GoTo 4
        B = Monitor.Width
        Monitor.top = 0
        Monitor.left = Me.Width - (B + 400)
        TimeLine.top = Me.Height - (TimeLine.Height + 900)
        TimeLine.left = 0
        TimeLine.Width = Me.Width - 400
4
End Sub


Private Sub mExit_Click()
        Unload Me
End Sub

Private Sub mExplorer_Click()
        Explorer.Show
End Sub

Private Sub mExport_Click()
        export.Show 1
End Sub

Private Sub mHelpm_Click()
Shell "explorer.exe" + " http://www.tcvb.tk/ir3help", vbNormalFocus
End Sub

Private Sub mImport_Click()
On Error GoTo 4
        Saved = True
        Main.Caption = "IranVideo 5.3*"
Monitor.Timer1.Enabled = False
 CommonDialog1.Filter = "all IranVideiFile|*.ir3"
 CommonDialog1.ShowOpen
 If CommonDialog1.filename = "" Then Exit Sub
 If LCase(right(CommonDialog1.filename, 3)) = "ir3" Then Import CommonDialog1.filename
4
End Sub

Private Sub mMonitor_Click()
        Monitor.Show
End Sub

Private Sub mOutPut_Click()
        export.Show 1
End Sub

Private Sub mMoveB_Click()
        Saved = True
        Main.Caption = "IranVideo 5.3*"
Monitor.Timer1.Enabled = False
        TimeLine.Ruler1.MoveBack
End Sub

Private Sub mMoveN_Click()
        Saved = True
        Main.Caption = "IranVideo 5.3*"
Monitor.Timer1.Enabled = False
        TimeLine.Ruler1.MoveNext
End Sub

Private Sub mNew_Click()
        If Saved = True Then
            f = MsgBox("Do you want to save project!", vbInformation + vbYesNoCancel, Me.Caption)
            If f = vbYes Then
                mSave_Click
            ElseIf f = vbCancel Then
                Cancel = True
                Exit Sub
            End If
        End If
        Saved = False
        Main.Caption = "IranVideo 5.3"
Monitor.Timer1.Enabled = False
        TimeLine.Ruler1.Clear
        Unload Monitor
End Sub

Private Sub mOpen_Click()
On Error GoTo 4
        If Saved = True Then
            f = MsgBox("Do you want to save project!", vbInformation + vbYesNoCancel, Me.Caption)
            If f = vbYes Then
                mSave_Click
            ElseIf f = vbCancel Then
                Cancel = True
                Exit Sub
            End If
        End If
        Saved = False
        Main.Caption = "IranVideo 5.3"
Monitor.Timer1.Enabled = False
        CommonDialog1.Filter = "all Ir3Project|*.irp"
        CommonDialog1.ShowOpen
        If CommonDialog1.filename <> "" Then irpOpen CommonDialog1.filename
4 End Sub

Private Sub mPause_Click()
        Monitor.Show
        Monitor.Timer1.Enabled = False
        Monitor.pus = True
End Sub


Private Sub mSave_Click()
On Error GoTo 4
Monitor.Timer1.Enabled = False
        Dim s$
        Dim Lst As ListBox
        TimeLine.Ruler1.SetList Lst
        Saved = False
        Main.Caption = "IranVideo 5.3"
        If Lst.ListCount = 0 Then Exit Sub
        CommonDialog1.Filter = "all Ir3Project|*.irp"
        If CommonDialog1.filename = "" Then CommonDialog1.ShowSave
        Open CommonDialog1.filename For Output As #1
        i = 0
        For i = 0 To Lst.ListCount - 1
            Print #1, Lst.List(i)
        Next
        Close
4 End Sub

Private Sub mSaveas_Click()
On Error GoTo 4
        Dim s$
        Dim Lst As ListBox
        Monitor.Timer1.Enabled = False
        TimeLine.Ruler1.SetList Lst
        Saved = False
        Main.Caption = "IranVideo 5.3"
        If Lst.ListCount = 0 Then Exit Sub
        CommonDialog1.Filter = "all Ir3Project|*.irp"
        CommonDialog1.ShowSave
        If CommonDialog1.filename = "" Then Exit Sub
        Open CommonDialog1.filename For Output As #1
        i = 0
        For i = 0 To Lst.ListCount - 1
            Print #1, Lst.List(i)
        Next
        Close
4 End Sub

Private Sub mShowPlayer_Click()
Monitor.Timer1.Enabled = False
        Command1 = ""
        Player.Show
End Sub

Private Sub msPlay_Click()
    If Monitor.pus = False Then
        TimeLine.Ruler1.MoveFirst
        TimeLine.Ruler1.CurentPic(Monitor.Image1).Refresh
    End If
    Monitor.pus = False
    Monitor.Timer1.Enabled = True
End Sub

Private Sub mStop_Click()
On Error GoTo 4
    Monitor.Timer1.Enabled = False
    If TimeLine.Ruler1.EndFilm = True Then Exit Sub
    TimeLine.Ruler1.MoveFirst
    TimeLine.Ruler1.CurentPic(Monitor.Image1).Refresh
    Monitor.pus = False
4 End Sub

Private Sub mSuport_Click()
Shell "explorer.exe" + " http://www.vbook.coo.ir/iranvideo", vbNormalFocus
End Sub

Private Sub mTimeLine_Click()
        TimeLine.Show
End Sub
Private Sub irpOpen(filename$)
On Error GoTo 4
Dim s$
TimeLine.Ruler1.Clear
        Open filename For Input As #1
        i = 0
        While EOF(1) = False
           Line Input #1, s
            If FSO.FileExists(s) = True Then TimeLine.Ruler1.AddPic s
        Wend
        Close #1
4 End Sub


