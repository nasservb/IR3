Attribute VB_Name = "ir3_Render"
'------Typing New data For Seearch File---------------------
Public Type BrowseInfo
    hWndOwner       As Long
    pIDLRoot        As Long
    pszDisplayName  As Long
    lpszTitle       As Long
    ulFlags         As Long
    lpfnCallback    As Long
    lParam          As Long
    iImage          As Long
End Type '----------------------
Public Enum Mode
             Readay = 0
             play = 1
             Puase = 2
             stopv = 3
             EndFilm = 4
             NoFile = 5
             FilUnknow = 6
End Enum
'---------------Conset For Seearch--------------------
Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BIF_DONTGOBELOWDOMAIN = 2
Public Const MAX_PATH = 260
Public Const SEE_MASK_INVOKEIDLIST = &HC
Public Const SEE_MASK_NOCLOSEPROCESS = &H40
Public Const SEE_MASK_FLAG_NO_UI = &H400
Public Const ATTR_NORMAL = 0
Public Const ATTR_READONLY = 1
Public Const ATTR_HIDDEN = 2
Public Const ATTR_SYSTEM = 4
Public Const ATTR_VOLUME = 8
Public Const ATTR_DIRECTORY = 16
Public Const ATTR_ARCHIVE = 32
'-----------------------Declareing API------------------------------------------
Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, _
        ByVal lpBuffer As String) As Long '-------------------------------------
Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, _
        ByVal lpString2 As String) As Long '------------------------------------
Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public FSO As New FileSystemObject
Public Btmp%
Public Command1$, Model As Mode, D1%, D2%
'---------------------------------
Public Sub Save(inf As RenderInfo)
On Error GoTo 4
        Dim d%, R%, VF%, SS%, i%, g%, H% ', D1%, D2%
        VF = inf.Query
        Can = False
        t6 = 21
        D1 = inf.Size_Width
        D2 = inf.Size_Height
        d = inf.List.ListCount
        frmRender.Picture3.Picture = LoadPicture(inf.List.List(0))
        Close '-----------------------------------
        Open inf.filename For Random Access Write As #2 Len = 2
        Put #2, , CInt("1" + Fex(D1))
        Put #2, , CInt(Fex(D2))
        Put #2, , d
        Put #2, , -7000
        For g = 1 To D1
            For H = 1 To D2
                R = Int(frmRender.Picture3.Point(g, H))
                Put #2, , R
            Next H
        Next g
        '----------shroe pardazesh---------------
        frmRender.ProgressBar1.Max = d
        frmRender.Label2.Caption = "Saving to:" & inf.filename
    For i = 1 To d - 1
        Put #2, , -7000
        If Can = True Then Exit For
        If frmRender.ProgressBar2.Value < frmRender.ProgressBar2.Max And i Mod 2 Then _
                        frmRender.ProgressBar2.Value = frmRender.ProgressBar2.Value + 1
       frmRender.ProgressBar1.Value = i
       frmRender.Picture2.Picture = LoadPicture(inf.List.List(i))
       frmRender.Caption = "Rendering..." & left(Val((i / d) * 100), 5) & "%"
       frmRender.Label1.Caption = "File's:" & inf.List.List(i)
       For g = 1 To D1  '---------------------Starting Save---------------------
                For H = 1 To D2
                    Btmp = frmRender.Picture2.Point(g, H)
                    If Abs(Btmp - (frmRender.Picture3.Point(g, H))) <= VF Then  '----agar tekrary bod pixel
                        SS = SS - 1
                        If H = D2 And SS < 0 Then
                            Put #2, , SS: SS = 0
                        End If
                    Else '-------------agar tekrary nabod pixel
                        If SS < 0 Then '-------agar mizan 0 bod
                            Put #2, , SS: SS = 0
                        End If
                        '------------------agar mizan 0 nabod
                        Put #2, , Btmp
                    End If '------------------------End Analiz of Frafrm--------------
                Next H
                If g Mod 50 = 0 Then DoEvents
       Next g '----------------
          frmRender.Picture3.Picture = frmRender.Picture2.Picture
        Next i '----------------------End Save------------------
4        Close
        Unload frmRender
        Main.Visible = True
        Command1 = inf.filename
        Player.Show
 End Sub
Public Sub Import(FNafrm$)
On Error GoTo 4
     Dim d%, R%, y6%, VR%, g%, H%, kk& ', D1%, D2%
     Close
    frmRender.Show
    Main.Enabled = False
    cunt2 = 0
    If FSO.FolderExists(App.Path & "\import") = False Then FSO.CreateFolder (App.Path & "\import")
    Open FNafrm For Random Access Read As #2 Len = 2
    Get #2, , y6
    D1 = CInt(right(CStr(y6), 4))
    Get #2, , y6
    D2 = CInt(right(CStr(y6), 4))
    Get #2, , y6
    frmRender.Picture3.Width = (D1 + 4) * 15
    frmRender.Picture3.Height = (D2 + 4) * 15
    Can = False
    frmRender.ProgressBar2.Max = y6
    frmRender.ProgressBar1.Max = D1
    frmRender.Caption = "Importing ... "
    While EOF(2) = False '---------starting play file
         Get #2, , Btmp
        If Btmp <> -7000 Then
            MsgBox "The File Format Is  Unknow!", vbInformation, "IranVideo"
            GoTo 4
        End If
        If frmRender.ProgressBar2.Value = frmRender.ProgressBar2.Max Then frmRender.ProgressBar2.Value = 0
        frmRender.ProgressBar2.Value = frmRender.ProgressBar2.Value + 1
        frmRender.ProgressBar2.Refresh: frmRender.ProgressBar1.Refresh
        If Can = True Or TimeLine.Ruler1.EndFilm = True Then GoTo 4 '-----canceling
        For g = 1 To D1
            For H = 1 To D2
                Get #2, , Btmp
                If Btmp < 0 Then
                    H = (H + Abs(Btmp)) - 1
                Else
                    kk = Btmp
                    kk = kk * 512
                    frmRender.Picture3.PSet (g, H), kk
                End If
            Next H
        If g Mod 30 = 0 Then
            DoEvents
            frmRender.ProgressBar1.Value = g
        End If
        Next g
        Set frmRender.Picture3.Picture = frmRender.Picture3.Image
        cunt2 = cunt2 + 1
        SavePicture frmRender.Picture3, App.Path & "\import\" & cunt2 & ".bmp"
        TimeLine.Ruler1.AddPic App.Path & "\import\" & cunt2 & ".bmp"
    Wend
4: If Err.Number Then MsgBox "unnknow file format!", vbExclamation, Main.Caption
        Close
        Unload frmRender
        Main.Enabled = True
End Sub
Private Function Fex(a%) As String
On Error GoTo 4
        Dim M%, i%
        M = Len(CStr(a))
        If M = 4 Then
                Fex = CStr(a)
                Exit Function
        End If
        Fex = CStr(a)
        For i = 1 To Abs(M - 4)
                Fex = "0" + Fex
        Next
4 End Function

Public Sub LoadIr3()
On Error GoTo 4
        Dim g%, H%, kk&  ', D1%, D2%
        If Can = True Or Model = stopv Then
            Close
            Model = stopv
            Player.Caption = "Play Stoped"
            Player.Timer1.Enabled = False
            Player.HScroll1.Value = 0
            Exit Sub
        End If
        If EOF(2) = True Then
            Model = EndFilm
            Player.Caption = "Play Completed"
            Player.Timer1.Enabled = False
            Player.HScroll1.Value = 0
            Exit Sub
        End If
        Player.Caption = "Playing ... " & _
        TTim(0, Player.HScroll1.Value) & "/" & ProsesTime(Player.HScroll1.Max)
        For g = 1 To D1
            For H = 1 To D2
                Get #2, , Btmp
                If Btmp < 0 Then
                    H = (H + Abs(Btmp)) - 1
                Else
                    kk = Btmp
                    kk = kk * 512
                    Player.Picture1.PSet (g, H), kk
                End If
            Next H
            If g Mod 30 = 0 Then DoEvents
         Next g
        Get #2, , Btmp
        Set Player.Picture1.Picture = Player.Picture1.Image
        Player.Image1.Picture = Player.Picture1.Picture
        If Player.HScroll1.Value < Player.HScroll1.Max Then _
        Player.HScroll1.Value = Player.HScroll1.Value + 1
        Exit Sub
4
Player.HScroll1.Value = 0
Player.Timer1.Enabled = False
Player.Caption = "Playing Error"
Close
        If Err.Number = 52 Then
            Model = NoFile
            Player.Caption = "No File For Play"
        End If
End Sub
Public Sub ZipIr3(inf As RenderInfo, ByRef picName As PictureBox)
On Error GoTo 4
        Dim Ii%, Jj%, lg&, n%
    frmRender.Picture2.Picture = picName.Picture
    For Ii = 1 To inf.Size_Width
        For Jj = 1 To inf.Size_Height
            lg = frmRender.Picture2.Point(Ii, Jj)
             n = (lg \ 512)
            frmRender.Picture2.PSet (Ii, Jj), n
        Next Jj
    If Ii Mod 30 = 0 Then DoEvents
    Next Ii
    Set frmRender.Picture2.Picture = frmRender.Picture2.Image
    picName.Picture = frmRender.Picture2.Picture
4 End Sub
Public Sub SeekIR3(FNafrm As String, Position As Integer)
On Error GoTo 4
Dim Fream As Integer, y6%, frem%
    If Position = 0 Then Exit Sub
    Close
    Open FNafrm For Random Access Read As #2 Len = 2
    Get #2, , y6
    D1 = CInt(right(CStr(y6), 4))
    Get #2, , y6
    D2 = CInt(right(CStr(y6), 4))
    Get #2, , Fream
    Player.HScroll1.Max = Fream
    While Fream = 0 Or EOF(2) = False
        Get #2, , Btmp
        If Btmp = -7000 Then
            frem = frem + 1
            DoEvents
            If frem = Position Then GoTo 4
        End If
    Wend
4:  If Err.Number Then
        MsgBox "unknow file format!", vbExclamation, Main.Caption
        Close
    End If
    Player.Timer1.Enabled = True
    Player.MousePointer = 0
End Sub


