Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Byte
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Type RGBComponent
        R As Integer
        g As Integer
        B As Integer
End Type
Public Enum EFilterTypes
    [_Min]
    eBlur
    eBlurMore
    eSoften
    eSoftenMore
    eSharpen
    eSharpenMore
    eUnSharp
    eEmboss
    eMedian
    eMinimum
    eMaximum
    eCount
    eCustom
    [_Max]
End Enum
Public m_eFilterType As EFilterTypes
Public CanFX As Boolean
Public Sub Brightnes(Picture1 As PictureBox, TxtBrightness%)
On Error GoTo 4
        Dim Brightness As Single
        Dim NewColor As Long
        Dim X%, Y As Integer
        Dim R%, g%, B As Integer
        Brightness = TxtBrightness / 100
        CanFX = False
        Explorer.Frame8.Visible = False
        Explorer.ProgressBar1.Max = Picture1.ScaleWidth
        For X = 0 To Picture1.ScaleWidth
                If CanFX = True Then Exit For
                If X Mod 30 = 0 Then
                    Explorer.ProgressBar1.Value = X
                    Explorer.ProgressBar1.Refresh
                    DoEvents
                End If
                For Y = 0 To Picture1.ScaleHeight
                        NewColor = GetPixel(Picture1.hDC, X, Y)
                        R = (NewColor Mod 256)
                        B = (Int(NewColor / 65536))
                        g = ((NewColor - (B * 65536) - R) / 256)
                        R = R * Brightness
                        B = B * Brightness
                        g = g * Brightness
                        If R > 255 Then R = 255
                        If R < 0 Then R = 0
                        If B > 255 Then B = 255
                        If B < 0 Then B = 0
                        'set the new pixel
                        If g > 255 Then g = 255
                        If g < 0 Then g = 0
                        SetPixelV Picture1.hDC, X, Y, RGB(R, g, B)
                Next Y
        Next X
        Set Picture1.Picture = Picture1.Image
        Explorer.Frame8.Visible = True
        Picture1.Refresh
4 End Sub
'---------------------------------------------------------------------------
Function AverageComponent(RGBComponents As RGBComponent) As Byte
On Error GoTo 4
        With RGBComponents
                AverageComponent = (.R + .g + .B) \ 3
        End With
4 End Function
Function ColorSplit(ByVal SplitColor As Long) As RGBComponent
On Error GoTo 4
        With ColorSplit
                .R = SplitColor And &HFF
                .g = (SplitColor \ &H100) And &HFF
                .B = (SplitColor \ &H10000) And &HFF
        End With
4 End Function
Function MaximalRGB(RGBComponents As RGBComponent) As Byte
On Error GoTo 4
        With RGBComponents
                If .R >= .g Then
                        If .R >= .B Then
                                MaximalRGB = .R
                        Else
                                MaximalRGB = .B
                        End If
                Else
                        If .g >= .B Then
                                MaximalRGB = .g
                        Else
                                MaximalRGB = .B
                        End If
                End If
        End With
4 End Function
Public Sub ColorPalette(ByVal PaletteColor As Long, ByRef srcPicture As PictureBox)
On Error GoTo 4
    Dim CurrentColor As Long
    Dim Max As Integer
    Dim nX As Integer
    Dim nY As Integer
    Dim R1 As Integer
    Dim G1 As Integer
    Dim B1 As Integer
    Dim VarAverageComponent As Long
    Dim VarCurrentColorRGBComponent As RGBComponent
    Dim VarRGBComponent As RGBComponent
    CanFX = False
    Explorer.Frame9.Visible = False
    Explorer.ProgressBar1.Max = srcPicture.ScaleHeight
    For nY = 0 To srcPicture.ScaleHeight
        If CanFX = True Then Exit For
        If Y Mod 30 = 0 Then
            Explorer.ProgressBar1.Value = nY
            Explorer.ProgressBar1.Refresh
            DoEvents
        End If
        For nX = 0 To srcPicture.ScaleWidth
            VarRGBComponent = ColorSplit(PaletteColor)
            Max = MaximalRGB(VarRGBComponent)
            CurrentColor = GetPixel(srcPicture.hDC, nX, nY)
            VarCurrentColorRGBComponent = ColorSplit(CurrentColor)
            VarAverageComponent = AverageComponent(VarCurrentColorRGBComponent)
            If Max Then
            R1 = VarAverageComponent * VarRGBComponent.R / Max
            G1 = VarAverageComponent * VarRGBComponent.g / Max
            B1 = VarAverageComponent * VarRGBComponent.B / Max
            Else
            R1 = VarAverageComponent
            G1 = VarAverageComponent
            B1 = VarAverageComponent
            End If
            SetPixel srcPicture.hDC, nX, nY, RGB(R1, G1, B1)
        Next nX, nY
        Explorer.Frame9.Visible = True
        Set srcPicture.Picture = srcPicture.Image
4 End Sub
Public Function TTim(Offset As Integer, Frame As Integer) As String '
Dim B1%, C1%
B1 = Frame
If (Offset >= 100) Then B1 = B1 + 1
If B1 < 60 Then
    TTim = "00:" + Fit(B1) + ":" + Fit(Offset): Exit Function
End If
C1 = B1 \ 60
B1 = B1 Mod 60
TTim = Fit(C1) + ":" + Fit(B1) + ":" + Fit(Offset)
End Function
Public Function Fit(A1 As Integer) As String
        If Len(Str(A1)) = 2 Then
        Fit = "0" + right(Str(A1), 1)
        Else: Fit = right(Str(A1), 2)
        End If
End Function
Public Function ProsesTime(A1 As Integer) As String
Dim B1%, C1%
B1 = A1
If B1 < 60 Then
ProsesTime = "00:" + Fit(B1) + ":00"
Exit Function
End If
C1 = B1 \ 60
B1 = B1 Mod 60
ProsesTime = Fit(C1) + ":" + Fit(B1) + ":00"
End Function



