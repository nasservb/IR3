VERSION 5.00
Begin VB.UserControl cpvSlider 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   CanGetFocus     =   0   'False
   ClientHeight    =   435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   705
   ClipControls    =   0   'False
   ScaleHeight     =   29
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   47
   ToolboxBitmap   =   "cpvSlider.ctx":0000
   Begin VB.Image iRailPicture 
      Height          =   300
      Left            =   240
      Top             =   120
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image Slider 
      Height          =   240
      Left            =   0
      Picture         =   "cpvSlider.ctx":0312
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "cpvSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'##  cpvSlider OCX v1.1  ##
'##                      ##
'##  Carles P.V. - 2001  ##
'##  carles_pv@terra.es  ##






Option Explicit
'
'## API declarations_
'
Private Declare Function DrawEdge Lib "user32" _
                        (ByVal hdc As Long, _
                         qrc As RECT, _
                         ByVal edge As Long, _
                         ByVal grfFlags As Long) As Long

Private Const BDR_SUNKEN = &HA
Private Const BDR_RAISED = &H5
Private Const BDR_SUNKENOUTER = &H2
Private Const BDR_RAISEDINNER = &H4

Private Const BF_LEFT = &H1
Private Const BF_RIGHT = &H4
Private Const BF_TOP = &H2
Private Const BF_BOTTOM = &H8
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Declare Function GetWindowRect Lib "user32" _
                        (ByVal hwnd As Long, _
                         lpRect As RECT) As Long
                         
Private Declare Function SetWindowPos Lib "user32" _
                        (ByVal hwnd As Long, _
                         ByVal hWndInsertAfter As Long, _
                         ByVal x As Long, ByVal y As Long, _
                         ByVal cx As Long, ByVal cy As Long, _
                         ByVal wFlags As Long) As Long
                         
Private Const HWND_TOP = 0
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
                         
Private Type RECT
             Left As Long
             Top As Long
             Right As Long
             Bottom As Long
End Type
'
'## UC Types and Constants:
'
Private Type Point
             x As Single
             y As Single
End Type

Public Enum sOrientationConstants
            [Horizontal]
            [Vertical]
End Enum

Public Enum sRailStyleConstants
            [Sunken]
            [Raised]
            [SunkenSoft]
            [RaisedSoft]
            [ByPicture] = 99
End Enum
'
'## Private Variables:
'
Private SliderHooked As Boolean '# Slider hooked
Private SliderOffset As Point   '# Slider anchor point

Private R As RECT               '# Rail rectangle
Private AbsCount As Long        '# AbsCount = Max - Min
Private LastValue As Long       '# Last slider value

'
'## Default Property Values:
'
Const m_def_Enabled = True
Const m_def_Orientation = 1     '# Vertical
Const m_def_RailStyle = 0       '# Sunken
Const m_def_ShowValueTip = True '# Show Tip
Const m_def_Min = 0             '# Min = 0
Const m_def_Max = 10            '# Max = 10
Const m_def_Value = 0           '# Value = 0
'
'## Property Variables:
'
Dim m_Enabled As Boolean
Dim m_Orientation As Variant
Dim m_RailStyle As Variant
Dim m_ShowValueTip As Boolean
Dim m_Min As Long
Dim m_Max As Long
Dim m_Value As Long

'
'## Event Declarations:
'
Public Event Click()
Public Event ArrivedFirst()
Public Event ArrivedLast()
Public Event ValueChanged()
Public Event MouseDown(Shift As Integer)
Public Event MouseUp(Shift As Integer)







'##
'## UserControl: InitProperties/ReadProperties/WriteProperties
'##

Private Sub UserControl_InitProperties()

    m_Enabled = m_def_Enabled
    m_Orientation = m_def_Orientation
    m_RailStyle = m_def_RailStyle
    m_ShowValueTip = m_def_ShowValueTip
    m_Min = m_def_Min
    m_Max = m_def_Max
    m_Value = m_def_Value
    
    AbsCount = 10
    ResetSlider
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_Orientation = PropBag.ReadProperty("Orientation", m_def_Orientation)
    m_RailStyle = PropBag.ReadProperty("RailStyle", m_def_RailStyle)
    m_ShowValueTip = PropBag.ReadProperty("ShowValueTip", m_def_ShowValueTip)
    m_Min = PropBag.ReadProperty("Min", m_def_Min)
    m_Max = PropBag.ReadProperty("Max", m_def_Max)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    
    Set Slider.Picture = PropBag.ReadProperty("SliderIcon", Nothing)
    Set iRailPicture = PropBag.ReadProperty("RailPicture", Nothing)
    '
    '# Get absolute count and set Slider position
    '
    AbsCount = m_Max - m_Min
    LastValue = m_Min
    Slider.Left = (m_Value - m_Min) * (ScaleWidth - Slider.Width) / AbsCount
    Slider.Top = (ScaleHeight - Slider.Height) - (m_Value - m_Min) * (ScaleHeight - Slider.Height) / AbsCount
  
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("SliderIcon", Slider.Picture, Nothing)
    Call PropBag.WriteProperty("Orientation", m_Orientation, m_def_Orientation)
    Call PropBag.WriteProperty("RailPicture", iRailPicture, Nothing)
    Call PropBag.WriteProperty("RailStyle", m_RailStyle, m_def_RailStyle)
    Call PropBag.WriteProperty("ShowValueTip", m_ShowValueTip, m_def_ShowValueTip)
    Call PropBag.WriteProperty("Min", m_Min, m_def_Min)
    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)

End Sub






'##
'## UserControl draw
'##

Private Sub UserControl_Show()
    '
    '## Draw control
    '
    Refresh
    
End Sub

Private Sub UserControl_Resize()

    On Error Resume Next
    '
    '## Resize control
    '
    If m_RailStyle = 99 And iRailPicture <> 0 Then
'        iRailPicture.Width = UserControl.Width
'        iRailPicture.Height = UserControl.Height
'        iRailPicture.Left = 0
        
        Select Case m_Orientation

            Case 0 '# Horizontal
                    If Slider.Height < iRailPicture.Height Then
                       Size iRailPicture.Width * 15 + 60, iRailPicture.Height * 15
                    Else
                       Size iRailPicture.Width * 15 + 60, Slider.Height * 15
                    End If

            Case 1 '# Vertical
                    If Slider.Width < iRailPicture.Width Then
                       Size iRailPicture.Width * 15, iRailPicture.Height * 15 + 60
                    Else
                       Size Slider.Width * 15, iRailPicture.Height * 15 + 60
                    End If

        End Select
    
    Else
    
        Select Case m_Orientation
            
            Case 0 '# Horizontal
                    If Width = 0 Then Width = Slider.Width * 15
                    Height = Slider.Height * 15
                    
            Case 1 '# Vertical
                    If Height = 0 Then Height = Slider.Height * 15
                    Width = (Slider.Width) * 15
            
        End Select
    
    End If
    '
    '## Update slider position
    '
    Select Case m_Orientation
    
        Case 0 '# Horizontal
                If Slider.Height < iRailPicture.Height And _
                   m_RailStyle = 99 And _
                   iRailPicture <> 0 Then
                   Slider.Top = (iRailPicture.Height - Slider.Height) * 0.5
                Else
                   Slider.Top = 0
                End If
                Slider.Left = (m_Value - m_Min) * (ScaleWidth - Slider.Width) / AbsCount
        
        Case 1 '# Vertical
                If Slider.Width < iRailPicture.Width And _
                   m_RailStyle = 99 And _
                   iRailPicture <> 0 Then
                   Slider.Left = (iRailPicture.Width - Slider.Width) * 0.5
                Else
                   Slider.Left = 0
                End If
                Slider.Top = ScaleHeight - Slider.Height - (m_Value - m_Min) * (ScaleHeight - Slider.Height) / AbsCount
   
    End Select
    '
    '## Define rail rectangle
    '
    Select Case m_Orientation
        
        Case 0 '# Horizontal
                R.Top = (Slider.Height - 4) * 0.5
                R.Bottom = R.Top + 4
                R.Left = Slider.Width * 0.5 - 2
                R.Right = R.Left + ScaleWidth - Slider.Width + 4
        
        Case 1 '# Vertical
                R.Top = Slider.Height * 0.5 - 2
                R.Bottom = R.Top + ScaleHeight - Slider.Height + 4
                R.Left = (Slider.Width - 4) * 0.5
                R.Right = R.Left + 4
    
    End Select
    '
    '# Refresh control
    '
    Refresh
    
    On Error GoTo 0
    
End Sub

Private Sub Refresh()
    '
    '## Clear control
    '
    Cls
    '
    '## Draw rail...
    '
    On Error Resume Next
    
    If m_RailStyle = 99 Then
    
        Select Case m_Orientation

            Case 0 '# Horizontal
                    PaintPicture iRailPicture, 2, (ScaleHeight - iRailPicture.Height) * 0.5

            Case 1 '# Vertical
                    PaintPicture iRailPicture, (ScaleWidth - iRailPicture.Width) * 0.5, 2

        End Select

    Else
    
       DrawEdge hdc, R, Choose(m_RailStyle + 1, &HA, &H5, &H2, &H4, 0), BF_RECT
    
    End If
    '
    '## ...and slider
    '
    PaintPicture Slider, Slider.Left, Slider.Top
    '
    '## Show value tip
    '
    If m_ShowValueTip And SliderHooked Then ShowTip
    
    On Error GoTo 0

End Sub






'##
'## Scrolling...
'##

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Not Me.Enabled Then Exit Sub
    
    With Slider
        '
        '# Hook slider, get offsets and show tip
        '
        If Button = vbLeftButton Then
           
            SliderHooked = True
            '
            '# If mouse over slider
            '
            If x >= .Left And x < .Left + .Width And _
               y >= .Top And y < .Top + .Height Then
               
                SliderOffset.x = x - .Left
                SliderOffset.y = y - .Top
            '
            '# If mouse over rail
            '
            Else
                SliderOffset.x = .Width / 2
                SliderOffset.y = .Height / 2
                UserControl_MouseMove Button, Shift, x, y
                
            End If
            '
            '# Show tip
            '
            If m_ShowValueTip Then
                 ShowTip
            End If
            
            RaiseEvent MouseDown(Shift)
           
        End If

    End With
 
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If SliderHooked Then
        '
        '## Check limits
        '
        With Slider
        Select Case m_Orientation
            
            Case 0 '# Horizontal
                    If x - SliderOffset.x < 0 Then
                        .Left = 0
                    ElseIf x - SliderOffset.x > ScaleWidth - .Width Then
                        .Left = ScaleWidth - .Width
                    Else
                        .Left = x - SliderOffset.x
                    End If
            
            Case 1 '# Vertical
                    If y - SliderOffset.y < 0 Then
                        .Top = 0
                    ElseIf y - SliderOffset.y > ScaleHeight - .Height Then
                        .Top = ScaleHeight - .Height
                    Else
                        .Top = y - SliderOffset.y
                    End If
        
        End Select
        End With
        '
        '## Get value from Slider position
        '
        Value = GetValue
    
    End If

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    '
    '# Click event (If mouse over control area)
    '
    If x >= 0 And x < ScaleWidth And _
       y >= 0 And y < ScaleHeight And _
       Button = vbLeftButton Then

       RaiseEvent Click
       
    End If
    '
    '# MouseUp event (If slider has been hooked)
    '
    If SliderHooked Then RaiseEvent MouseUp(Shift)
    '
    '## Unhook slider and hide value tip
    '
    SliderHooked = False
    Unload frmValueTip
    
End Sub






'##
'## Properties
'##

'## Enabled
    Public Property Get Enabled() As Boolean
        Enabled = m_Enabled
    End Property
    
    Public Property Let Enabled(ByVal New_Enabled As Boolean)
        m_Enabled = New_Enabled
        PropertyChanged "Enabled"
    End Property

'## BackColor
    Public Property Get BackColor() As OLE_COLOR
        BackColor = UserControl.BackColor
    End Property
    
    Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
        UserControl.BackColor() = New_BackColor
        Refresh
        PropertyChanged "BackColor"
    End Property
    
'## Max
    Public Property Get Max() As Long
        Max = m_Max
    End Property
    
    Public Property Let Max(ByVal New_Max As Long)
    
        If New_Max <= m_Min Then Err.Raise 380
        
        m_Max = New_Max
        AbsCount = m_Max - m_Min
        PropertyChanged "Max"
        
    End Property

'## Min
    Public Property Get Min() As Long
        Min = m_Min
    End Property
    
    Public Property Let Min(ByVal New_Min As Long)
    
        If New_Min >= m_Max Then Err.Raise 380
        
        m_Min = New_Min
        Value = New_Min
        AbsCount = m_Max - m_Min
        PropertyChanged "Min"
        
    End Property

'## Value
    Public Property Get Value() As Long
Attribute Value.VB_UserMemId = 0
        Value = m_Value
    End Property
    
    Public Property Let Value(ByVal New_Value As Long)
    
        If New_Value < m_Min Or New_Value > m_Max Then Err.Raise 380
        m_Value = New_Value
            
            If m_Value <> LastValue Then
                
                If Not SliderHooked Then
                           
                    Select Case m_Orientation
        
                        Case 0 '# Horizontal
                                Slider.Left = (New_Value - m_Min) * (ScaleWidth - Slider.Width) / AbsCount
                        
                        Case 1 '# Vertical
                                Slider.Top = ScaleHeight - Slider.Height - (New_Value - m_Min) * (ScaleHeight - Slider.Height) / AbsCount
       
                    End Select
    
                End If
                
                Refresh
                LastValue = m_Value
                
                RaiseEvent ValueChanged
                If m_Value = m_Max Then RaiseEvent ArrivedLast
                If m_Value = m_Min Then RaiseEvent ArrivedFirst
        
            End If
            
        PropertyChanged "Value"
        
    End Property

'## Orientation
    Public Property Get Orientation() As sOrientationConstants
        Orientation = m_Orientation
    End Property
    
    Public Property Let Orientation(ByVal New_Orientation As sOrientationConstants)
        
        m_Orientation = New_Orientation
        
        ResetSlider
        UserControl_Resize
        
        PropertyChanged "Orientation"
        
    End Property
    
'## RailStyle
    Public Property Get RailStyle() As sRailStyleConstants
        RailStyle = m_RailStyle
    End Property
    
    Public Property Let RailStyle(ByVal New_RailStyle As sRailStyleConstants)
    
        m_RailStyle = New_RailStyle
       
        UserControl_Resize
        
        PropertyChanged "RailStyle"
        
    End Property
    
'## SliderIcon
    Public Property Get SliderIcon() As Picture
Attribute SliderIcon.VB_Description = "Devuelve o establece el gráfico que se mostrará en un control."
        Set SliderIcon = Slider.Picture
    End Property
    
    Public Property Set SliderIcon(ByVal New_SliderIcon As Picture)
    
        Set Slider.Picture = New_SliderIcon
        
        UserControl_Resize
        
        PropertyChanged "SliderIcon"
        
    End Property
    
'## RailPicture
    Public Property Get RailPicture() As Picture
        Set RailPicture = iRailPicture.Picture
    End Property
    
    Public Property Set RailPicture(ByVal New_RailPicture As Picture)
    
        Set iRailPicture.Picture = New_RailPicture
        
        UserControl_Resize
        
        PropertyChanged "RailPicture"
        
    End Property
    
'## ShowValueTip
    Public Property Get ShowValueTip() As Boolean
        ShowValueTip = m_ShowValueTip
    End Property
    
    Public Property Let ShowValueTip(ByVal New_ShowValueTip As Boolean)
        m_ShowValueTip = New_ShowValueTip
        PropertyChanged "ShowValueTip"
    End Property






'##
'## Private Functions/Subs
'##

'#
'# Get value from Slider position
'#
Private Function GetValue() As Long
    
    On Error Resume Next
    Select Case m_Orientation
    
        Case 0 '# Horizontal
                GetValue = Slider.Left / (ScaleWidth - Slider.Width) * AbsCount + m_Min
                Slider.Left = (GetValue - m_Min) * (ScaleWidth - Slider.Width) / AbsCount
        
        Case 1 '# Vertical
                GetValue = (ScaleHeight - Slider.Height - Slider.Top) / (ScaleHeight - Slider.Height) * AbsCount + m_Min
                Slider.Top = ScaleHeight - Slider.Height - (GetValue - m_Min) * (ScaleHeight - Slider.Height) / AbsCount
   
    End Select
    On Error GoTo 0
    
End Function
'#
'# Reset slider position
'#
Private Sub ResetSlider()

    Select Case m_Orientation
        
        Case 0 '# Horizontal
                Slider.Move 0, 0
             
        Case 1 '# Vertical
                Slider.Move 0, ScaleHeight - Slider.Height
             
    End Select
    
End Sub
'#
'# Show value tip
'#
Private Sub ShowTip()
    
    Dim ucR As RECT
    Dim x As Long, y As Long

    On Error Resume Next
    
    GetWindowRect hwnd, ucR
    
    With frmValueTip
    
        .lblTip.Width = .TextWidth(m_Value)
        .lblTip.Caption = m_Value
        .lblTip.Refresh
        
        Select Case m_Orientation
            
            Case 0 '# Horizontal
                    x = ucR.Left + Slider.Left + (Slider.Width - .lblTip.Width - 4) * 0.5
                    y = ucR.Top + Slider.Top - .lblTip.Height - 5
                 
            Case 1 '# Vertical
                    x = ucR.Left + Slider.Left - .lblTip.Width - 6
                    y = ucR.Top + Slider.Top + (Slider.Height - .lblTip.Height - 4) * 0.5
                 
        End Select
        '
        '# Set Tip position...
        '
        .Move x * 15, y * 15, (.lblTip.Width + 4) * 15, (.lblTip.Height + 3) * 15
        '
        '# ...and show it
        '
        SetWindowPos .hwnd, HWND_TOP, 0, 0, 0, 0, _
                            SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
    
    End With
    
    On Error GoTo 0

End Sub
