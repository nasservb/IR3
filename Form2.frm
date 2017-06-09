VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   10470
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   11640
      Left            =   0
      TabIndex        =   0
      Top             =   -4920
      Width           =   10455
      Begin VB.TextBox Text1 
         Height          =   2175
         Index           =   4
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   5
         Text            =   "Form2.frx":0000
         Top             =   9360
         Width           =   10095
      End
      Begin VB.TextBox Text1 
         Height          =   2175
         Index           =   3
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   4
         Text            =   "Form2.frx":0006
         Top             =   7080
         Width           =   10095
      End
      Begin VB.TextBox Text1 
         Height          =   2175
         Index           =   2
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   3
         Text            =   "Form2.frx":000C
         Top             =   4800
         Width           =   10095
      End
      Begin VB.TextBox Text1 
         Height          =   2175
         Index           =   1
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   2
         Text            =   "Form2.frx":0012
         Top             =   2520
         Width           =   10095
      End
      Begin VB.TextBox Text1 
         Height          =   2175
         Index           =   5
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Text            =   "Form2.frx":0018
         Top             =   240
         Width           =   10095
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CX!
Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
CX = Y
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Frame1.Move 8, (Frame1.Top + Y) - CX
End Sub

