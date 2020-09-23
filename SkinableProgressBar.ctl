VERSION 5.00
Begin VB.UserControl SkinableProgressBar 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3600
      ScaleHeight     =   465
      ScaleWidth      =   225
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox BG 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   3585
      TabIndex        =   1
      Top             =   480
      Width           =   3615
      Begin VB.Image BG1 
         Height          =   495
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3615
      End
   End
   Begin VB.PictureBox FG 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   3585
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.Image FG1 
         Height          =   495
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3615
      End
   End
End
Attribute VB_Name = "SkinableProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Made by HardStream Productions ©
' This is release 2.0
' No timers!
' Better formatted
' Faster code
'
' EnJoY It!
'
' SkinableProgressBar © HardStream Productions 12-7-2004

Option Explicit

Public Value As String

Private Sub UserControl_Initialize()
BG.Top = 0
BG.Left = -3600
SetValue 0
BG1.Left = BG.Left
BG1.Top = BG.Top
FG1.Left = FG.Left
FG1.Top = FG.Top
BG1.Width = BG.Width
BG1.Height = BG.Height
FG1.Width = FG.Width
FG1.Height = FG.Height
End Sub

Private Sub UserControl_Resize()
BG.Width = UserControl.Width
BG.Height = UserControl.Height
FG.Width = UserControl.Width
FG.Height = UserControl.Height
Picture1.Left = BG.Width - 10
FG1.Left = FG.Left
FG1.Top = FG.Top
BG1.Width = BG.Width
BG1.Height = BG.Height
FG1.Width = FG.Width
FG1.Height = FG.Height
Value = 0
End Sub

Public Sub SetValue(Number As String)
If Number = "" Then
MsgBox "No Number", vbCritical, "Error"
Exit Sub
End If
BG.Left = FG.Width / 100 * Number
Value = -1 + 1 + Number
If Value >= 100 Then
Value = 100
End If
FG1.Left = FG.Left
FG1.Top = FG.Top
FG1.Width = FG.Width
FG1.Height = FG.Height
BG1.Width = BG.Width
BG1.Height = BG.Height
End Sub

Public Sub AddValue(Number As String)
If Number = "" Then
MsgBox "No Number", vbCritical, "Error"
Exit Sub
End If
BG.Left = BG.Left + FG.Width / 100 * Number
Value = -1 + 1 + Value + Number
If Value >= 100 Then
Value = 100
End If
FG1.Left = FG.Left
FG1.Top = FG.Top
FG1.Width = FG.Width
FG1.Height = FG.Height
BG1.Width = BG.Width
BG1.Height = BG.Height
End Sub

Public Sub SetColorFG(Color As String)
If Color = "" Then
MsgBox "No Color", vbCritical, "Error"
Exit Sub
End If
FG.BackColor = Color
End Sub

Public Sub SetColorBG(Color As String)
If Color = "" Then
MsgBox "No Color", vbCritical, "Error"
Exit Sub
End If
BG.BackColor = Color
End Sub

Public Sub SetImageFG(Image As String)
If Image = "" Then
MsgBox "No Image", vbCritical, "Error"
Exit Sub
End If
FG1.Picture = LoadPicture(Image)
End Sub

Public Sub SetImageBG(Image As String)
If Image = "" Then
MsgBox "No Image", vbCritical, "Error"
Exit Sub
End If
BG1.Picture = LoadPicture(Image)
End Sub

Public Sub SetPictureFG(Picture As String)
If Picture = "" Then
MsgBox "No Picture", vbCritical, "Error"
Exit Sub
End If
FG.Picture = LoadPicture(Picture)
End Sub

Public Sub SetPictureBG(Picture As String)
If Picture = "" Then
MsgBox "No Picture", vbCritical, "Error"
Exit Sub
End If
BG.Picture = LoadPicture(Picture)
End Sub


