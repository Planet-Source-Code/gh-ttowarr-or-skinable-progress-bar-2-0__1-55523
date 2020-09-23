VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Demo 
   Caption         =   "Demo"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Caption         =   "SetBGPicture"
      Height          =   375
      Left            =   5640
      TabIndex        =   16
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "SetFGPicture"
      Height          =   375
      Left            =   5640
      TabIndex        =   15
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "SetBGImage"
      Height          =   375
      Left            =   4320
      TabIndex        =   14
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "SetFGImage"
      Height          =   375
      Left            =   4320
      TabIndex        =   13
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "SetFGColor"
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SetBGColor"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "AddValue -->"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SetValue -->"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   3360
      Width           =   1215
   End
   Begin ProgressBar.SkinableProgressBar PBar 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   7020
      _ExtentX        =   12383
      _ExtentY        =   661
   End
   Begin MSComDlg.CommonDialog com2 
      Left            =   240
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog com1 
      Left            =   240
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog com4 
      Left            =   840
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog com3 
      Left            =   840
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog com6 
      Left            =   1440
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog com5 
      Left            =   1440
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   2355
      Left            =   120
      Top             =   120
      Width           =   7020
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "HardStream Productions ©"
      Height          =   255
      Left            =   960
      TabIndex        =   12
      Top             =   5640
      Width           =   5295
   End
   Begin VB.Label Label5 
      Caption         =   "programs but we hold the copyright"
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   5280
      Width           =   5295
   End
   Begin VB.Label Label4 
      Caption         =   "this one is the ultimate SkinableProgressBar. You can use it in your"
      Height          =   255
      Left            =   960
      TabIndex        =   10
      Top             =   5040
      Width           =   5295
   End
   Begin VB.Label Label3 
      Caption         =   "make a SkinableProgressBar. We've made a little ProgressBar before but"
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   4800
      Width           =   5295
   End
   Begin VB.Label Label2 
      Caption         =   "I've made this Skinable ProgressBar because a Friend of me wanted to"
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   4560
      Width           =   5295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   7020
   End
End
Attribute VB_Name = "Demo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

Private Sub Command5_Click()
com4.ShowOpen
PBar.SetImageFG com4.FileName
End Sub

Private Sub Command6_Click()
com3.ShowOpen
PBar.SetImageBG com3.FileName
End Sub

Private Sub Command7_Click()
com6.ShowOpen
PBar.SetPictureFG com6.FileName
End Sub

Private Sub Command8_Click()
com5.ShowOpen
PBar.SetPictureBG com5.FileName
End Sub

Private Sub Form_Load()
Demo.Height = 4755
Label1.Caption = PBar.Value & " %"
Image1.Picture = LoadPicture(App.Path & "\logo.bmp")
Demo.Icon = LoadPicture(App.Path & "\logo.ico")
End Sub

Private Sub Command1_Click()
PBar.SetValue Text1.Text
Label1.Caption = PBar.Value & " %"
End Sub

Private Sub Command2_Click()
PBar.AddValue Text2.Text
Label1.Caption = PBar.Value & " %"
End Sub

Private Sub Command3_Click()
com1.ShowColor
PBar.SetColorBG com1.Color
End Sub

Private Sub Command4_Click()
com2.ShowColor
PBar.SetColorFG com2.Color
End Sub

