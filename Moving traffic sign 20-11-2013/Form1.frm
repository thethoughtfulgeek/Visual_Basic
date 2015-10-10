VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   5595
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Start/Stop"
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   600
      Top             =   5160
   End
   Begin VB.Image Image2 
      Height          =   1455
      Left            =   3480
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   3
      Left            =   600
      Picture         =   "Form1.frx":0442
      Top             =   4200
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   2
      Left            =   600
      Picture         =   "Form1.frx":0884
      Top             =   3000
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   600
      Picture         =   "Form1.frx":0CC6
      Top             =   1800
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   600
      Picture         =   "Form1.frx":1108
      Top             =   600
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer1.Enabled = Not (Timer1.Enabled)
End Sub

Private Sub Timer1_Timer()
Static picnum As Integer
picnum = picnum + 1
If picnum > 3 Then
picnum = 0
End If
Image2.Picture = Image1(picnum).Picture
End Sub
