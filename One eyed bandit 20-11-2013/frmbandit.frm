VERSION 5.00
Begin VB.Form frmbandit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "One-Buttoned Bandit"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timspin 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4200
      Top             =   360
   End
   Begin VB.Timer timdone 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   5040
      Top             =   360
   End
   Begin VB.CommandButton cmdexit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdspin 
      Caption         =   "&Spin It"
      Default         =   -1  'True
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Image imgbandit 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Index           =   2
      Left            =   4200
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Image imgbandit 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Index           =   1
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Image imgbandit 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Index           =   0
      Left            =   120
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Image imgchoice 
      Height          =   480
      Index           =   3
      Left            =   1080
      Picture         =   "frmbandit.frx":0000
      Top             =   3840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgchoice 
      Height          =   480
      Index           =   2
      Left            =   1080
      Picture         =   "frmbandit.frx":0442
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgchoice 
      Height          =   480
      Index           =   1
      Left            =   240
      Picture         =   "frmbandit.frx":0884
      Top             =   3840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgchoice 
      Height          =   480
      Index           =   0
      Left            =   240
      Picture         =   "frmbandit.frx":0CC6
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblbank 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2760
      TabIndex        =   3
      Top             =   360
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Bankroll"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Width           =   1305
   End
End
Attribute VB_Name = "frmbandit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bankroll As Integer

Private Sub cmdexit_Click()
MsgBox "You have ended with" + Str(bankroll) + "points.", vbOKOnly, "Game over"
End
End Sub

Private Sub cmdspin_Click()
If bankroll = 0 Then
MsgBox "Out of Cash!!", vbOKOnly + vbCritical, "Game over"
End
End If
bankroll = bankroll - 1
lblbank.Caption = Str(bankroll)
timspin.Enabled = True
timdone.Enabled = True
End Sub

Private Sub Form_Load()
Randomize Timer
bankroll = lblbank.Caption
End Sub

Private Sub timdone_Timer()
Dim p0 As Integer, p1 As Integer, p2 As Integer
Dim winnings As Integer
Const face = 3
timspin.Enabled = False
timdone.Enabled = False
p0 = Int(Rnd * 4)
p1 = Int(Rnd * 4)
p2 = Int(Rnd * 4)
imgbandit(0).Picture = imgchoice(p0).Picture
imgbandit(1).Picture = imgchoice(p1).Picture
imgbandit(2).Picture = imgchoice(p2).Picture
If p0 = face Then
winnings = 1
    If p1 = face Then
    winnings = 2
        If p3 = face Then
        winnings = 10
        End If
    End If
ElseIf p0 = p1 Then
    winnings = 2
    If p1 = p2 Then
    winnings = 4
    End If
End If
bankroll = bankroll + winnings
lblbank.Caption = Str(bankroll)
End Sub

Private Sub timspin_Timer()
imgbandit(0).Picture = imgchoice(Int(4 * Rnd)).Picture
imgbandit(1).Picture = imgchoice(Int(4 * Rnd)).Picture
imgbandit(2).Picture = imgchoice(Int(4 * Rnd)).Picture
End Sub
