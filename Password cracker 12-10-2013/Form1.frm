VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4320
      TabIndex        =   6
      Top             =   1080
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2040
      Top             =   3240
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Reset"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Generate"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Terminate"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "Password Cracked ! Login Successful !"
      Height          =   195
      Left            =   2025
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   2805
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim crackpass As String, code1 As String, code2 As String, code3 As String
Dim password As String

Private Sub Command1_Click()
Timer2.Enabled = False
Text2.Text = "Process terminated"
Text2.BackColor = &HFF&
End Sub

Private Sub Command2_Click()
Timer2.Enabled = True
End Sub

Private Sub Command3_Click()
Text2.Text = ""
Text2.BackColor = vbWhite
End Sub

Private Sub Form_Load()
password = abc
End Sub
Private Sub Timer2_Timer()
Randomize Timer
code1 = Int(Rnd * 255)
code1 = Chr(code1)
code2 = Int(Rnd * 255)
code2 = Chr(code2)
code3 = Int(Rnd * 255)
code3 = Chr(code3)
crackpass = code1 + code2 + code3
If crackpass = password Then
Timer2.Enabled = False
Text1.Text = crackpass
Label1.Visible = True
Else
Text2.Text = "please wait"
End If
End Sub

