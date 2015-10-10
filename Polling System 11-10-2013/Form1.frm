VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10050
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   10050
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   8400
      TabIndex        =   27
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   8400
      TabIndex        =   26
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   8400
      TabIndex        =   25
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   8400
      TabIndex        =   24
      Top             =   3480
      Width           =   855
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   8400
      TabIndex        =   23
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Vote"
      Height          =   615
      Left            =   2040
      TabIndex        =   22
      Top             =   6120
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   7320
      TabIndex        =   21
      Top             =   5160
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   5760
      TabIndex        =   20
      Text            =   "Total Votes= "
      Top             =   5160
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   7320
      TabIndex        =   19
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   7320
      TabIndex        =   18
      Top             =   3480
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   7320
      TabIndex        =   17
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   7320
      TabIndex        =   16
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   7320
      TabIndex        =   15
      Top             =   600
      Width           =   855
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H8000000A&
      Height          =   255
      Left            =   4800
      ScaleHeight     =   195
      ScaleWidth      =   1755
      TabIndex        =   14
      Top             =   4320
      Width           =   1815
   End
   Begin VB.PictureBox Picture4 
      Height          =   255
      Left            =   4800
      ScaleHeight     =   195
      ScaleWidth      =   1755
      TabIndex        =   13
      Top             =   3480
      Width           =   1815
   End
   Begin VB.PictureBox Picture3 
      Height          =   255
      Left            =   4800
      ScaleHeight     =   195
      ScaleWidth      =   1755
      TabIndex        =   12
      Top             =   2520
      Width           =   1815
   End
   Begin VB.PictureBox Picture2 
      Height          =   255
      Left            =   4800
      ScaleHeight     =   195
      ScaleWidth      =   1755
      TabIndex        =   11
      Top             =   1560
      Width           =   1815
   End
   Begin VB.OptionButton Option5 
      Height          =   675
      Left            =   3240
      TabIndex        =   9
      Top             =   4320
      Width           =   1095
   End
   Begin VB.OptionButton Option4 
      Height          =   675
      Left            =   3240
      TabIndex        =   8
      Top             =   3360
      Width           =   1095
   End
   Begin VB.OptionButton Option3 
      Height          =   675
      Left            =   3240
      TabIndex        =   7
      Top             =   2400
      Width           =   1095
   End
   Begin VB.OptionButton Option2 
      Height          =   675
      Left            =   3240
      TabIndex        =   6
      Top             =   1440
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Height          =   675
      Left            =   3240
      TabIndex        =   5
      Top             =   480
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   4800
      ScaleHeight     =   195
      ScaleWidth      =   1755
      TabIndex        =   10
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Bad"
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Satisfactory"
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Good"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Very Good"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Excellent"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim opt1_perc, opt2_perc, opt3_perc, opt4_perc, opt5_perc As Single
Dim total, opt1_tot, opt2_tot, opt3_tot, opt4_tot, opt5_tot As Integer
Private Sub Command1_Click()
If Option1.Value = True Then
opt1_tot = opt1_tot + 1
Text1.Text = opt1_tot
ElseIf Option2.Value = True Then
opt2_tot = opt2_tot + 1
Text2.Text = opt2_tot
ElseIf Option3.Value = True Then
opt3_tot = opt3_tot + 1
Text3.Text = opt3_tot
ElseIf Option4.Value = True Then
opt4_tot = opt4_tot + 1
Text4.Text = opt4_tot
ElseIf Option5.Value = True Then
opt5_tot = opt5_tot + 1
Text5.Text = opt5_tot
End If
total = opt1_tot + opt2_tot + opt3_tot + opt4_tot + opt5_tot
Text7.Text = total

opt1_perc = opt1_tot / total
opt2_perc = opt2_tot / total
opt3_perc = opt3_tot / total
opt4_perc = opt4_tot / total
opt5_perc = opt5_tot / total

Text12.Text = Format(opt1_perc, "percent")
Text11.Text = Format(opt2_perc, "percent")
Text10.Text = Format(opt3_perc, "percent")
Text9.Text = Format(opt4_perc, "percent")
Text8.Text = Format(opt5_perc, "percent")
Picture1.Line (4800, 600)-(6615 * opt1_perc, 345), vbRed, BF
Picture2.Line (4800, 1560)-(6615 * opt2_perc, 1305), vbBlack, BF
Picture3.Line (4800, 2520)-(6615 * opt3_perc, 2265), vbBlue, BF
Picture4.Line (4800, 3480)-(6615 * opt4_perc, 3225), vbGreen, BF
Picture5.Line (4800, 4320)-(6615 * opt5_perc, 4065), vbYellow, BF
End Sub

