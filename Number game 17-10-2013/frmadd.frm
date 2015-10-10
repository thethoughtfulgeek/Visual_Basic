VERSION 5.00
Begin VB.Form frmadd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Flash Card Addition"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtanswer 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7440
      MaxLength       =   2
      TabIndex        =   9
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "&Next Problem"
      Enabled         =   0   'False
      Height          =   855
      Left            =   3120
      TabIndex        =   8
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
      Height          =   735
      Left            =   6480
      TabIndex        =   7
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label lblnum1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2280
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblnum2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3840
      TabIndex        =   4
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Score:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lblscore 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label lblmessage 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      TabIndex        =   0
      Top             =   2160
      Width           =   6975
   End
End
Attribute VB_Name = "frmadd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sum As Integer
Dim numprob As Integer, numright As Integer

Private Sub cmdexit_Click()
End
End Sub

Private Sub cmdnext_Click()
'generate next addition problem
Dim number1 As Integer
Dim number2 As Integer
txtanswer.Text = ""
lblmessage.Caption = ""
numprob = numprob + 1
number1 = Int(Rnd * 21)
number2 = Int(Rnd * 21)
lblnum1.Caption = Format(number1, "#0")
lblnum2.Caption = Format(number2, "#0")
'find sum
sum = number1 + number2
cmdnext.Enabled = False
txtanswer.SetFocus


End Sub

Private Sub Form_Activate()
Call cmdnext_Click
End Sub

Private Sub Form_Load()
'generate random numbers for addition
Randomize Timer
numprob = 0
numright = 0
End Sub

Private Sub txtanswer_KeyPress(KeyAscii As Integer)
Dim ans As Integer
'check for number only and for return key
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then
Exit Sub
ElseIf KeyAscii = vbKeyReturn Then
'check answer
ans = Val(txtanswer.Text)
If ans = sum Then
numright = numright + 1
lblmessage.Caption = "answer Is correct"
Else
lblmessage.Caption = "Answer is " + Format(sum, "#0")
End If
lblscore.Caption = Format(100 * numright / numprob, "##0")
cmdnext.Enabled = True
cmdnext.SetFocus
Else
KeyAscii = 0
End If
End Sub
