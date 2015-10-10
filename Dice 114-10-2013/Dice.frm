VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   6060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Roll Dice"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   0
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H80000004&
      Height          =   255
      Index           =   6
      Left            =   3120
      Shape           =   3  'Circle
      Top             =   1920
      Width           =   255
   End
   Begin VB.Shape Shape2 
      Height          =   255
      Index           =   5
      Left            =   3120
      Shape           =   3  'Circle
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape Shape2 
      Height          =   255
      Index           =   4
      Left            =   3120
      Shape           =   3  'Circle
      Top             =   720
      Width           =   255
   End
   Begin VB.Shape Shape2 
      Height          =   255
      Index           =   3
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape Shape2 
      Height          =   255
      Index           =   2
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   1920
      Width           =   255
   End
   Begin VB.Shape Shape2 
      Height          =   255
      Index           =   1
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape Shape2 
      Height          =   255
      Index           =   0
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   720
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   1935
      Left            =   1560
      Shape           =   5  'Rounded Square
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Randomize Timer
n = Int(1 + Rnd * 6)
For i = 0 To 6
Shape2(i).Visible = False
Next
If n = 1 Then
Shape2(3).Visible = True
ElseIf n = 2 Then
Shape2(1).Visible = True
Shape2(5).Visible = True
ElseIf n = 3 Then
Shape2(1).Visible = True
Shape2(3).Visible = True
Shape2(5).Visible = True
ElseIf n = 4 Then
Shape2(0).Visible = True
Shape2(4).Visible = True
Shape2(2).Visible = True
Shape2(6).Visible = True
ElseIf n = 5 Then
Shape2(0).Visible = True
Shape2(4).Visible = True
Shape2(2).Visible = True
Shape2(6).Visible = True
Shape2(3).Visible = True
ElseIf n = 6 Then
Shape2(0).Visible = True
Shape2(4).Visible = True
Shape2(2).Visible = True
Shape2(6).Visible = True
Shape2(1).Visible = True
Shape2(5).Visible = True

End If

End Sub
