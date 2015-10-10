VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   8250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   735
      Left            =   4200
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Sports"
      Height          =   615
      Left            =   1080
      TabIndex        =   3
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Computer"
      Height          =   615
      Left            =   1080
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Reading"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Please select your topic/s of interest:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 Then
MsgBox ("you like reading, computer and sports")
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 0 Then
MsgBox ("you like reading and computer")
ElseIf Check1.Value = 1 And Check2.Value = 0 And Check3.Value = 1 Then
MsgBox ("you like reading and sports")
ElseIf Check1.Value = 0 And Check2.Value = 1 And Check3.Value = 1 Then
MsgBox ("you like computer and sports")
ElseIf Check1.Value = 0 And Check2.Value = 0 And Check3.Value = 1 Then
MsgBox ("you like sports only")
ElseIf Check1.Value = 0 And Check2.Value = 1 And Check3.Value = 0 Then
MsgBox ("you like computer only")
ElseIf Check1.Value = 1 And Check2.Value = 0 And Check3.Value = 0 Then
MsgBox ("you like reading only")
Else
MsgBox ("you idiot, go and get some hobby you fool!!")
End If
End Sub
