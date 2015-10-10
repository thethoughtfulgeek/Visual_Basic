VERSION 5.00
Begin VB.Form frmpassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password Validation"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   11715
   StartUpPosition =   3  'Windows Default
   Tag             =   "Jaineel"
   Begin VB.CommandButton cmdexit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   615
      Left            =   5880
      TabIndex        =   3
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdvalid 
      Caption         =   "&Validate"
      Default         =   -1  'True
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox txtpassword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      IMEMode         =   3  'DISABLE
      Left            =   3600
      PasswordChar    =   "*"
      TabIndex        =   1
      Tag             =   "jaineel"
      Top             =   1320
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Please enter your password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   11055
   End
End
Attribute VB_Name = "frmpassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim incorrect As Integer
Const tot_attempt = 3

Private Sub cmdexit_Click()
End
End Sub

Private Sub cmdvalid_Click()
'this procedure checks the input password
Dim response As Integer
If txtpassword.Text = txtpassword.Tag Then
'if correct, display message box
MsgBox "Correct Password", vbOKOnly + vbExclamation, "Access Granted"
Else
'if incorrect allow to chance to try again
response = MsgBox("Incorrect Password", vbRetryCancel + vbCritical, "Access Denied")
incorrect = incorrect + 1
If incorrect >= tot_attempt Then
MsgBox "you have entered 3 incorrect passwords, program is exiting", vbCritical, "Intruder Alert"
End
End If
If response = vbRetry Then
txtpassword.Text = ""
Else
End
End If
End If
txtpassword.SetFocus
End Sub

Private Sub Form_Activate()
txtpassword.SetFocus
End Sub


