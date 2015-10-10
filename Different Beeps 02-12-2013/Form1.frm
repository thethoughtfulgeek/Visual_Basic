VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long
Private Const MB_ICONASTERISK = &H40&
Private Const MB_ICONEXCLAMATION = &H30&
Private Const MB_ICONHAND = &H10&
Private Const MB_ICONINFORMATION = MB_ICONASTERISK
Private Const MB_ICONSTOP = MB_ICONHAND
Private Const MB_ICONQUESTION = &H20&
Private Const MB_OK = &H0&


Private Sub Command1_Click()
Dim beeptype As Long, rtnvalue As Long
Select Case Val(Text1.Text)
Case 0
    beeptype = MB_OK
Case 1
    beeptype = MB_ICONINFORMATION
Case 2
    beeptype = MB_ICONEXCLAMATION
Case 3
    beeptype = MB_ICONQUESTION
Case 4
    beeptype = MB_ICONSTOP
End Select
rtnvalue = MessageBeep(beeptype)
MsgBox "This is a test", beeptype, "Beep Test"
End Sub
