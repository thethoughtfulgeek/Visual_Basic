VERSION 5.00
Begin VB.Form frmsavings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Savings account"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   9540
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdreset 
      Caption         =   "&Reset"
      Height          =   495
      Left            =   5880
      TabIndex        =   10
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   3840
      TabIndex        =   9
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdcalculate 
      Caption         =   "&Calculate"
      Height          =   495
      Left            =   3840
      TabIndex        =   8
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txtfinal 
      Height          =   495
      Left            =   5640
      TabIndex        =   7
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtmonths 
      Height          =   495
      Left            =   5640
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtinterest 
      Height          =   495
      Left            =   5640
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtdeposit 
      Height          =   495
      Left            =   5640
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Final Balance"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "No. of months"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Yearly Interest"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Monthly Deposit"
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmsavings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const vbkeydecpt = 46
'there is no constant for ascii value of decimal point. We define it
'using a custom-made constant called vbkeydecpt
Dim interest As Single
Dim months As Single
Dim final As Single
Dim deposit As Single

Private Sub cmdcalculate_Click()
Dim intrate As Single
'read values from textboxes
deposit = Val(txtdeposit.Text)
months = Val(txtmonths.Text)
interest = Val(txtinterest.Text)
intrate = interest / 100
'compute final value and put it in text box
final = deposit * ((1 + intrate) ^ months - 1) / intrate
txtfinal.Text = Format(final, "#####0.00")
End Sub

Private Sub cmdexit_Click()
End
End Sub

Private Sub cmdreset_Click()
txtdeposit.Text = ""
txtmonths.Text = ""
txtinterest.Text = ""
txtfinal.Text = ""
End Sub

Private Sub txtdeposit_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbkeydecpt Or KeyAscii = vbKeyBack Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub

Private Sub txtinterest_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbkeydecpt Or KeyAscii = vbKeyBack Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub

Private Sub txtmonths_KeyPress(KeyAscii As Integer)
'concept of keytrapping
'only it will read the required keys. it will not accept any other keys
'only allow numbers, decimal point and backspace key
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbkeydecpt Or KeyAscii = vbKeyBack Then
Exit Sub
Else
KeyAscii = 0
Beep
'make a beep sound
End If
End Sub
