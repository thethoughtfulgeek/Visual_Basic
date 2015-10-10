VERSION 5.00
Begin VB.Form frmstats 
   Caption         =   "Mean and Standard deviation"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtinput 
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
      Height          =   615
      Left            =   2400
      TabIndex        =   6
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton cmdcompute 
      Caption         =   "&Compute"
      Height          =   615
      Left            =   2400
      TabIndex        =   5
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "&New Sequence"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton cmdaccept 
      Caption         =   "&Accept Number"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lblnumber 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   615
      Left            =   2565
      TabIndex        =   11
      Top             =   360
      Width           =   1125
   End
   Begin VB.Label lblstddev 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   2400
      TabIndex        =   10
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label lblmean 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   2400
      TabIndex        =   9
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Number of Values"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Standard Deviation"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Mean"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Enter Number"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "frmstats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const vbkeydecpt = 46
Const vbkeyminus = 45
Dim sumx2 As Single
Dim sumx As Single
Dim numvalues As Integer


Private Sub cmdaccept_Click()
Dim value As Single
txtinput.SetFocus
numvalues = numvalues + 1
lblnumber.Caption = Str(numvalues)
value = Val(txtinput.Text)
sumx = sumx + value
sumx2 = sumx2 + value ^ 2
txtinput.Text = ""
End Sub

Private Sub cmdcompute_Click()
Dim mean As Single
Dim stddev As Single
txtinput.SetFocus
If numvalues < 2 Then
Beep
End If
mean = sumx / numvalues
lblmean.Caption = Str(mean)
stddev = Sqr((numvalues * sumx2 - sumx ^ 2) / (numvalues * (numvalues - 1)))
lblstddev.Caption = Str(stddev)
End Sub

Private Sub cmdexit_Click()
End
End Sub

Private Sub cmdnew_Click()
txtinput.SetFocus
numvalues = 0
lblnumber.Caption = "0"
txtinput.Text = ""
lblmean.Caption = ""
lblstddev.Caption = ""
sumx = 0
sumx2 = 0
End Sub
Private Sub txtinput_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbkeydecpt Or KeyAscii = vbkeyminus Or KeyAscii = vbKeyBack) Then
Exit Sub
ElseIf KeyAscii = vbKeyReturn Then
Call cmdaccept_Click
Else
KeyAscii = 0
End If
End Sub
