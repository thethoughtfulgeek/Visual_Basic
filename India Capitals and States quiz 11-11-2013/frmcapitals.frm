VERSION 5.00
Begin VB.Form frmcapitals 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8355
   ClientLeft      =   585
   ClientTop       =   885
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   8040
   Begin VB.TextBox txtanswer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   11
      Top             =   2760
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.CommandButton cmdexit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   7560
      Width           =   1815
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "&Next Question"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   6600
      Width           =   1815
   End
   Begin VB.TextBox lblcomment 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   2640
      TabIndex        =   2
      Top             =   1800
      Width           =   3495
   End
   Begin VB.Label lblanswer 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   240
      TabIndex        =   10
      Top             =   5640
      Width           =   7335
   End
   Begin VB.Label lblanswer 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   4680
      Width           =   7335
   End
   Begin VB.Label lblanswer 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   3720
      Width           =   7335
   End
   Begin VB.Label lblanswer 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   7335
   End
   Begin VB.Label lblgiven 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   7335
   End
   Begin VB.Label lblscore 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3540
      TabIndex        =   5
      Top             =   6840
      Width           =   825
   End
   Begin VB.Label lblheadanswer 
      AutoSize        =   -1  'True
      Caption         =   "Capital:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   360
      TabIndex        =   1
      Top             =   1920
      Width           =   1770
   End
   Begin VB.Label lblheadgiven 
      AutoSize        =   -1  'True
      Caption         =   "State:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1350
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnufilenew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnufilebar 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuoptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuoptionscapitals 
         Caption         =   "Name &Capitals"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuoptionsstates 
         Caption         =   "Name &States"
      End
      Begin VB.Menu mnupotionsbar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuoptionsmc 
         Caption         =   "&Multiple Choice answers"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuoptionstype 
         Caption         =   "&Type in answers"
      End
   End
End
Attribute VB_Name = "frmcapitals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim numans As Integer, correctanswer As Integer, numcorrect As Integer
Dim state(27) As String, capital(27) As String

Private Sub cmdexit_Click()
End
End Sub

Private Sub Form_Activate()
'this function is called every time form becomes the active window
Call mnufilenew_Click
End Sub

Private Sub Form_Load()
 Randomize Timer
 'load state/capital arrays
 state(0) = "Andhra Pradesh": capital(0) = "Hyderabad"
 ': sign indicates two different lines which are written side by side (cascading)
 state(1) = "Arunachal Pradesh": capital(1) = "Itanagar"
 state(2) = "Assam": capital(2) = "Dispur"
 state(3) = "Bihar": capital(3) = "Patna"
 state(4) = "Chhatisgarh": capital(4) = "Raipur"
 state(5) = "Goa": capital(5) = "Panjim"
 state(6) = "Gujarat": capital(6) = "Gandhinagar"
 state(7) = "Haryana": capital(7) = "Chandigarh"
 state(8) = "Himachal Pradesh": capital(8) = "Shimla"
 state(9) = "Jammu and Kashmir": capital(9) = "Srinagar"
 state(10) = "Jharkhand": capital(10) = "Ranchi"
 state(11) = "Karnataka": capital(11) = "Bangalore"
 state(12) = "Kerala": capital(12) = "Thiruvanthapuram"
 state(13) = "Madhya Pradesh": capital(13) = "Bhopal"
 state(14) = "Maharashtra": capital(14) = "Mumbai"
 state(15) = "Manipur": capital(15) = "Imphal"
 state(16) = "Meghalaya": capital(16) = "Shillong"
 state(17) = "Mizoram": capital(17) = "Aizawl"
 state(18) = "Nagaland": capital(18) = "Kohima"
 state(19) = "Orissa": capital(19) = "Bhubaneshwar"
 state(20) = "Punjab": capital(20) = "Chandigarh"
 state(21) = "Rajasthan": capital(21) = "Jaipur"
 state(22) = "Sikkim": capital(22) = "Gangtok"
 state(23) = "Tamil Nadu": capital(23) = "Chennai"
 state(24) = "Tripura": capital(24) = "Agartala"
 state(25) = "Uttar Pradesh": capital(25) = "Lucknow"
 state(26) = "Uttarakhand": capital(26) = "Dehradun"
 state(27) = "West Bengal": capital(27) = "Kolkatta"
 End Sub

Private Sub lblanswer_Click(Index As Integer)
'check multiple choice answers
Dim iscorrect As Integer
'if already answered exit
If cmdnext.Enabled = True Then
Exit Sub
iscorrect = 0
If mnuoptionscapitals.Checked = True Then
    If lblanswer(Index).Caption = capital(correctanswer) Then
        iscorrect = 1
Else
    If lblanswer(Index).Caption = state(correctanswer) Then
    iscorrect = 1
End If
Call update_score(iscorrect)
End If
End If
End If
End Sub

Private Sub mnufileexit_Click()
Call cmdexit_Click
End Sub

Private Sub mnufilenew_Click()
'Reset the score and start again
    numans = 0
    numcorrect = 0
    lblscore.Caption = "0%"
    lblcomment.Text = ""
    cmdnext.Enabled = False
    Call next_question(correctanswer)
    End Sub

Private Sub mnuoptionscapitals_Click()
'setup that provides capitals and state quiz based on given state
mnuoptionsstates.Checked = False
mnuoptionscapitals.Checked = True
lblheadgiven.Caption = "State:"
lblheadanswer.Caption = "Capital:"
Call mnufilenew_Click
End Sub

Private Sub mnuoptionsmc_Click()
'setup for MCQ answers
Dim i As Integer
mnuoptionsmc.Checked = True
mnuoptionstype.Checked = False
For i = 0 To 3
lblanswer(i).Visible = True
Next i
txtanswer.Visible = False
Call mnufilenew_Click
End Sub

Private Sub mnuoptionsstates_Click()
'setup that provides capitals and state quiz based on given capital
mnuoptionsstates.Checked = True
mnuoptionscapitals.Checked = False
lblheadgiven.Caption = "Capital:"
lblheadanswer.Caption = "State:"
Call mnufilenew_Click
End Sub

Private Sub mnuoptionstype_Click()
'setup for typein answers
Dim i As Integer
mnuoptionsmc.Checked = False
mnuoptionstype.Checked = True
For i = 0 To 3
lblanswer(i).Visible = False
Next i
txtanswer.Visible = True
Call mnufilenew_Click
End Sub

Private Sub next_question(answer As Integer)
Dim vused(27) As Integer, i As Integer, j As Integer
Dim Index(3)
lblcomment.Text = ""
numans = numans + 1
'generate the next question based on selected options
answer = Int(Rnd * 27) + 1
If mnuoptionscapitals.Checked = True Then
lblgiven.Caption = state(answer)
Else
lblgiven.Caption = capital(answer)
End If
If mnuoptionsmc.Checked = True Then
'we will have multiple choice answers
'Vused array will be used to see which possible choices are used to as possible answers in options
For i = 0 To 27
vused(i) = 0
Next i
'pick 4 different state indices (j) at random
'these are used to setup multiple choice answers
' these selected answers are stored in the index array
i = 0
Do
    Do
        j = Int(Rnd * 27) + 1
'loop until both the below conditions are satisfied.
'i.e Both the below conditions become true
    Loop Until vused(j) = 0 And j <> answer
    vused(j) = 1
    Index(i) = j
    i = i + 1
Loop Until i = 4
'now replace one index at random with correct answer
Index(Int(Rnd * 4)) = answer
'display multiple choice answers in label boxes
For i = 0 To 3
If mnuoptionscapitals.Checked = True Then
lblanswer(i).Caption = capital(Index(i))
Else
lblanswer(i).Caption = state(Index(i))
End If
Next i
If mnuoptionstype.Checked = True Then
txtanswer.Locked = False
txtanswer.Text = ""
txtanswer.SetFocus
End If
End If
End Sub

Private Sub update_score(iscorrect As Integer)
Dim i As Integer
'check if answer is correct
cmdnext.Enabled = True
cmdnext.SetFocus
If iscorrect = 1 Then
numcorrect = numcorrect + 1
lblcomment.Text = "Correct!"
Else
lblcomment.Text = "Wrong answer, Sorry!"
End If
'display correct and update score
If mnuoptionsmc.Checked = True Then
    For i = 0 To 3
    If mnuoptionscapitals.Checked = True Then
        If lblanswer(i).Caption <> capital(correctanswer) Then
            lblanswer(i).Caption = ""
        End If
    ElseIf mnuoptionsstates.Checked = True Then
        If lblanswer(i).Caption <> capital(correctanswer) Then
            lblanswer(i).Caption = ""
        End If
    End If
    Next i
Else
If mnuoptionsstates.Checked = True Then
    txtanswer.Text = state(correctanswer)
Else
    txtanswer.Text = capital(correctanswer)
End If
End If
lblscore.Caption = Format(numcorrect / numans, "##0%")
End Sub

