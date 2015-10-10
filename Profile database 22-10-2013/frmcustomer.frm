VERSION 5.00
Begin VB.Form frmcustomer 
   Caption         =   "Customer Profile"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10155
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   10155
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   8040
      TabIndex        =   22
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton cmdshow 
      Caption         =   "&Show Profile"
      Height          =   495
      Left            =   8040
      TabIndex        =   21
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "&New Profile"
      Height          =   495
      Left            =   8040
      TabIndex        =   20
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "Athletic Level"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   4200
      TabIndex        =   15
      Top             =   3240
      Width           =   2775
      Begin VB.OptionButton optlevel 
         Caption         =   "Extreme"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   19
         Top             =   1920
         Width           =   1095
      End
      Begin VB.OptionButton optlevel 
         Caption         =   "Advanced"
         Height          =   495
         Index           =   1
         Left            =   360
         TabIndex        =   18
         Top             =   1320
         Width           =   1095
      End
      Begin VB.OptionButton optlevel 
         Caption         =   "Intermediate"
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   17
         Top             =   840
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optlevel 
         Caption         =   "Beginner"
         Height          =   495
         Index           =   3
         Left            =   360
         TabIndex        =   16
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Activities"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   240
      TabIndex        =   8
      Top             =   3120
      Width           =   3615
      Begin VB.CheckBox chkact 
         Caption         =   "Running"
         Height          =   435
         Index           =   5
         Left            =   1920
         TabIndex        =   14
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox chkact 
         Caption         =   "Walking"
         Height          =   495
         Index           =   4
         Left            =   1920
         TabIndex        =   13
         Top             =   720
         Width           =   1455
      End
      Begin VB.CheckBox chkact 
         Caption         =   "Biking"
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   12
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CheckBox chkact 
         Caption         =   "Swimming"
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   11
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox chkact 
         Caption         =   "Skiing"
         Height          =   495
         Index           =   1
         Left            =   360
         TabIndex        =   10
         Top             =   840
         Width           =   1335
      End
      Begin VB.CheckBox chkact 
         Caption         =   "Skating"
         Height          =   495
         Index           =   0
         Left            =   360
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.ComboBox cbocity 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   3360
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   7
      Text            =   "City of Residence"
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sex"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   2295
      Begin VB.OptionButton optsex 
         Caption         =   "Female"
         Height          =   615
         Index           =   1
         Left            =   360
         TabIndex        =   6
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton optsex 
         Caption         =   "Male"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.TextBox txtage 
      Height          =   495
      Left            =   7560
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtname 
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Age"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "SansSerif"
         Size            =   18
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmcustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim activity As String

Private Sub cmdexit_Click()
End
End Sub

Private Sub cmdnew_Click()
Dim i As Integer
'blank out name and age and reset all check boxes
txtname.Text = ""
txtage.Text = ""
For i = 0 To 5
chkact(i).Value = vbUnchecked
Next i
End Sub

Private Sub cmdshow_Click()
Dim noact As Integer, i As Integer
Dim msg As String, pronoun As String
'check to make sure name is entered
If txtname.Text = "" Then
MsgBox "The profile requires a name", vbOKOnly + vbCritical, "No name entered"
Exit Sub
End If
'check to make sure age is entered
If txtage.Text = "" Then
MsgBox "The profile requires an age", vbOKOnly + vbCritical, "No age entered"
Exit Sub
End If
'put together custom profile message
msg = txtname.Text + "is" + Str(txtage.Text) + "years old." + vbCr
'str() converts the number to a string. For adding 2 strings one needs to convert the number
'into a string
If optsex(0).Value = True Then pronoun = "He"
Else
pronoun = "she"
msg = msg + "pronoun" + "lives in" + cbocity.Text + "." + vbCr
msg = msg + pronoun + "is a"
If optlevel(3).Value = False Then
msg = msg + "n"
Else
msg = msg + ""
msg = msg + activity + "level athlete." + vbCr
noact = 0
For i = 0 To 5
If chkact(i).Value = vbChecked Then
noact = noact + 1
Next i
If noact > 0 Then
msg = msg + "activities include:" + vbCr
For i = 0 To 5
If chkact(i).Value = vbChecked Then
msg = msg + String$(10, 32) + chkact(i).Caption + vbCr
Next i
Else
msg = msg + vbCr
End If
MsgBox msg, vbOKOnly, "Customer Profile"
End Sub

Private Sub Form_Load()
'Load combo box with potential city names
cbocity.AddItem "Seattle"
cbocity.Text = "Seattle"
cbocity.AddItem "Bellevue"
cbocity.AddItem "Kirkland"
cbocity.AddItem "Everett"
cbocity.AddItem "Mercer Island"
cbocity.AddItem "Renton"
cbocity.AddItem "Issaquah"
cbocity.AddItem "Kent"
cbocity.AddItem "Bothell"
cbocity.AddItem "Tukwila"
cbocity.AddItem "West Seattle"
cbocity.AddItem "Edmonds"
cbocity.AddItem "Tacoma"
cbocity.AddItem "Federal Way"
cbocity.AddItem "Burien"
cbocity.AddItem "SeaTac"
cbocity.AddItem "Woodinville"
activity = "intermediate"
End Sub

Private Sub optlevel_Click(Index As Integer)
'Determine activity level
Select Case Index
Case 0
activity = "extreme"
Case 1
activity = "advanced"
Case 2
activity = "intermediate"
Case 3
activity = "beginner"
End Select
End Sub

Private Sub txtage_KeyPress(KeyAscii As Integer)
'Only allow numbers for age
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Then
Exit Sub
Else
KeyAscii = 0
End If
End Sub
