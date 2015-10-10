VERSION 5.00
Begin VB.Form frmflight 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Flight Planner"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdexit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   495
      Left            =   3720
      TabIndex        =   7
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdassign 
      Caption         =   "&Assign"
      Default         =   -1  'True
      Height          =   495
      Left            =   1680
      TabIndex        =   6
      Top             =   3600
      Width           =   1335
   End
   Begin VB.ComboBox cbomeal 
      Height          =   1545
      Left            =   5280
      Style           =   1  'Simple Combo
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.ComboBox cboseat 
      Height          =   315
      Left            =   3120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
   Begin VB.ListBox lstcities 
      Height          =   2205
      Left            =   360
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Meal Preference"
      Height          =   495
      Left            =   5520
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Seat Preference"
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Destination City"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmflight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdassign_Click()
'build message box that gives the message
Dim message As String
message = "Destination:" + lstcities.Text + vbCr
message = message + "Seat Location:" + cboseat.Text + vbCr
message = message + "Food preference:" + cbomeal.Text + vbCr
MsgBox message, vbOKOnly + vbInformation, "Your assignment"
End Sub

Private Sub cmdexit_Click()
End
End Sub

Private Sub Form_Load()

'add list of cities to list box
lstcities.Clear
lstcities.AddItem "Toronto"
lstcities.AddItem "Pittsburgh"
lstcities.AddItem "santa barbara"
lstcities.AddItem "los angeles"
lstcities.AddItem "boston"
lstcities.AddItem "worcester"
lstcities.AddItem " houston"
lstcities.AddItem "san diego"
lstcities.AddItem "new york"
lstcities.AddItem "berkeley"
lstcities.AddItem "illinois"
lstcities.AddItem "urbana champain"
lstcities.AddItem "omaha"
lstcities.AddItem "bakersfield"
lstcities.AddItem "san franscisco"
lstcities.ListIndex = 0

'add seat types to first combo box
cboseat.AddItem "Aisle"
cboseat.AddItem "Middle"
cboseat.AddItem "window"
cboseat.ListIndex = 0

'add meal types to second combo box
cbomeal.AddItem "veg only"
cbomeal.AddItem "Jain"
cbomeal.AddItem "Non-veg"
cbomeal.AddItem "Fruit Plate"
cbomeal.AddItem "no preference"
End Sub
