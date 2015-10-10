VERSION 5.00
Begin VB.Form frmpizza 
   Caption         =   "Pizza Order"
   ClientHeight    =   6135
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   12090
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   12090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdnew 
      Caption         =   "&New order"
      Height          =   615
      Left            =   9000
      TabIndex        =   18
      Top             =   4800
      Width           =   1815
   End
   Begin VB.OptionButton optwhere 
      Caption         =   "Take away"
      Height          =   495
      Index           =   1
      Left            =   7560
      TabIndex        =   17
      Top             =   3240
      Width           =   1575
   End
   Begin VB.OptionButton optwhere 
      Caption         =   "Eat in"
      Height          =   495
      Index           =   0
      Left            =   4920
      TabIndex        =   16
      Top             =   3240
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Size"
      Height          =   2415
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   3015
      Begin VB.OptionButton optsize 
         Caption         =   "Small"
         Height          =   495
         Index           =   0
         Left            =   840
         TabIndex        =   13
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optsize 
         Caption         =   "Medium"
         Height          =   495
         Index           =   1
         Left            =   840
         TabIndex        =   12
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton optsize 
         Caption         =   "Large"
         Height          =   495
         Index           =   2
         Left            =   840
         TabIndex        =   11
         Top             =   1680
         Width           =   1095
      End
   End
   Begin VB.CheckBox chktop 
      Caption         =   "Tomatoes"
      Height          =   495
      Index           =   5
      Left            =   8160
      TabIndex        =   9
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CheckBox chktop 
      Caption         =   "Green Pepper"
      Height          =   495
      Index           =   4
      Left            =   8160
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CheckBox chktop 
      Caption         =   "Onions"
      Height          =   495
      Index           =   2
      Left            =   5640
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CheckBox chktop 
      Caption         =   "Black Olives"
      Height          =   495
      Index           =   3
      Left            =   8160
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.CheckBox chktop 
      Caption         =   "Mushrooms"
      Height          =   495
      Index           =   1
      Left            =   5640
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CheckBox chktop 
      Caption         =   "Extra Cheese"
      Height          =   495
      Index           =   0
      Left            =   5640
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdexit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   615
      Left            =   6720
      TabIndex        =   3
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton cmdbuild 
      Caption         =   "Build &Pizza"
      Default         =   -1  'True
      Height          =   615
      Left            =   3960
      TabIndex        =   2
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      Caption         =   "Toppings"
      Height          =   2295
      Left            =   4320
      TabIndex        =   1
      Top             =   360
      Width           =   6375
   End
   Begin VB.Frame Frame2 
      Caption         =   "Crust Type"
      Height          =   1215
      Left            =   360
      TabIndex        =   0
      Top             =   3720
      Width           =   3135
      Begin VB.OptionButton optcrust 
         Caption         =   "Thick Crust"
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   15
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton optcrust 
         Caption         =   "Thin Crust"
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmpizza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pizzasize As String
Dim pizzawhere As String
Dim pizzacrust As String
Private Sub cmdbuild_Click()
'this procedure shall build a message box that displays your pizza type
Dim message As String
Dim I As Integer
message = pizzawhere + vbCr
message = message + pizzasize + "Pizza" + vbCr
message = message + pizzacrust + vbCr
For I = 0 To 5
If chktop(I).Value = vbChecked Then message = message + chktop(I).Caption + vbCr
Next I
'no topping selected then cheese only should be displayed on order
If chktop(0).Value = vbUnchecked And chktop(1).Value = vbUnchecked And chktop(2).Value = vbUnchecked And chktop(3).Value = vbUnchecked And chktop(4).Value = vbUnchecked And chktop(5).Value = vbUnchecked Then
message = message + "Cheese only"
End If
MsgBox message, vbOKOnly, "Your Pizza"
End Sub

Private Sub cmdexit_Click()
End
End Sub

Private Sub cmdnew_Click()
Dim a As Integer
pizzasize = "Small"
pizzacrust = "Thin Crust"
pizzawhere = "Eat in"
optsize(0).Value = True
optwhere(0).Value = True
optcrust(0).Value = True
For a = 0 To 5
chktop(a).Value = vbUnchecked
Next a
End Sub

Private Sub Form_Load()
'initialize pizza parameters
pizzasize = "Small"
pizzacrust = "Thin Crust"
pizzawhere = "Eat in"
End Sub

Private Sub optcrust_Click(Index As Integer)
'read pizza crust
pizzacrust = optcrust(Index).Caption
End Sub

Private Sub optsize_Click(Index As Integer)
'read pizza size
pizzasize = optsize(Index).Caption
End Sub

Private Sub optwhere_Click(Index As Integer)
'read pizza eating location
pizzawhere = optwhere(Index).Caption
End Sub
