VERSION 5.00
Begin VB.Form tmpform 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Temperature Conversion"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdexit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   2520
      Width           =   1455
   End
   Begin VB.VScrollBar vsbtemp 
      Height          =   2295
      LargeChange     =   10
      Left            =   1920
      Max             =   -60
      Min             =   120
      TabIndex        =   0
      Top             =   480
      Value           =   32
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFF80&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      FillColor       =   &H000000FF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   2775
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lbltempc 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3480
      TabIndex        =   4
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label lbltempf 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "32"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   750
      TabIndex        =   3
      Top             =   1560
      Width           =   435
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "Celsius"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   360
      Left            =   3195
      TabIndex        =   2
      Top             =   720
      Width           =   945
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "Fahrenheit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   360
      Left            =   165
      TabIndex        =   1
      Top             =   720
      Width           =   1470
   End
End
Attribute VB_Name = "tmpform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tempf As Integer
Dim tempc As Integer
Private Sub cmdexit_Click()
End
End Sub
Private Sub vsbtemp_Change()
tempf = vsbtemp.Value
Call showtemps
End Sub
Private Sub vsbtemp_Scroll()
'read F and convert to C
Call vsbtemp_Change
End Sub

Private Sub showtemps()
lbltempf.Caption = Str(tempf)
tempc = degf_to_degc(tempf)
lbltempc.Caption = Str(tempc)
End Sub
