VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calendar"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   15180
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timdisplay 
      Interval        =   1000
      Left            =   3000
      Top             =   3600
   End
   Begin VB.Label lbltime 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "00:00:00 PM"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1905
      TabIndex        =   4
      Top             =   2280
      Width           =   2685
   End
   Begin VB.Label lblmonth 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "March"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   10275
      TabIndex        =   3
      Top             =   480
      Width           =   1425
   End
   Begin VB.Label lblyear 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "1998"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1665
      Left            =   9330
      TabIndex        =   2
      Top             =   6960
      Width           =   3195
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   10200
      TabIndex        =   1
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label lblday 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Sunday"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2160
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub timdisplay_Timer()
Dim today As Variant
today = Now
lblday.Caption = Format(today, "dddd")
lblmonth.Caption = Format(today, "mmmm")
lblyear.Caption = Format(today, "yyyy")
lbltime.Caption = Format(today, "hh:mm:ss ampm")
lbldate.Caption = Format(today, "dd")

End Sub
