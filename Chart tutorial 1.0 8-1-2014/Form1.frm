VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9375
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   14250
   LinkTopic       =   "Form1"
   ScaleHeight     =   9375
   ScaleWidth      =   14250
   StartUpPosition =   3  'Windows Default
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   4215
      Left            =   240
      OleObjectBlob   =   "Form1.frx":0000
      TabIndex        =   0
      ToolTipText     =   "yoo yoo "
      Top             =   240
      Width           =   12255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim x(1 To 3, 1 To 6) As Variant
x(1, 2) = "Rice"
x(1, 3) = "Corn"
x(1, 4) = "Lentils"
x(1, 5) = "Wheat"
x(1, 6) = "Rye"
x(2, 1) = "January"
x(3, 1) = "February"
x(2, 2) = 2
x(2, 3) = 3
x(2, 4) = 4
x(2, 5) = 5
x(2, 6) = 6
x(3, 2) = 4
x(3, 3) = 6
x(3, 4) = 8
x(3, 5) = 10
x(3, 6) = 12
MSChart1.ChartData = x
MSChart1.ShowLegend = True
End Sub
