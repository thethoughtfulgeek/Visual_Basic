VERSION 5.00
Begin VB.Form frmplot 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plotting Examples"
   ClientHeight    =   6630
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   8430
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picplot 
      BackColor       =   &H00FFFFFF&
      Height          =   6375
      Left            =   120
      ScaleHeight     =   6315
      ScaleWidth      =   8115
      TabIndex        =   0
      Top             =   120
      Width           =   8175
   End
   Begin VB.Menu mnuplot 
      Caption         =   "&Plot"
      Begin VB.Menu mnuplotline 
         Caption         =   "&Line Chart"
      End
      Begin VB.Menu mnuplotbar 
         Caption         =   "&Bar Chart"
      End
      Begin VB.Menu mnuplotspiral 
         Caption         =   "&Spiral Chart"
      End
      Begin VB.Menu mnuplotdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuplotexit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmplot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim N As Integer
Dim X(199) As Single
Dim Y(199) As Single
Dim YD(199) As Single

Private Sub Form_Load()
'loading all arrays with points to plot
Dim i As Integer
Const pi = 3.14
N = 200
For i = 0 To N - 1
X(i) = i
Y(i) = Exp(-0.01 * i) * Sin(pi * i / 10)
YD(i) = Exp(-0.01 * i) * (pi * Cos(pi * i / 10) / 10 - 0.01 * Sin(pi * i / 10))
Next i
End Sub

Private Sub mnuplotbar_Click()
Call barchart(picplot, N, X, Y)
End Sub

Private Sub mnuplotexit_Click()
End
End Sub

Private Sub mnuplotline_Click()
Call linechart(picplot, N, X, Y)
End Sub

Private Sub mnuplotspiral_Click()
Call spiralchart(picplot, N, X, YD)
End Sub

Private Sub barchart(objectname As Control, N As Integer, X() As Single, Y() As Single)
Dim xmin As Single, ymin As Single
Dim xmax As Single, ymax As Single
Dim i As Integer
xmin = X(0)
ymin = Y(0)
xmax = X(0)
ymax = Y(0)
For i = 0 To N - 1
    If X(i) < xmin Then xmin = X(i)
    If X(i) > xmax Then xmax = X(i)
    If Y(i) < ymin Then ymin = Y(i)
    If Y(i) > ymax Then ymax = Y(i)
Next i
ymin = (1 - 0.05 * (Sgn(ymin))) * ymin
'extend ymin by 5 percent
ymax = (1 - 0.05 * (Sgn(ymax))) * ymax
'extend ymax by 5 percent
picplot.Cls
picplot.Scale (xmin, ymax)-(xmax, ymin)
picplot.PSet (0, 0)
For i = 0 To N - 1
    picplot.Line (X(i), 0)-(X(i), Y(i)), vbRed
Next i
End Sub


Public Sub linechart(objectname As Control, N As Integer, X() As Single, Y() As Single)
Dim i As Integer
Dim xmin As Single, ymin As Single
Dim xmax As Single, ymax As Single
xmin = X(0): ymin = Y(0)
xmax = X(0): ymax = Y(0)
For i = 0 To N - 1
    If X(i) < xmin Then xmin = X(i)
    If X(i) > xmax Then xmax = X(i)
    If Y(i) < ymin Then ymin = Y(i)
    If Y(i) > ymax Then ymax = Y(i)
Next i
ymin = (1 - 0.05 * (Sgn(ymin))) * ymin
'extend ymin by 5 percent
ymax = (1 - 0.05 * (Sgn(ymax))) * ymax
'extend ymax by 5 percent
picplot.Cls
picplot.Scale (xmin, ymax)-(xmax, ymin)
picplot.PSet (0, 0)
For i = 0 To N - 1
    picplot.Line -(X(i), Y(i)), vbGreen
Next i
End Sub
