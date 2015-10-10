VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Motor Database"
   ClientHeight    =   9390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15060
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   15060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdchart 
      Caption         =   "S&how Chart"
      Height          =   495
      Left            =   12960
      TabIndex        =   4
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton cmdexit2 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   495
      Left            =   12960
      TabIndex        =   3
      Top             =   8160
      Width           =   1455
   End
   Begin VB.TextBox txtsheet 
      Height          =   495
      Left            =   12960
      MaxLength       =   5
      TabIndex        =   2
      Top             =   6120
      Width           =   1455
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   5295
      Left            =   120
      OleObjectBlob   =   "Form2.frx":0000
      TabIndex        =   1
      Top             =   0
      Width           =   14775
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3735
      Left            =   240
      TabIndex        =   0
      Top             =   5280
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   6588
      _Version        =   393216
      ScrollTrack     =   -1  'True
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim j As Integer, count1 As Integer

Private Sub cmdchart_Click()
Dim x(1 To 50) As Variant
MSFlexGrid1.Rows = count1
MSChart1.chartType = VtChChartType2dLine
MSChart1.RowCount = count1
MSChart1.ColumnCount = 1
For j = 1 To MSFlexGrid1.Rows - 1
    MSFlexGrid1.Col = 1
    MSFlexGrid1.Row = j
    x(j) = MSFlexGrid1.Text
    MSChart1.Column = 1
    MSChart1.Row = j
    MSChart1.Data = x(j)
    MSChart1.RowLabel = j
Next j
MSChart1.ShowLegend = True
MSChart1.Visible = True
End Sub

Private Sub cmdexit2_Click()
End
End Sub

Private Sub Form_Load()
    MSChart1.Visible = False
    MSFlexGrid1.Rows = 100
    MSFlexGrid1.Cols = 10
    MSFlexGrid1.FixedRows = 1
    MSFlexGrid1.FixedCols = 1
    MSFlexGrid1.AllowUserResizing = flexResizeBoth
    MSFlexGrid1.WordWrap = True
    MSFlexGrid1.Row = 0
    count1 = 1
    For i = 1 To MSFlexGrid1.Cols - 1
        MSFlexGrid1.Col = i
        MSFlexGrid1.Text = Chr(Asc("A") - 1 + i)
' this prints alphabets on row 0
    Next i
    MSFlexGrid1.Col = 0
    For i = 1 To MSFlexGrid1.Rows - 1
        MSFlexGrid1.Row = i
        MSFlexGrid1.Text = Str(i)
' this prints numbers on column 0
    Next i
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Col = 0
    MSFlexGrid1.Visible = True
End Sub

Public Sub transfer_vel()
    count1 = count1 + 1
    MSFlexGrid1.Col = 1
    MSFlexGrid1.Row = MSFlexGrid1.Row + 1
    MSFlexGrid1.Text = txtsheet.Text
End Sub

