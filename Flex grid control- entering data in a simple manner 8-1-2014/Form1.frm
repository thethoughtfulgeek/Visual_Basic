VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   11370
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1935
      Left            =   3840
      TabIndex        =   1
      Top             =   600
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   3413
      _Version        =   393216
      Rows            =   7
      Cols            =   7
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   5415
      Left            =   2160
      OleObjectBlob   =   "Form1.frx":0000
      TabIndex        =   0
      Top             =   3480
      Width           =   10215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim items(6) As String
Dim intloopindex As Integer
items(1) = "Item 1"
items(2) = "Item 2"
items(3) = "Item 3"
items(4) = "Item 4"
items(5) = "Item 5"
items(6) = "Total"
For intloopindex = 1 To MSFlexGrid1.Rows - 1
    MSFlexGrid1.Col = 0
    MSFlexGrid1.Row = intloopindex
    MSFlexGrid1.Text = Str(intloopindex)
    MSFlexGrid1.Col = 1
    MSFlexGrid1.Text = items(intloopindex)
Next intloopindex
MSFlexGrid1.Row = 0
For intloopindex = 1 To MSFlexGrid1.Cols - 1
    MSFlexGrid1.Col = intloopindex
    MSFlexGrid1.Text = Chr(Asc("A") - 1 + intloopindex)
Next intloopindex
MSFlexGrid1.Row = 1
MSFlexGrid1.Col = 1
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
    Dim introwindex As Integer
    Dim sum As Integer
    MSFlexGrid1.Text = MSFlexGrid1.Text + Chr(KeyAscii)
    oldrow = MSFlexGrid1.Row
    oldcol = MSFlexGrid1.Col
    MSFlexGrid1.Col = 2
    sum = 0
    For introwindex = 1 To MSFlexGrid1.Rows - 2
        MSFlexGrid1.Row = introwindex
        sum = sum + Val(MSFlexGrid1.Text)
    Next introwindex
    MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
    MSFlexGrid1.Text = Str(sum)
    MSFlexGrid1.Row = oldrow
    MSFlexGrid1.Col = oldcol
End Sub
