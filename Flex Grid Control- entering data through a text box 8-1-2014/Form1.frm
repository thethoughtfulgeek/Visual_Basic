VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13620
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   13620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   8640
      TabIndex        =   2
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   3600
      TabIndex        =   1
      Top             =   3600
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2535
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   4471
      _Version        =   393216
      Rows            =   10
      Cols            =   10
      GridLines       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ISOCP3"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim coldrop, rowdrop As Integer
Private Sub Command1_Click()
    MSFlexGrid1.Col = 1
    MSFlexGrid1.Sort = 1
End Sub

Private Sub Form_Load()
Dim i As Integer
MSFlexGrid1.Row = 0
For i = 1 To MSFlexGrid1.Cols - 1
    MSFlexGrid1.Col = i
    MSFlexGrid1.Text = Chr(Asc("A") - 1 + i)
Next i
MSFlexGrid1.Col = 0
For i = 1 To MSFlexGrid1.Rows - 1
    MSFlexGrid1.Row = i
    MSFlexGrid1.Text = Str(i)
Next i
End Sub

Private Sub MSFlexGrid1_DragDrop(Source As Control, x As Single, y As Single)
MSFlexGrid1.ColPosition(coldrop) = MSFlexGrid1.MouseCol
MSFlexGrid1.RowPosition(rowdrop) = MSFlexGrid1.MouseRow
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
Text1.Text = Text1.Text + Chr(KeyAscii)
Text1.SelStart = 1
Text1.Move MSFlexGrid1.CellLeft + MSFlexGrid1.Left, MSFlexGrid1.CellTop + MSFlexGrid1.Top, MSFlexGrid1.CellWidth, MSFlexGrid1.CellHeight
Text1.Visible = True
Text1.SetFocus
End Sub

Private Sub MSFlexGrid1_LeaveCell()
If Text1.Visible = False Then
    Exit Sub
End If
MSFlexGrid1.Text = Text1.Text
Text1.Visible = False
Text1.Text = ""
End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
coldrop = MSFlexGrid1.MouseCol
rowdrop = MSFlexGrid1.MouseRow
MSFlexGrid1.Drag 1
End Sub
