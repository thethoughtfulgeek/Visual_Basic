VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15180
   LinkTopic       =   "Form1"
   ScaleHeight     =   9375
   ScaleWidth      =   15180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsheet 
      Caption         =   "S&how Sheet"
      Height          =   495
      Left            =   720
      TabIndex        =   9
      Top             =   7680
      Width           =   1935
   End
   Begin VB.TextBox txtcol 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12840
      TabIndex        =   8
      Top             =   7920
      Width           =   615
   End
   Begin VB.TextBox txtrow 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12840
      TabIndex        =   7
      Top             =   7320
      Width           =   615
   End
   Begin VB.CommandButton cmdexit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   6960
      Width           =   1935
   End
   Begin VB.CommandButton cmdchart 
      Caption         =   "&Show Chart"
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   6240
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   10680
      TabIndex        =   2
      Top             =   6480
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1575
      Left            =   3720
      TabIndex        =   1
      Top             =   6360
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   2778
      _Version        =   393216
      Rows            =   6
      Cols            =   6
      AllowUserResizing=   3
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   4935
      Left            =   1680
      OleObjectBlob   =   "Form1.frx":0000
      TabIndex        =   0
      Top             =   480
      Width           =   10695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No. of Columns"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10800
      TabIndex        =   6
      Top             =   7920
      Width           =   1845
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No. of Rows"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10905
      TabIndex        =   5
      Top             =   7320
      Width           =   1485
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x(1 To 5, 1 To 5) As Variant
Dim flag As Boolean
Dim i As Integer, j As Integer


Private Sub cmdexit_Click()
    End
End Sub

Private Sub cmdsheet_Click()
    MSFlexGrid1.Rows = Val(txtrow.Text)
    MSFlexGrid1.Cols = Val(txtcol.Text)
    MSFlexGrid1.Row = 0
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
    MSFlexGrid1.Row = 1
    MSFlexGrid1.Col = 1
    MSFlexGrid1.Visible = True
End Sub

Private Sub cmdchart_Click()
    MSChart1.chartType = VtChChartType2dLine
    MSChart1.ColumnCount = MSFlexGrid1.Cols - 1
' don't send fixed column data
    MSChart1.RowCount = MSFlexGrid1.Rows - 1
' don't send fixed row data
    For i = 1 To MSFlexGrid1.Rows - 1
        MSFlexGrid1.Row = i
        For j = 1 To MSFlexGrid1.Cols - 1
            MSFlexGrid1.Col = j
            x(i, j) = MSFlexGrid1.Text
            MSChart1.Column = j
            MSChart1.Row = i
            MSChart1.Data = x(i, j)
            MSChart1.ColumnLabel = Chr(Asc("A") - 1 + j)
        Next j
        MSChart1.RowLabel = Str(i)
    Next i
    MSChart1.ShowLegend = True
    MSChart1.Visible = True
End Sub


Private Sub Form_Load()
    MSChart1.Visible = False
    MSFlexGrid1.Visible = False
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
' only when some key is pressed inside a cell we shall call the textbox here
    Text1.Text = Text1.Text & Chr(KeyAscii)
    Text1.SelStart = 1
    Text1.Move MSFlexGrid1.Left + MSFlexGrid1.CellLeft, MSFlexGrid1.Top + MSFlexGrid1.CellTop, MSFlexGrid1.CellWidth, MSFlexGrid1.CellHeight
' left gives distance between leftmost end of grid and form
' top gives distance between topmost end of grid and form
' celltop gives distance between topmost end of selected cell and grid
' cellleft gives distance between leftmost end of selected cell and grid
' cellwidth gives width of selected cell
' cellheight gives height of selected cell
    Text1.Visible = True
    Text1.SetFocus
End Sub

Private Sub MSFlexGrid1_LeaveCell()
    If Text1.Visible = False Then
' no data will be entered inside a cell if a user is randomly moving from one textbox to other.
' hence in case of this kind of leavecell event no data will be moved from textbox to the cell
        Exit Sub
    End If
    MSFlexGrid1.Text = Text1.Text
    Text1.Text = ""
    Text1.Visible = False
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
' This allows using Down Direction key for browsing cells in flexgrid
            MSFlexGrid1.SetFocus
            If MSFlexGrid1.Row = MSFlexGrid1.Rows - 1 Then
                MSFlexGrid1.Row = MSFlexGrid1.FixedRows
' on reaching the last row come back to the top one
                flag = True
' the flag ensures that when we reach the first row from last row we don't automatically move to the 2nd row due to
' next if condition
            End If
            If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 And flag = False Then
' if last row has not been reached then keep on moving down whenever down key is pressed
' only if flag is not set you can move to next row. if flag is set, it means that just now first row has come from the
' last one
                MSFlexGrid1.Row = MSFlexGrid1.Row + 1
            End If
            flag = False
        Case vbKeyUp
' This allows using Up Direction key for browsing cells in flexgrid
            MSFlexGrid1.SetFocus
            If MSFlexGrid1.Row = MSFlexGrid1.FixedRows Then
                MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
                flag = True
            End If
            If MSFlexGrid1.Row > MSFlexGrid1.FixedRows And flag = False Then
                MSFlexGrid1.Row = MSFlexGrid1.Row - 1
            End If
            flag = False
        Case vbKeyLeft
' This allows using Left Direction key for browsing cells in flexgrid
            MSFlexGrid1.SetFocus
            If MSFlexGrid1.Col = MSFlexGrid1.FixedCols Then
                MSFlexGrid1.Col = MSFlexGrid1.Cols - 1
                flag = True
            End If
            If MSFlexGrid1.Col > MSFlexGrid1.FixedCols And flag = False Then
                MSFlexGrid1.Col = MSFlexGrid1.Col - 1
            End If
            flag = False
        Case vbKeyRight
' This allows using Right Direction key for browsing cells in flexgrid
            MSFlexGrid1.SetFocus
            If MSFlexGrid1.Col = MSFlexGrid1.Cols - 1 Then
                MSFlexGrid1.Col = MSFlexGrid1.FixedCols
                flag = True
            End If
            If MSFlexGrid1.Col < MSFlexGrid1.Cols - 1 And flag = False Then
                MSFlexGrid1.Col = MSFlexGrid1.Col + 1
            End If
            flag = False
    End Select
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
' This is because enter key cannot be detected by keydown event. hence if we want to move one row down on pressing
' enter key, we will use this event
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
'keyascii=0 suppreses the BEEP sound that occurs on pressing enter in a textbox
        Call Text1_KeyDown(vbKeyDown, 0)
    End If
End Sub
