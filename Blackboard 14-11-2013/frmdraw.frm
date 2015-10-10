VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmdraw 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Blackboard"
   ClientHeight    =   7680
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picdraw 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      Height          =   7335
      Left            =   120
      ScaleHeight     =   7275
      ScaleWidth      =   9555
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      Begin MSComDlg.CommonDialog cdl 
         Left            =   8880
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
   End
   Begin VB.Label lblcolor 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   7
      Left            =   9960
      TabIndex        =   8
      Top             =   5040
      Width           =   615
   End
   Begin VB.Label lblcolor 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   6
      Left            =   9960
      TabIndex        =   7
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lblcolor 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   5
      Left            =   9960
      TabIndex        =   6
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lblcolor 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   4
      Left            =   9960
      TabIndex        =   5
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label lblcolor 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   3
      Left            =   9960
      TabIndex        =   4
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblcolor 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   2
      Left            =   9960
      TabIndex        =   3
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblcolor 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   1
      Left            =   9960
      TabIndex        =   2
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblcolor 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   0
      Left            =   9960
      TabIndex        =   1
      Top             =   0
      Width           =   615
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnufilenew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnufileopen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnufilesave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnufilesep 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileexit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmdraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim drawon As Boolean

Private Sub Form_Load()
'load drawing colors into control array
lblcolor(0).BackColor = vbBlack
lblcolor(1).BackColor = vbRed
lblcolor(2).BackColor = vbGreen
lblcolor(3).BackColor = vbYellow
lblcolor(4).BackColor = vbBlue
lblcolor(5).BackColor = vbMagenta
lblcolor(6).BackColor = vbCyan
lblcolor(7).BackColor = vbWhite
'loading backcolor and fore colors on blackboard
picdraw.BackColor = vbbblack
picdraw.ForeColor = vbWhite
End Sub

Private Sub lblcolor_Click(Index As Integer)
'when a label color is clicked change the drawing color
'first make an audible tone and then change color
Beep
picdraw.ForeColor = lblcolor(Index).BackColor
End Sub

Private Sub mnufileexit_Click()
Dim response As Integer
response = MsgBox("Do you really want to exit?", vbYesNo + vbQuestion, "Exit application?")
If response = vbYes Then
Call mnufilesave_Click
End
End If
End Sub

Private Sub mnufilenew_Click()
Dim response As Integer
response = MsgBox("Do you really want to start over?", vbYesNo + vbQuestion, "New file?")
If response = vbYes Then
Call mnufilesave_Click
picdraw.Cls
picdraw.Picture = LoadPicture("")
End If
End Sub

Private Sub mnufileopen_Click()
'NOTE: autoredraw property must be set to true for the file to recover and open an old saved file
Dim response As Integer
'in case if cancel is pressed then error will be produced since cancelerror is true
On Error GoTo errhandler
response = MsgBox("Do you wish to open a new file?", vbYesNo + vbQuestion, "Open file?")
If response = vbYes Then
Call mnufilesave_Click
cdl.DialogTitle = "Open File"
cdl.Filter = ""
cdl.ShowOpen
picdraw.Picture = LoadPicture(cdl.FileName)
cdl.FileName = ""
End If
Exit Sub
errhandler:
'user pressed cancel button
Exit Sub
End Sub

Private Sub mnufilesave_Click()
'NOTE: autoredraw property must be set to true for the file to recover and open an old saved file
Dim response As Integer
On Error GoTo errhandler
response = MsgBox("Do you wish to save current file?", vbYesNo + vbQuestion, "Save file?")
If response = vbYes Then
    cdl.DialogTitle = "Save file"
    cdl.ShowSave
    cdl.Filter = "BMP (*.bmp)|*.bmp"
    If cdl.FileName <> "" Then
        SavePicture picdraw.Image, cdl.FileName
'We save the Image property,not the Picture property, since this is where Visual Basic maintains the persistent graphics.
        End If
Else
    Exit Sub
    End If
Exit Sub
errhandler:
'user pressed cancel button
Exit Sub
End Sub

Private Sub picdraw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'when left mouse button is pressed drawing begins
If Button = vbLeftButton Then
drawon = True
picdraw.CurrentX = X
picdraw.CurrentY = Y
End If
End Sub

Private Sub picdraw_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'if mouse is being moved while drawon is true then draw lines in current color
If drawon = True Then
picdraw.Line -(X, Y), picdraw.ForeColor
End If
End Sub

Private Sub picdraw_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'when left mouse button is released drawing ends by toggling drawon
If Button = vbLeftButton Then
drawon = False
End If
End Sub
