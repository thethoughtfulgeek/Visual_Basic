VERSION 5.00
Begin VB.Form frmimage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Image Viewer"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   9030
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmexit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   495
      Left            =   7200
      TabIndex        =   8
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdshow 
      Caption         =   "&Show Image"
      Default         =   -1  'True
      Height          =   495
      Left            =   5400
      TabIndex        =   7
      Top             =   4800
      Width           =   1215
   End
   Begin VB.FileListBox filimage 
      Height          =   3210
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   2055
   End
   Begin VB.DirListBox dirimage 
      Height          =   1890
      Left            =   2400
      TabIndex        =   5
      Top             =   1920
      Width           =   2415
   End
   Begin VB.DriveListBox drvimage 
      Height          =   315
      Left            =   2400
      TabIndex        =   4
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Image imgimage 
      BorderStyle     =   1  'Fixed Single
      Height          =   3615
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   600
      Width           =   2655
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   3615
      Left            =   5760
      Top             =   600
      Width           =   2655
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   5040
      X2              =   5040
      Y1              =   0
      Y2              =   5520
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Drives:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2520
      TabIndex        =   3
      Top             =   4080
      Width           =   1230
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Directories:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2520
      TabIndex        =   2
      Top             =   1200
      Width           =   2040
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Files:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   990
   End
   Begin VB.Label lblimage 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FF0000&
      FillStyle       =   4  'Upward Diagonal
      Height          =   4335
      Left            =   5400
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "frmimage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdshow_Click()
'put image file name together and load image into image box
Dim imagename As String
'check to see if at root directory
If Right(filimage.Path, 1) = "\" Then
imagename = filimage.Path + filimage.FileName
Else
imagename = filimage.Path + "\" + filimage.FileName
End If
lblimage.Caption = imagename
imgimage.Picture = LoadPicture(imagename)
End Sub

Private Sub cmexit_Click()
End
End Sub

Private Sub dirimage_Change()
'if directory changes then change file path too
filimage.Path = dirimage.Path
End Sub

Private Sub drvimage_Change()
'if drive changes then change directory
dirimage.Path = drvimage.Drive
End Sub

Private Sub filimage_DblClick()
Call cmdshow_Click
End Sub
