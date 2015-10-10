VERSION 5.00
Begin VB.Form frmdispose 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Letter Disposal"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdreset 
      Caption         =   "&Reset"
      Height          =   615
      Left            =   1800
      TabIndex        =   0
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Image imgletter 
      DragIcon        =   "Form1.frx":0000
      DragMode        =   1  'Automatic
      Height          =   1440
      Left            =   240
      Picture         =   "Form1.frx":0442
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1800
   End
   Begin VB.Image imgburn 
      Height          =   480
      Left            =   4920
      Picture         =   "Form1.frx":0884
      Top             =   3960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgtrash 
      Height          =   480
      Left            =   4080
      Picture         =   "Form1.frx":0CC6
      Top             =   3960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgcan 
      Height          =   1560
      Left            =   4200
      Picture         =   "Form1.frx":1108
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1200
   End
End
Attribute VB_Name = "frmdispose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdreset_Click()
'reset to trash can picture
imgcan.Picture = imgtrash.Picture
imgletter.Visible = True
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
Source.Move X, Y
End Sub

Private Sub imgcan_DragDrop(Source As Control, X As Single, Y As Single)
'burn mail and make it disappear
imgcan.Picture = imgburn.Picture
Source.Visible = False
End Sub
