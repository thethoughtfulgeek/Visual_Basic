VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   8445
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image3 
      Height          =   1440
      Left            =   480
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   1200
   End
   Begin VB.Image Image2 
      Height          =   825
      Left            =   600
      Picture         =   "Form1.frx":0442
      Stretch         =   -1  'True
      Top             =   2160
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   825
      Left            =   600
      Picture         =   "Form1.frx":0884
      Stretch         =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image3_Click()
Static picnum As Integer
If picnum = 0 Then
Image3.Picture = Image1.Picture
picnum = 1
Else
Image3.Picture = Image2.Picture
picnum = 0
End If
End Sub
