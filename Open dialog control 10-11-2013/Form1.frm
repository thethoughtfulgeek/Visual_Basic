VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmcommon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Common Dialog Examples"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   7755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmddisplay 
      Caption         =   "&Display Box"
      Default         =   -1  'True
      Height          =   735
      Left            =   3000
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog cdlexample 
      Left            =   6240
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "bmp"
      DialogTitle     =   "Save as Example"
   End
   Begin VB.Label lblexample 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7335
   End
End
Attribute VB_Name = "frmcommon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmddisplay_Click()
cdlexample.ShowSave
lblexample.Caption = cdlexample.FileName
End Sub
