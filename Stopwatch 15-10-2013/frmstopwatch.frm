VERSION 5.00
Begin VB.Form frmstopwatch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "StopWatch Application"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   7740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton cmdend 
      Caption         =   "&End Timing"
      Enabled         =   0   'False
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton cmdstart 
      Caption         =   "&Start Timing"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label belapsed 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   4440
      TabIndex        =   8
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label bend 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   4440
      TabIndex        =   7
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label bstart 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   4440
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Elapsed Time"
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "End Time"
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Start Time"
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmstopwatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim starttime As Variant, endtime As Variant, elapsedtime As Variant
Option Explicit

Private Sub cmdend_Click()
'find the ending time, compute the elapsed time
'put both values in label boxes
endtime = Now
elapsedtime = endtime - starttime
bend.Caption = Format(endtime, "hh:mm:ss")
belapsed.Caption = Format(elapsedtime, "hh:mm:ss")
cmdstart.Enabled = True
End Sub

Private Sub cmdexit_Click()
End
End Sub

Private Sub cmdstart_Click()
'establish and print starting time
starttime = Now
bstart.Caption = Format(starttime, "hh:mm:ss")
cmdend.Enabled = True
cmdstart.Enabled = False
bend.Caption = ""
belapsed.Caption = ""
End Sub


