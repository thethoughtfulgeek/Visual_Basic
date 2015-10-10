VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5370
   FillColor       =   &H0000FFFF&
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   307
   ScaleMode       =   0  'User
   ScaleWidth      =   348
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Sub Form_Resize()
Dim rtnvalue As Long
Form1.Cls
rtnvalue = Ellipse(Form1.hdc, 0.1 * ScaleWidth, 0.1 * ScaleHeight, 0.9 * ScaleWidth, 0.9 * ScaleHeight)
End Sub
