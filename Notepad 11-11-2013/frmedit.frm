VERSION 5.00
Begin VB.Form frmedit 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Note Editor"
   ClientHeight    =   5655
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7305
   Icon            =   "frmedit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   7305
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtedit 
      Height          =   5295
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   6615
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnufilenew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnufilebar 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuformat 
      Caption         =   "F&ormat"
      Begin VB.Menu mnuformatbold 
         Caption         =   "&Bold"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuformatitalic 
         Caption         =   "&Italic"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuformatunderline 
         Caption         =   "&Underline"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuformatsize 
         Caption         =   "&Siize"
         Begin VB.Menu mnuformatsizesmall 
            Caption         =   "&Small"
            Checked         =   -1  'True
            Shortcut        =   ^S
         End
         Begin VB.Menu mnuformatsizemedium 
            Caption         =   "&Medium"
            Shortcut        =   ^M
         End
         Begin VB.Menu mnuformatsizelarge 
            Caption         =   "&Large"
            Shortcut        =   ^L
         End
      End
   End
End
Attribute VB_Name = "frmedit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnufileexit_Click()
'check if the user really wants to exit the file
Dim response As Integer
response = MsgBox("Do you want to exit?", vbYesNo + vbQuestion, "Exit file")
If response = vbYes Then
End
End If
End Sub

Private Sub mnufilenew_Click()
'check if the user really wants to open a new file
Dim response As Integer
response = MsgBox("Are you sure you want a new file?", vbYesNo + vbQuestion, "New file")
If response = vbYes Then
txtedit.Text = ""
End If
End Sub

Private Sub mnuformatbold_Click()
'toggle bold font status
mnuformatbold.Checked = Not (mnuformatbold.Checked)
txtedit.FontBold = Not (txtedit.FontBold)
End Sub

Private Sub mnuformatitalic_Click()
'toggle italic font status
mnuformatitalic.Checked = Not (mnuformatitalic.Checked)
txtedit.FontItalic = Not (txtedit.FontItalic)
End Sub

Private Sub mnuformatsizelarge_Click()
'check large size and uncheck all others
'also activate large size fonts
mnuformatsizesmall.Checked = False
mnuformatsizemedium.Checked = False
mnuformatsizelarge.Checked = True
txtedit.FontSize = 18
End Sub

Private Sub mnuformatsizemedium_Click()
'check medium size and uncheck all others
'also activate medium size fonts
mnuformatsizesmall.Checked = False
mnuformatsizemedium.Checked = True
mnuformatsizelarge.Checked = False
txtedit.FontSize = 12
End Sub

Private Sub mnuformatsizesmall_Click()
'check small size and uncheck all others
'also activate small fonts
mnuformatsizesmall.Checked = True
mnuformatsizemedium.Checked = False
mnuformatsizelarge.Checked = False
txtedit.FontSize = 8
End Sub

Private Sub mnuformatunderline_Click()
'toggle underline font status
mnuformatunderline.Checked = Not (mnuformatunderline.Checked)
txtedit.FontUnderline = Not (txtedit.FontUnderline)
End Sub


