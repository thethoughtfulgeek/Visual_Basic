VERSION 5.00
Begin VB.MDIForm frmparent 
   BackColor       =   &H8000000C&
   Caption         =   "MDI Example"
   ClientHeight    =   7050
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10245
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnufilenew 
         Caption         =   "&New"
      End
   End
   Begin VB.Menu mnuarrange 
      Caption         =   "&Arrange"
      Begin VB.Menu mnuarrangeitem 
         Caption         =   "&Cascade"
         Index           =   0
      End
      Begin VB.Menu mnuarrangeitem 
         Caption         =   "&Horizontal Tile"
         Index           =   1
      End
      Begin VB.Menu mnuarrangeitem 
         Caption         =   "&Vertical Tile"
         Index           =   2
      End
      Begin VB.Menu mnuarrangeitem 
         Caption         =   "&Arrange Icons"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmparent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuarrangeitem_Click(Index As Integer)
Arrange Index
End Sub

Private Sub mnufilenew_Click()
Dim newdoc As New frmchild
newdoc.Show
End Sub
