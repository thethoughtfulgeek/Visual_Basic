VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6720
   ClientLeft      =   5340
   ClientTop       =   2040
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   10185
   Begin VB.TextBox Text4 
      Height          =   735
      Left            =   6960
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   3960
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   7200
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   8160
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1920
      Width           =   855
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   2280
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   3
      DTREnable       =   0   'False
      InBufferSize    =   10
      InputLen        =   10
      OutBufferSize   =   0
      ParityReplace   =   48
      RThreshold      =   10
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1200
      Top             =   4560
   End
   Begin VB.CommandButton Command1 
      Caption         =   "asd"
      Height          =   615
      Left            =   2160
      TabIndex        =   1
      Top             =   3000
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1800
      Width           =   5295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim serdata As String
Dim lenserdata As Integer
Dim b As Integer
Dim y As Integer

Private Sub Form_Load()

            MSComm1.CommPort = 3
            BaudR = 9600
            MSComm1.Settings = "9600,n,8,1"
            MSComm1.InBufferSize = 180
            
            MSComm1.InputLen = 0
            MSComm1.RThreshold = 14 'W
            MSComm1.SThreshold = 10
            MSComm1.ParityReplace = ""
            MSComm1.Handshaking = comNone
            MSComm1.EOFEnable = False
            MSComm1.InputMode = comInputModeText
            
            If MSComm1.PortOpen = False Then
                MSComm1.PortOpen = True
            End If




End Sub

Private Sub MSComm1_OnComm()
    Select Case MSComm1.CommEvent
        ' Event messages.
        Case comEvReceive
        serdata = MSComm1.Input
        Text1.Text = serdata
        lenserdata = Len(serdata)
        Text2.Text = lenserdata
       Text4.Text = ""
        For b = 1 To lenserdata
          Text3.Text = (Mid(serdata, b, 1))
          If Text3.Text = "x" Then
          
              For y = 1 To 9
                 
                  Text4.Text = Text4.Text + (Mid(serdata, b + y, 1))
             Next y
        
          End If
     Next b
        
'NL = Chr(10) + Chr(10)
'Text1.Text = NL

'ComVartext3b1 = MSComm1.Input




        
    End Select
    

End Sub


'datalenb1 = Len(ComVartext3b1)   ' GET THE LENGTH OF THE DATA
'
'For SPxb1 = 1 To datalenb1
'    ComSerValb1(SPxb1) = Int(Asc(Mid(ComVartext3b1, SPxb1, 1))) ' byte 1
'        If SPxb1 < datalenb1 Then 'check to see that there is sufficient data available
'            ComSerValb1(SPxb1 + 1) = Int(Asc(Mid(ComVartext3b1, SPxb1 + 1, 1))) 'byte 2
'        End If
'    If ComSerValb1(SPxb1) = 250 And ComSerValb1(SPxb1 + 1) = 251 Then
'        Datapoint1b1 = SPxb1
'    End If
'Next SPxb1








