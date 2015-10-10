VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Motor Control"
   ClientHeight    =   4350
   ClientLeft      =   5325
   ClientTop       =   2025
   ClientWidth     =   7575
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   7575
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   3840
      Top             =   3360
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdexit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   615
      Left            =   5160
      TabIndex        =   7
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton cmdstart 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Start"
      Default         =   -1  'True
      Height          =   615
      Left            =   720
      TabIndex        =   6
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3120
      Top             =   3360
   End
   Begin VB.TextBox txtdir 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox txtspeed 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   3240
      Top             =   1200
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
   Begin VB.TextBox txtdist 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   600
      Width           =   3015
   End
   Begin VB.Shape motoron 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape motoroff 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Direction"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   360
      TabIndex        =   5
      Top             =   2040
      Width           =   2760
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Speed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   840
      TabIndex        =   4
      Top             =   1320
      Width           =   1725
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Distance Travelled"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   3465
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim serdata As String, test As String, test1 As String, check As String
Dim lenserdata As Integer
Dim b As Integer
Dim Y As Integer

Private Sub cmdexit_Click()
End
End Sub

Private Sub cmdstart_Click()
txtspeed.Text = ""
txtdir.Text = ""
txtdist.Text = ""
Text1.Text = ""
End Sub
Private Sub Form_Load()
            Timer1.Enabled = True
            Timer2.Enabled = False
            MSComm1.CommPort = 14
' Commport is used for selecting the Port. Enter a particular port number
            BaudR = 9600
            MSComm1.Settings = "9600,n,8,1"
' Settings are used for setting Baudrate,Parity Bit, Total Data-bits and Stop bit <-In that order
            MSComm1.InBufferSize = 180
' InBufferSize=180 Bytes. The received data shall be stored in Inbuffer in bytes
            MSComm1.InputLen = 0
' When Input property is used InputLen=10 will ensure that only 10 bytes from InbufferSize are
' taken inside. If InputLen=0 then all bytes present in InBufferSize shall be taken inside
            MSComm1.RThreshold = 15
' Used to produce an interrupt (Commevent)
' When total data actually received in InBuffersize>=Rthreshold
' Commevent will be fired
          '  MSComm1.SThreshold = 10
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
        lenserdata = Len(serdata)
        Text1.Text = lenserdata
        txtspeed.Text = ""
        For b = 1 To lenserdata
            test = (Mid(serdata, b, 1))
            If test = "1" Then
                test = (Mid(serdata, b, 4))
                If test = "1000" Then
                    'MSComm1.InputLen = 25
                    For Y = 1 To 8
                        test1 = (Mid(serdata, b + 3 + Y, 1))
                        txtdist.Text = txtdist.Text + test1
                    Next Y
                End If
            End If
            If test = "2" Then
                test = (Mid(serdata, b, 4))
                If test = "2000" Then
                    'MSComm1.InputLen = 10
                    For Y = 1 To 4
                        test1 = (Mid(serdata, b + 4 + Y, 1))
                        txtspeed.Text = txtspeed.Text + test1
                    Next Y
                End If
            End If
            If test = "3" Then
                test = (Mid(serdata, b, 6))
                If test = "333000" Then
                    'MSComm1.InputLen = 8
                    test1 = (Mid(serdata, b + 6, 1))
                  '  txtdir.Text = test1
                    If test1 = "0" Then
                        txtdir.Text = "Forward"
                    End If
                    If test1 = "1" Then
                        txtdir.Text = "Reverse"
                    End If
                End If
            End If
            If test = "4" Then
                test = (Mid(serdata, b, 6))
                If test = "444000" Then
                    test1 = (Mid(serdata, b + 6, 1))
                        If test1 = "0" Then
                       ' motoroff.Visible = True
                        'motoron.Visible = False
                            motoron.Visible = False
                            Timer1.Enabled = True
                            Timer2.Enabled = False
                        End If
                        If test1 = "1" Then
                        'motoron.Visible = True
                        'motoroff.Visible = False
                            motoroff.Visible = False
                            Timer1.Enabled = False
                            Timer2.Enabled = True
                        End If
                End If
            End If
        Next b
    End Select
    

End Sub

Private Sub Timer1_Timer()
motoroff.Visible = Not (motoroff.Visible)
End Sub

Private Sub Timer2_Timer()
motoron.Visible = Not (motoron.Visible)
End Sub
