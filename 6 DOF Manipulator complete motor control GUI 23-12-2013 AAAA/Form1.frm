VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Motor Control "
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   12315
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txttest 
      Height          =   1215
      Left            =   6960
      MultiLine       =   -1  'True
      TabIndex        =   23
      Top             =   120
      Width           =   4815
   End
   Begin VB.Timer tim_inputdata 
      Interval        =   4100
      Left            =   9120
      Top             =   5160
   End
   Begin VB.Timer tim_mot_on 
      Interval        =   500
      Left            =   8400
      Top             =   5160
   End
   Begin VB.Timer tim_mot_off 
      Interval        =   500
      Left            =   7560
      Top             =   5160
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   6600
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.TextBox txtinputdata 
      Height          =   495
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   4320
      Width           =   4695
   End
   Begin VB.CommandButton cmdrefresh 
      Caption         =   "&Refresh"
      Height          =   615
      Left            =   480
      TabIndex        =   21
      Top             =   5280
      Width           =   1575
   End
   Begin VB.TextBox txtoutputlength 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox txtoutputdist 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2880
      Width           =   2415
   End
   Begin VB.TextBox txtoutputdir 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox txtoutputvel 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdexit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   615
      Left            =   4080
      TabIndex        =   15
      Top             =   5280
      Width           =   1695
   End
   Begin VB.TextBox txtinputdcc 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox txtinputacc 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtinputdist 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2880
      Width           =   2775
   End
   Begin VB.TextBox txtinputdir 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox txtinputvel 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   7
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Length of Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   6480
      TabIndex        =   20
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Shape motoron 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   11280
      Shape           =   3  'Circle
      Top             =   5160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape motoroff 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   10200
      Shape           =   3  'Circle
      Top             =   5160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Distance Travelled"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   6360
      TabIndex        =   14
      Top             =   2880
      Width           =   2610
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
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
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   6720
      TabIndex        =   13
      Top             =   2160
      Width           =   1515
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      Caption         =   "Current Velocity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   6435
      TabIndex        =   12
      Top             =   1440
      Width           =   2265
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Deceleration Rate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   2535
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Acceleration Rate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Width           =   2505
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Distance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   525
      TabIndex        =   4
      Top             =   2880
      Width           =   1245
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
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
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   600
      TabIndex        =   3
      Top             =   2160
      Width           =   1275
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Final Velocity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   6120
      X2              =   6120
      Y1              =   0
      Y2              =   6120
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Output"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   8520
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Input"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   2400
      TabIndex        =   0
      Top             =   360
      Width           =   1185
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim serdata As String, test As String, test1 As String
Dim i As Integer, lenserdata As Integer, b As Integer, Y As Integer

Private Sub cmdexit_Click()
    End
End Sub

Private Sub cmdrefresh_Click()
    i = 1
    txtinputvel.Text = ""
    txtinputdir.Text = ""
    txtinputdist.Text = ""
    txtinputacc.Text = ""
    txtinputdcc.Text = ""
    txtoutputvel.Text = ""
    txtoutputdir.Text = ""
    txtoutputdist.Text = ""
    txtoutputlength.Text = ""
    txtinputdata.Text = ""
    txttest.Text = ""
    motoroff.Visible = False
    motoron.Visible = False
    tim_mot_on.Enabled = False
    tim_mot_off.Enabled = True
    txtinputvel.Locked = False
    MSComm1.OutBufferCount = 0
    MSComm1.InBufferCount = 0
End Sub

Private Sub Form_Load()
    i = 1
    tim_mot_off.Enabled = True
    tim_mot_on.Enabled = False
    tim_inputdata.Enabled = False
    MSComm1.CommPort = 2
' VISUAL BASIC can only detect COMM PORT between 1 to 16
    BaudR = 9600
    MSComm1.EOFEnable = False
    MSComm1.Handshaking = comNone
    MSComm1.InBufferSize = 180
    MSComm1.InputLen = 0
' InputLen=0 means that all characters in the Inbuffer shall be read into the required variable when specified
' Inputlen=x means that x number of characters in the Inbuffer shall be read into the required variable when specified
    MSComm1.InputMode = comInputModeText
    MSComm1.OutBufferSize = 180
    MSComm1.ParityReplace = ""
    MSComm1.RThreshold = 24
'Commevent->comEvReceive event shall be fired each time there are RThreshold number of characters in Inbuffer
    MSComm1.Settings = "9600,N,8,1"
    MSComm1.SThreshold = 0
    If MSComm1.PortOpen = False Then
        MSComm1.PortOpen = True
    End If
End Sub

Private Sub MSComm1_OnComm()
    Select Case MSComm1.CommEvent
        Case comEvReceive
            serdata = MSComm1.Input
' sending data from receive buffer to serdata string variable
            txttest.Text = serdata
            lenserdata = Len(serdata)
' length of received data
            txtoutputlength.Text = lenserdata
            txtoutputvel.Text = ""
            txtoutputdir.Text = ""
            For b = 1 To lenserdata
                test = (Mid(serdata, b, 1))
                If test = "9" Then
' Printing TOTAL DISTANCE travelled
                    test = (Mid(serdata, b, 4))
                    If test = "999 " Then
                        For Y = 1 To 8
                            test1 = (Mid(serdata, b + 3 + Y, 1))
' Mid command picks up the specified characters from serdata variable starting from b+3+Y number and upto 1 length.
' ie. picks out a single character
                            txtoutputdist.Text = txtoutputdist.Text + test1
                        Next Y
                    End If
                End If
                If test = "2" Then
' Printing Final Velocity in PWM
                    test = (Mid(serdata, b, 3))
                    If test = "200" Then
                        For Y = 1 To 4
                            test1 = (Mid(serdata, b + 3 + Y, 1))
                            txtoutputvel.Text = txtoutputvel.Text + test1
                        Next Y
                    End If
                End If
                If test = "3" Then
' Printing Direction
                    test = (Mid(serdata, b, 6))
                    If test = "333333" Then
                        test1 = (Mid(serdata, b + 6, 1))
                        If test1 = "0" Then
                            txtoutputdir.Text = "Forward"
                        End If
                        If test1 = "1" Then
                            txtoutputdir.Text = "Reverse"
                        End If
                    End If
                End If
                If test = "4" Then
' used to indicate whether motor is running or stopped
                    test = (Mid(serdata, b, 6))
                    If test = "444444" Then
                        test1 = (Mid(serdata, b + 6, 1))
                        If test1 = "0" Then
                            motoron.Visible = False
                            tim_mot_on.Enabled = False
                            tim_mot_off.Enabled = True
                        End If
                        If test1 = "1" Then
                            motoroff.Visible = False
                            tim_mot_on.Enabled = True
                            tim_mot_off.Enabled = False
                        End If
                    End If
                End If
            Next b
        End Select
End Sub

Private Sub tim_mot_off_Timer()
    motoroff.Visible = Not (motoroff.Visible)
End Sub

Private Sub tim_mot_on_Timer()
    motoron.Visible = Not (motoron.Visible)
End Sub

Private Sub tim_inputdata_Timer()
' this delay is used to transmit characters without any error
' By locking all textboxes during the delay this timer ensures that no data is sent while other data is being transmitted
    Select Case i
        Case 1
            tim_inputdata.Enabled = False
            MSComm1.InBufferCount = 0
' Keil has auto-echo on. Meaning whatever is sent to scanf command in Keil is received back to print on the screen
' This data is received in Receiving buffer. Hence after transmitting each data we will clear the buffer to prevent this
' data from being printed
' There is no need to clear Transmitting buffer because it will be cleared automatically each time the data is sent completely
            txtinputdata.Text = txtinputvel.Text + " "
            txtinputdir.SetFocus
' Transfers cursor to txtinputdir
            txtinputdir.Locked = False
        Case 2
            tim_inputdata.Enabled = False
            MSComm1.InBufferCount = 0
            txtinputdata.Text = txtinputdata.Text + " " + txtinputdir.Text
            txtinputdist.SetFocus
            txtinputdist.Locked = False
        Case 3
            tim_inputdata.Enabled = False
            MSComm1.InBufferCount = 0
            txtinputdata.Text = txtinputdata.Text + " " + txtinputdist.Text
            txtinputacc.SetFocus
            txtinputacc.Locked = False
        Case 4
            tim_inputdata.Enabled = False
            MSComm1.InBufferCount = 0
            txtinputdata.Text = txtinputdata.Text + " " + txtinputacc.Text
            txtinputdcc.SetFocus
            txtinputdcc.Locked = False
        Case 5
            tim_inputdata.Enabled = False
' the timer is used as a stopwatch. Started each time the data is sent and and stopped once the operation is completed.
            MSComm1.InBufferCount = 0
            txtinputdata.Text = txtinputdata.Text + " " + txtinputdcc.Text
    End Select
End Sub

Private Sub txtinputdcc_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        txtinputdcc.Locked = True
        MSComm1.Output = txtinputdcc.Text & vbCr
' Sending data back to controller. It is necessary to enter vbCr with each data that is to be sent.
        i = 5
        tim_inputdata.Enabled = True
    End If
    If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Then
        Exit Sub
    Else
         KeyAscii = 0
        Beep
    End If
End Sub

Private Sub txtinputacc_KeyPress(KeyAscii As Integer)
' Keytrapping event.Does not allow to type anything except the required keys
    If KeyAscii = vbKeyReturn Then
        txtinputacc.Locked = True
        MSComm1.Output = txtinputacc.Text & vbCr
        i = 4
        tim_inputdata.Enabled = True
    End If
    If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Then
        Exit Sub
    Else
        KeyAscii = 0
        Beep
    End If
End Sub

Private Sub txtinputdist_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtinputdist.Locked = True
        MSComm1.Output = txtinputdist.Text & vbCr
        i = 3
        tim_inputdata.Enabled = True
    End If
    If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Then
        Exit Sub
    Else
       KeyAscii = 0
       Beep
   End If
End Sub

Private Sub txtinputdir_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtinputdir.Locked = True
        MSComm1.Output = txtinputdir.Text & vbCr
        i = 2
        tim_inputdata.Enabled = True
    End If
    If KeyAscii = vbKey0 Or KeyAscii = vbKey1 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Then
        Exit Sub
    Else
        KeyAscii = 0
        Beep
    End If
End Sub

Private Sub txtinputvel_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtinputvel.Locked = True
        MSComm1.Output = txtinputvel.Text & vbCr
        i = 1
        tim_inputdata.Enabled = True
    End If
    If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Then
        Exit Sub
    Else
        KeyAscii = 0
        Beep
    End If
End Sub

