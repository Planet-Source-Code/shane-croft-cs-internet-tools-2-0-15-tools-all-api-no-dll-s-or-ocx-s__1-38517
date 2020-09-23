VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form FrmPortListen 
   Caption         =   "Port Listener"
   ClientHeight    =   3450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4530
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPortListen.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3450
   ScaleWidth      =   4530
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   600
      TabIndex        =   6
      Top             =   0
      Width           =   3495
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Listen"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdDisconnect 
         Caption         =   "Stop"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton optTCP 
         Caption         =   "TCP/IP"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optUDP 
         Caption         =   "UDP"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox port3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   600
         MaxLength       =   5
         TabIndex        =   0
         Text            =   "1979"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Close 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   960
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   2760
         Picture         =   "FrmPortListen.frx":1D12
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Protocol:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Port:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.TextBox txtStatus 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   1815
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1560
      Width           =   4335
   End
   Begin MSWinsockLib.Winsock ws1 
      Left            =   4080
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmPortListen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Close_Click()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Close_Click

Unload Me

EXIT_Close_Click:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_Close_Click:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in Close_Click" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_Close_Click
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_Close_Click

End Sub

Private Sub cmdConnect_Click()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_cmdConnect_Click

If port3.text = "" Then
MsgBox "Please Enter A Port Number", vbCritical
Exit Sub
End If
cmdConnect.Enabled = False
port3.Enabled = False
cmdDisconnect.Enabled = True
txtStatus = ""
If optTCP = True Then
    ws1.protocol = sckTCPProtocol
End If
If optUDP = True Then
    ws1.protocol = sckUDPProtocol
End If
On Error GoTo PortIsOpen
ws1.Close
ws1.LocalPort = port3.text
ws1.listen
Exit Sub
PortIsOpen:
ws1.Close
If Err.Number = 10048 Then
    txtStatus = "The port " & port3.text & " is already open."
Else
    txtStatus = "Error: " & Err.Number & vbCrLf & "   " & Err.Description
End If
cmdDisconnect.Enabled = False
port3.Enabled = True
cmdConnect.Enabled = True

EXIT_cmdConnect_Click:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_cmdConnect_Click:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in cmdConnect_Click" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_cmdConnect_Click
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_cmdConnect_Click

End Sub

Private Sub cmdDisconnect_Click()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_cmdDisconnect_Click

ws1.Close
cmdDisconnect.Enabled = False
port3.Enabled = True
cmdConnect.Enabled = True

EXIT_cmdDisconnect_Click:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_cmdDisconnect_Click:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in cmdDisconnect_Click" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_cmdDisconnect_Click
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_cmdDisconnect_Click

End Sub


Private Sub Form_Load()

Me.Height = 3855
Me.Width = 4650

optTCP = True

End Sub


Private Sub Form_Resize()

On Error Resume Next
Frame1.Move Me.ScaleWidth / 2 - Frame1.Width / 2, Frame1.Top
txtStatus.Move txtStatus.Left, txtStatus.Top, Me.ScaleWidth - 200, Me.ScaleHeight - 1600

End Sub

Private Sub port3_KeyDown(KeyCode As Integer, Shift As Integer)

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_port3_KeyDown

If KeyCode = vbKeyReturn Then
 Call cmdConnect_Click
 DoEvents
 End If

EXIT_port3_KeyDown:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_port3_KeyDown:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in port3_KeyDown" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_port3_KeyDown
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_port3_KeyDown

End Sub

Private Sub port3_LostFocus()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_port3_LostFocus

On Error Resume Next
port3.text = Replace(port3.text, " ", "", 1, , vbTextCompare)

EXIT_port3_LostFocus:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_port3_LostFocus:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in port3_LostFocus" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_port3_LostFocus
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_port3_LostFocus

End Sub

Private Sub ws1_ConnectionRequest(ByVal requestID As Long)
 'If ws1.State <> sckClosed Then ws1.Close
 'ws1.Accept (requestID)

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_ws1_ConnectionRequest

 txtStatus.text = txtStatus.text & vbCrLf & "Connection" & " - " & Date & " " & Time

EXIT_ws1_ConnectionRequest:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_ws1_ConnectionRequest:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in ws1_ConnectionRequest" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_ws1_ConnectionRequest
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_ws1_ConnectionRequest

End Sub

Private Sub ws1_DataArrival(ByVal bytesTotal As Long)

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_ws1_DataArrival

Dim strdata As String
ws1.GetData strdata
txtStatus.text = txtStatus.text & vbCrLf & " - " & strdata & " - " & Date & " " & Time

EXIT_ws1_DataArrival:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_ws1_DataArrival:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in ws1_DataArrival" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_ws1_DataArrival
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_ws1_DataArrival

End Sub

Private Sub ws1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_ws1_Error

txtStatus = txtStatus.text & vbCrLf & "Winsock Error: " & Number & vbCrLf & "   " & descriptoin & " - " & Date & " " & Time

EXIT_ws1_Error:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_ws1_Error:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in ws1_Error" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_ws1_Error
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_ws1_Error

End Sub
