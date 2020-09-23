VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmPortScanner 
   Caption         =   "Port Scanner"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6735
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPortScanner.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4305
   ScaleWidth      =   6735
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   120
      TabIndex        =   20
      Top             =   0
      Width           =   6495
      Begin VB.TextBox txtUpperBound 
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
         Left            =   5760
         MaxLength       =   5
         TabIndex        =   3
         Text            =   "65535"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtLowerBound 
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
         Left            =   4680
         MaxLength       =   5
         TabIndex        =   2
         Text            =   "1"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtIP 
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
         Left            =   1200
         TabIndex        =   1
         Text            =   "127.0.0.1"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Text            =   "IP Address:"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   195
         Left            =   3480
         TabIndex        =   21
         Text            =   "Ports to scan:"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblTo 
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Left            =   5400
         TabIndex        =   23
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.ListBox lstOpenPorts 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   4695
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   120
      TabIndex        =   14
      Top             =   2880
      Width           =   4695
      Begin VB.TextBox Portn 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000012&
         Height          =   210
         Left            =   1440
         TabIndex        =   16
         Text            =   "0"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   210
         Left            =   120
         TabIndex        =   15
         Text            =   "Ports Scanned:"
         Top             =   240
         Width           =   1215
      End
      Begin MSComctlLib.ProgressBar PB1 
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label Label4 
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
         Height          =   255
         Left            =   2520
         TabIndex        =   19
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "0%"
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
         Left            =   4080
         TabIndex        =   18
         Top             =   600
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   4920
      TabIndex        =   12
      Top             =   600
      Width           =   1695
      Begin VB.CommandButton Command1 
         Caption         =   "Save To File"
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
         Left            =   240
         TabIndex        =   8
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtMaxConnections 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   600
         MaxLength       =   3
         TabIndex        =   4
         Text            =   "99"
         Top             =   480
         Width           =   495
      End
      Begin VB.CommandButton cmdScan 
         Caption         =   "Scan"
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
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdStop 
         Cancel          =   -1  'True
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
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdClearList 
         Caption         =   "Clear List"
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
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Text            =   "Max Connections:"
         Top             =   240
         Width           =   1455
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
         Left            =   240
         TabIndex        =   9
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   600
         Picture         =   "FrmPortScanner.frx":1982
         Top             =   2640
         Width           =   480
      End
   End
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6120
      Top             =   480
   End
   Begin MSWinsockLib.Winsock wskSocket 
      Index           =   0
      Left            =   5880
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label5 
      Caption         =   "Open Ports:"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   600
      Width           =   4695
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3960
      Width           =   3135
   End
   Begin VB.Label Label3 
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
      Height          =   255
      Left            =   3480
      TabIndex        =   10
      Top             =   3960
      Width           =   3135
   End
End
Attribute VB_Name = "FrmPortScanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngNextPort As Long
Dim intI As Integer
Dim intI2 As Integer

Public Sub cmdClearList_Click()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_cmdClearList_Click

   Me.lstOpenPorts.Clear
   Label4.Caption = ""
   Label3.Caption = ""
   Label2.Caption = ""
   Command1.Enabled = False

EXIT_cmdClearList_Click:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_cmdClearList_Click:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in cmdClearList_Click" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_cmdClearList_Click
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_cmdClearList_Click

End Sub

Public Sub cmdScan_Click()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_cmdScan_Click

On Error Resume Next
If txtip.text = "" Then
MsgBox "Please Enter A IP", vbCritical
Exit Sub
End If
If txtLowerBound.text = "" Then
MsgBox "Please Enter A Begining Port", vbCritical
Exit Sub
End If
If txtUpperBound.text = "" Then
MsgBox "Please Enter A Ending Port", vbCritical
Exit Sub
End If
If txtMaxConnections.text = "" Then
MsgBox "Please Enter A Max Connection", vbCritical
Exit Sub
End If
   Dim intI2 As Integer
   Command1.Enabled = False
   cmdClearList.Enabled = False
   Me.txtip.Enabled = False
   Me.txtLowerBound.Enabled = False
   Me.txtUpperBound.Enabled = False
   Me.txtMaxConnections.Enabled = False
   Label4.Caption = ""
   Label2.Caption = "Scan Started at " & Date & " " & Time
   Label3.Caption = ""
   lstOpenPorts.Clear
   lngNextPort = Val(Me.txtLowerBound)
   PB1.Max = txtUpperBound.text
   PB1.Min = txtLowerBound.text
   For intI2 = 1 To Val(Me.txtMaxConnections)
   
      Load Me.wskSocket(intI2)
     
      lngNextPort = lngNextPort + 1
      
      Me.wskSocket(intI2).connect Me.txtip, lngNextPort
   
   Next intI2
timTimer.Enabled = True
 cmdStop.Enabled = True
cmdScan.Enabled = False

EXIT_cmdScan_Click:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_cmdScan_Click:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in cmdScan_Click" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_cmdScan_Click
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_cmdScan_Click

End Sub

Public Sub cmdStop_Click()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_cmdStop_Click

On Error Resume Next
   Dim intI As Integer
   timTimer.Enabled = False
   For intI = 1 To Val(Me.txtMaxConnections)
   
      Me.wskSocket(intI).Close
 
      Unload Me.wskSocket(intI)
   
   Next intI
   
cmdStop.Enabled = False
cmdScan.Enabled = True
Command1.Enabled = True
cmdClearList.Enabled = True
Label1.Caption = "0%"
Portn.text = "0"
PB1.Value = txtLowerBound.text
Label4.Caption = "Scan Stopped By User!"
Label3.Caption = "Scan Stopped at " & Date & " " & Time
FrmPortScanner.Caption = "Port Scanner"
   Me.txtip.Enabled = True
   Me.txtLowerBound.Enabled = True
   Me.txtUpperBound.Enabled = True
   Me.txtMaxConnections.Enabled = True

EXIT_cmdStop_Click:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_cmdStop_Click:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in cmdStop_Click" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_cmdStop_Click
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_cmdStop_Click

End Sub

Private Sub Command1_Click()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Command1_Click

On Error Resume Next
FrmReport.Show
FrmReport.SetFocus
DoEvents
FrmReport.List1.Clear
DoEvents
FrmReport.List1.AddItem "Address Scanned: " & FrmPortScanner.txtip.text
FrmReport.List1.AddItem ""
FrmReport.List1.AddItem FrmPortScanner.Label2.Caption
FrmReport.List1.AddItem FrmPortScanner.Label3.Caption
FrmReport.List1.AddItem ""
FrmReport.List1.AddItem "Ports Scanned: " & FrmPortScanner.txtLowerBound.text & " To " & FrmPortScanner.txtUpperBound.text
FrmReport.List1.AddItem ""
FrmReport.List1.AddItem "Total Ports Found Open: " & FrmPortScanner.lstOpenPorts.ListCount
FrmReport.List1.AddItem ""
FrmReport.List1.AddItem "Current Ports Found Open:"
FrmPortScanner.lstOpenPorts.ListIndex = 0
Do Until FrmPortScanner.lstOpenPorts.ListIndex = FrmPortScanner.lstOpenPorts.ListCount - 1
FrmReport.List1.AddItem FrmPortScanner.lstOpenPorts.text
FrmPortScanner.lstOpenPorts.ListIndex = FrmPortScanner.lstOpenPorts.ListIndex + 1
Loop
FrmReport.List1.AddItem FrmPortScanner.lstOpenPorts.text

EXIT_Command1_Click:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_Command1_Click:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in Command1_Click" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_Command1_Click
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_Command1_Click

End Sub

Private Sub Form_Load()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Form_Load

Me.Height = 4710
Me.Width = 6825
   
      Me.wskSocket(0).Close

EXIT_Form_Load:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_Form_Load:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in Form_Load" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_Form_Load
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_Form_Load

End Sub

Private Sub Form_Resize()

On Error Resume Next
Frame3.Move Me.ScaleWidth / 2 - Frame3.Width / 2, Frame3.Top
Frame1.Move Me.ScaleWidth - Frame1.Width - 95, Frame1.Top
Frame2.Move Frame2.Left, Me.ScaleHeight - 1430
Label2.Move Label2.Left, Me.ScaleHeight - 325, Me.ScaleWidth / 2
Label3.Move Label2.Left + Label2.Width, Me.ScaleHeight - 325, Me.ScaleWidth / 2
Label5.Move Label5.Left, Label5.Top, lstOpenPorts.Width
lstOpenPorts.Move lstOpenPorts.Left, lstOpenPorts.Top, Me.ScaleWidth - 2000, Me.ScaleHeight - 2200

End Sub

Private Sub Form_Unload(Cancel As Integer)

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Form_Unload

cmdStop_Click
DoEvents
Unload Me

EXIT_Form_Unload:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_Form_Unload:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in Form_Unload" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_Form_Unload
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_Form_Unload

End Sub

Private Sub timTimer_Timer()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_timTimer_Timer

On Error Resume Next
   Label1.Caption = Int((lngNextPort - Me.txtLowerBound.text) / (Me.txtUpperBound.text - Me.txtLowerBound.text) * 100) & " %" '
   FrmPortScanner.Caption = Label1.Caption & " Port Scanner"

EXIT_timTimer_Timer:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_timTimer_Timer:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in timTimer_Timer" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_timTimer_Timer
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_timTimer_Timer

End Sub

Private Sub txtIP_KeyDown(KeyCode As Integer, Shift As Integer)

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_txtIP_KeyDown

If KeyCode = vbKeyReturn Then
 Call cmdScan_Click
 DoEvents
 End If

EXIT_txtIP_KeyDown:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_txtIP_KeyDown:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in txtIP_KeyDown" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_txtIP_KeyDown
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_txtIP_KeyDown

End Sub

Private Sub txtIP_LostFocus()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_txtIP_LostFocus

On Error Resume Next
txtip.text = Replace(txtip.text, " ", "", 1, , vbTextCompare)

EXIT_txtIP_LostFocus:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_txtIP_LostFocus:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in txtIP_LostFocus" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_txtIP_LostFocus
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_txtIP_LostFocus

End Sub

Private Sub txtLowerBound_KeyDown(KeyCode As Integer, Shift As Integer)

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_txtLowerBound_KeyDown

If KeyCode = vbKeyReturn Then
 Call cmdScan_Click
 DoEvents
 End If

EXIT_txtLowerBound_KeyDown:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_txtLowerBound_KeyDown:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in txtLowerBound_KeyDown" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_txtLowerBound_KeyDown
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_txtLowerBound_KeyDown

End Sub

Private Sub txtLowerBound_LostFocus()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_txtLowerBound_LostFocus

On Error Resume Next
txtLowerBound.text = Replace(txtLowerBound.text, " ", "", 1, , vbTextCompare)

EXIT_txtLowerBound_LostFocus:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_txtLowerBound_LostFocus:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in txtLowerBound_LostFocus" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_txtLowerBound_LostFocus
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_txtLowerBound_LostFocus

End Sub

Private Sub txtMaxConnections_KeyDown(KeyCode As Integer, Shift As Integer)

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_txtMaxConnections_KeyDown

If KeyCode = vbKeyReturn Then
 Call cmdScan_Click
 DoEvents
 End If

EXIT_txtMaxConnections_KeyDown:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_txtMaxConnections_KeyDown:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in txtMaxConnections_KeyDown" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_txtMaxConnections_KeyDown
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_txtMaxConnections_KeyDown

End Sub

Private Sub txtMaxConnections_LostFocus()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_txtMaxConnections_LostFocus

On Error Resume Next
txtMaxConnections.text = Replace(txtMaxConnections.text, " ", "", 1, , vbTextCompare)

EXIT_txtMaxConnections_LostFocus:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_txtMaxConnections_LostFocus:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in txtMaxConnections_LostFocus" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_txtMaxConnections_LostFocus
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_txtMaxConnections_LostFocus

End Sub

Private Sub txtUpperBound_KeyDown(KeyCode As Integer, Shift As Integer)

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_txtUpperBound_KeyDown

If KeyCode = vbKeyReturn Then
 Call cmdScan_Click
 DoEvents
 End If

EXIT_txtUpperBound_KeyDown:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_txtUpperBound_KeyDown:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in txtUpperBound_KeyDown" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_txtUpperBound_KeyDown
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_txtUpperBound_KeyDown

End Sub

Private Sub txtUpperBound_LostFocus()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_txtUpperBound_LostFocus

On Error Resume Next
txtUpperBound.text = Replace(txtUpperBound.text, " ", "", 1, , vbTextCompare)

EXIT_txtUpperBound_LostFocus:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_txtUpperBound_LostFocus:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in txtUpperBound_LostFocus" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_txtUpperBound_LostFocus
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_txtUpperBound_LostFocus

End Sub

Private Sub wskSocket_Connect(Index As Integer)

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_wskSocket_Connect

   Me.lstOpenPorts.AddItem "Port: " & Format(Me.wskSocket(Index).RemotePort, "00000")
   'Me.Portn.Text = Format(Me.wskSocket(Index).RemotePort, "00000")
   'PB1.Value = Me.wskSocket(intI2).RemotePort
   DoEvents
   Try_Next_Port (Index)

EXIT_wskSocket_Connect:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_wskSocket_Connect:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in wskSocket_Connect" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_wskSocket_Connect
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_wskSocket_Connect

End Sub

Private Sub wskSocket_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_wskSocket_Error

   DoEvents
   Try_Next_Port (Index)

EXIT_wskSocket_Error:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_wskSocket_Error:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in wskSocket_Error" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_wskSocket_Error
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_wskSocket_Error

End Sub

Private Sub Try_Next_Port(Index As Integer)

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Try_Next_Port

On Error Resume Next
   Me.wskSocket(Index).Close

   If lngNextPort <= Val(Me.txtUpperBound) Then
      
      Me.wskSocket(Index).connect , lngNextPort
      
      lngNextPort = lngNextPort + 1
Me.Portn.text = lngNextPort
PB1.Value = lngNextPort
DoEvents
   Else

      Unload Me.wskSocket(Index)
      Me.cmdScan.Enabled = True
      Me.cmdStop.Enabled = False
      Command1.Enabled = True
      Me.timTimer.Enabled = False
      cmdClearList.Enabled = True
      Label4.Caption = "Scan Finished!"
      Label1.Caption = "0%"
      Portn.text = "0"
      PB1.Value = txtLowerBound.text
      Label3.Caption = "Scan Finished at " & Date & " " & Time
      FrmPortScanner.Caption = "Port Scanner"
   Me.txtip.Enabled = True
   Me.txtLowerBound.Enabled = True
   Me.txtUpperBound.Enabled = True
   Me.txtMaxConnections.Enabled = True

   End If

EXIT_Try_Next_Port:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_Try_Next_Port:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in Try_Next_Port" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_Try_Next_Port
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_Try_Next_Port

End Sub

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
