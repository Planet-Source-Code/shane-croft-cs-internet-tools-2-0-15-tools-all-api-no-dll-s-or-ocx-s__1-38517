VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form FrmTime 
   Caption         =   "Time Sync"
   ClientHeight    =   1425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3555
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmTime.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1425
   ScaleWidth      =   3555
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   240
         Width           =   2310
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Synch"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2505
         TabIndex        =   1
         Top             =   240
         Width           =   780
      End
      Begin MSWinsockLib.Winsock StinkySock 
         Left            =   2160
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   2760
         Picture         =   "FrmTime.frx":1982
         Top             =   600
         Width           =   480
      End
      Begin VB.Label Label1 
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   2415
      End
   End
End
Attribute VB_Name = "FrmTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetSystemTime Lib "kernel32" _
   (lpSystemTime As SYSTEMTIME) As Long
   
Private Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type

Dim sNTP As String ' the 32bit time stamp returned by the server
Dim TimeDelay As Single 'the time between the acknowledgement of
                        'the connection and the data received.
                        'we compensate by adding half of the round
                        'trip latency

Private Sub Command1_Click()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Command1_Click

On Error Resume Next
Label1.Caption = "Please Wait..."
DoEvents
StinkySock.Close
sNTP = Empty
StinkySock.RemoteHost = Combo1.text
StinkySock.RemotePort = 37 'NTP servers port
StinkySock.connect

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

On Error Resume Next
Me.Height = 1830
Me.Width = 3675

Combo1.AddItem "time.ien.it"
Combo1.AddItem "ntp.cs.mu.oz.au"
Combo1.AddItem "tock.usno.navy.mil"
Combo1.AddItem "tick.usno.navy.mil"
Combo1.AddItem "swisstime.ethz.ch"
Combo1.AddItem "ntp-cup.external.hp.com"
Combo1.AddItem "ntp1.fau.de"
Combo1.AddItem "ntps1-0.cs.tu-berlin.de"
Combo1.AddItem "ntps1-1.rz.Uni-Osnabrueck.DE"
Combo1.AddItem "tempo.cstv.to.cnr.it"
Combo1.ListIndex = 0

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Form_QueryUnload

On Error Resume Next
DoEvents
StinkySock.Close
sNTP = Empty

EXIT_Form_QueryUnload:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_Form_QueryUnload:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in Form_QueryUnload" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_Form_QueryUnload
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_Form_QueryUnload

End Sub

Private Sub Form_Resize()
On Error Resume Next
Frame1.Move Me.ScaleWidth / 2 - Frame1.Width / 2, Me.ScaleHeight / 2 - Frame1.Height / 2

End Sub

Private Sub StinkySock_DataArrival(ByVal bytesTotal As Long)

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_StinkySock_DataArrival

Dim DATA As String

StinkySock.GetData DATA, vbString
sNTP = sNTP & DATA

EXIT_StinkySock_DataArrival:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_StinkySock_DataArrival:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in StinkySock_DataArrival" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_StinkySock_DataArrival
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_StinkySock_DataArrival

End Sub

Private Sub StinkySock_Connect()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_StinkySock_Connect

On Error Resume Next
TimeDelay = Timer

EXIT_StinkySock_Connect:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_StinkySock_Connect:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in StinkySock_Connect" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_StinkySock_Connect
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_StinkySock_Connect

End Sub

Private Sub StinkySock_Close()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_StinkySock_Close

On Error Resume Next
Do Until StinkySock.State = sckClosed
 StinkySock.Close
 DoEvents
Loop
TimeDelay = ((Timer - TimeDelay) / 2)
Call SyncClock(sNTP)

EXIT_StinkySock_Close:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_StinkySock_Close:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in StinkySock_Close" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_StinkySock_Close
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_StinkySock_Close

End Sub

Private Sub SyncClock(tStr As String)

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_SyncClock

On Error Resume Next
Dim NTPTime As Double
Dim UTCDATE As Date
Dim LngTimeFrom1990 As Long
Dim ST As SYSTEMTIME
     
tStr = Trim(tStr)
If Len(tStr) <> 4 Then
 Label1.Caption = "NTP Server returned an invalid response."
 Exit Sub
End If

NTPTime = Asc(Left$(tStr, 1)) * 256 ^ 3 + Asc(Mid$(tStr, 2, 1)) * 256 ^ 2 + Asc(Mid$(tStr, 3, 1)) * 256 ^ 1 + Asc(Right$(tStr, 1))
      
LngTimeFrom1990 = NTPTime - 2840140800#

UTCDATE = DateAdd("s", CDbl(LngTimeFrom1990 + CLng(TimeDelay)), #1/1/1990#)

ST.wYear = Year(UTCDATE)
ST.wMonth = Month(UTCDATE)
ST.wDay = Day(UTCDATE)
ST.wHour = Hour(UTCDATE)
ST.wMinute = Minute(UTCDATE)
ST.wSecond = Second(UTCDATE)

Call SetSystemTime(ST)
Label1.Caption = "Clock synchronised succesfully."

EXIT_SyncClock:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_SyncClock:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in SyncClock" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_SyncClock
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_SyncClock

End Sub

