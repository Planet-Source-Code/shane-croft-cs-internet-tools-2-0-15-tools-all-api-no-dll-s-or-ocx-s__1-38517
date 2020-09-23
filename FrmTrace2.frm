VERSION 5.00
Begin VB.Form FrmTrace2 
   Caption         =   "Trace Route Built Into Windows"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7560
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmTrace2.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3720
   ScaleWidth      =   7560
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   600
      TabIndex        =   4
      Top             =   0
      Width           =   6135
      Begin VB.CheckBox Check1 
         Caption         =   "Resolve Ip To Host."
         Height          =   210
         Left            =   1080
         TabIndex        =   3
         Top             =   600
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CommandButton TraceRT2 
         Caption         =   "Trace Route"
         Default         =   -1  'True
         Height          =   255
         Left            =   3495
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Close 
         Caption         =   "Close"
         Height          =   255
         Left            =   4800
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Host 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   840
         TabIndex        =   0
         Top             =   240
         Width           =   2535
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   4800
         Picture         =   "FrmTrace2.frx":1982
         Top             =   840
         Width           =   480
      End
      Begin VB.Label lblIPHost 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "IP/Host:"
         Height          =   225
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   4215
      End
   End
   Begin VB.TextBox txtNS 
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   2175
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   1440
      Width           =   7335
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6975
      Top             =   0
   End
End
Attribute VB_Name = "FrmTrace2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
' This project shows netstat instrution of commnand and show it in Windows Forms
'
' Declare all API, first

Private Declare Function CreatePipe Lib "kernel32" ( _
    phReadPipe As Long, _
    phWritePipe As Long, _
    lpPipeAttributes As Any, _
    ByVal nSize As Long) As Long

Private Declare Function ReadFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByVal lpBuffer As String, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Any) As Long

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadID As Long
End Type

Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
   lpApplicationName As Long, ByVal lpCommandLine As String, _
   lpProcessAttributes As Any, lpThreadAttributes As Any, _
   ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
   ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
   lpStartupInfo As Any, lpProcessInformation As Any) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" _
    (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal _
   hObject As Long) As Long

Const SW_SHOWMINNOACTIVE = 7
Const STARTF_USESHOWWINDOW = &H1
Const INFINITE = -1&
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const STARTF_USESTDHANDLES = &H100&

' to execute the browser
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_SHOWDEFAULT = 10
'delay function
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Function ExecCmdPipe(ByVal CmdLine As String) As String
    'Executes the command, and when it finish returns value to VB

    Dim proc As PROCESS_INFORMATION, ret As Long, bSuccess As Long
    Dim start As STARTUPINFO
    Dim sa As SECURITY_ATTRIBUTES
    Dim hReadPipe As Long, hWritePipe As Long
    Dim bytesread As Long, mybuff As String
    Dim i As Integer
    
    Dim sReturnStr As String
    
    ' the lenght of the string must be 10 * 1024
    
    mybuff = String(10 * 1024, Chr$(65))
    sa.nLength = Len(sa)
    sa.bInheritHandle = 1&
    sa.lpSecurityDescriptor = 0&
    ret = CreatePipe(hReadPipe, hWritePipe, sa, 0)
    If ret = 0 Then
        '===Error
        ExecCmdPipe = "Error: CreatePipe failed. " & Err.LastDllError
        Exit Function
    End If
    start.cb = Len(start)
    start.hStdOutput = hWritePipe
    start.dwFlags = STARTF_USESTDHANDLES + STARTF_USESHOWWINDOW
    start.wShowWindow = SW_SHOWMINNOACTIVE
    
    ' Start the shelled application:
    ret& = CreateProcessA(0&, CmdLine$, sa, sa, 1&, _
        NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
    If ret <> 1 Then
        '===Error
        sReturnStr = "Error: CreateProcess failed. " & Err.LastDllError
    End If
    
    ' Wait for the shelled application to finish:
    ret = WaitForSingleObject(proc.hProcess, INFINITE)
    

    bSuccess = ReadFile(hReadPipe, mybuff, Len(mybuff), bytesread, 0&)
    If bSuccess = 1 Then
        sReturnStr = Left(mybuff, bytesread)
    Else
        '===Error
        sReturnStr = "Error: ReadFile failed. " & Err.LastDllError
    End If
    ret = CloseHandle(proc.hProcess)
    ret = CloseHandle(proc.hThread)
    ret = CloseHandle(hReadPipe)
    ret = CloseHandle(hWritePipe)
    
    'returns to VB
    ExecCmdPipe = sReturnStr
    Me.Label1.Caption = "Done"
DoEvents
End Function

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

Private Sub Form_Load()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Form_Load

Me.Height = 4125
Me.Width = 7680

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
Frame1.Move Me.ScaleWidth / 2 - Frame1.Width / 2, Frame1.Top
txtNS.Move txtNS.Left, txtNS.Top, Me.ScaleWidth - 200, Me.ScaleHeight - 1500

End Sub

Private Sub Timer1_Timer()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Timer1_Timer

If Host.text = "" Then
TraceRT2.Enabled = False
Else
TraceRT2.Enabled = True
End If

EXIT_Timer1_Timer:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_Timer1_Timer:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in Timer1_Timer" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_Timer1_Timer
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_Timer1_Timer

End Sub

Private Sub TraceRT2_Click()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_TraceRT2_Click

If Me.Host.text = "" Then
MsgBox "Please enter a domain name.", vbCritical
Me.Host.SetFocus
Exit Sub
End If

Me.Label1.Caption = "Please Wait..."
DoEvents
Me.txtNS.text = ""
DoEvents
If Me.Check1.Value = 1 Then
Me.txtNS = ExecCmdPipe("Tracert " & Me.Host.text)
Else
Me.txtNS = ExecCmdPipe("Tracert " & Me.Host.text & " -d")
End If

EXIT_TraceRT2_Click:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_TraceRT2_Click:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in TraceRT2_Click" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_TraceRT2_Click
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_TraceRT2_Click

End Sub
