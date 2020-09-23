VERSION 5.00
Begin VB.Form FrmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CS Internet Tools Settings"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6135
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   8
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "Misc. Settings"
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   2655
      Begin VB.CheckBox Check1 
         Caption         =   "Start Program Maximized"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Bandwidth Meter"
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   3735
      Begin VB.CheckBox Check4 
         Caption         =   "Enable Bandiwdth Meter"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.Image Image2 
         Height          =   870
         Left            =   2760
         Picture         =   "FrmSettings.frx":1782
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sound"
      Height          =   1215
      Left            =   2880
      TabIndex        =   0
      Top             =   720
      Width           =   3135
      Begin VB.CheckBox Check3 
         Caption         =   "Enable Shutdown Sound"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   2295
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Enable Startup Sound"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2175
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   2520
         Picture         =   "FrmSettings.frx":1D6D
         Top             =   360
         Width           =   480
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   1080
      TabIndex        =   9
      Top             =   120
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "FrmSettings.frx":36EF
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "FrmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OSGetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function OSGetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function OSGetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function OSWritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function OSWritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Declare Function OSGetProfileInt Lib "kernel32" Alias "GetProfileIntA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Long) As Long
Private Declare Function OSGetProfileSection Lib "kernel32" Alias "GetProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function OSGetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long

Private Declare Function OSWriteProfileSection Lib "kernel32" Alias "WriteProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String) As Long
Private Declare Function OSWriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long

Private Const nBUFSIZEINI = 1024
Private Const nBUFSIZEINIALL = 4096
Private FilePathName As String

Private Sub Command1_Click()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Command1_Click

Dim fFile As Integer
fFile = FreeFile
 
Open App.Path & "\Settings.inf" For Output As fFile
Print #fFile, "[settings]"
Print #fFile, "maximized=" & Me.Check1.Value
Print #fFile, "startupsound=" & Me.Check2.Value
Print #fFile, "shutdownsound=" & Me.Check3.Value
Print #fFile, "bandwidth=" & Me.Check4.Value
Close fFile
DoEvents
DoEvents

If Check4.Value = 1 Then
FrmMain.Timer1.Enabled = True
End If

If Check4.Value = 0 Then
FrmMain.Timer1.Enabled = False
DoEvents
FrmMain.StatusBar1.Panels(4).Picture = FrmMain.ImageList1.ListImages(4).Picture
End If

DoEvents
Me.Hide

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

Private Sub Command2_Click()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Command2_Click

Me.Hide

EXIT_Command2_Click:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_Command2_Click:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in Command2_Click" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_Command2_Click
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_Command2_Click

End Sub

Private Sub Form_Load()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Form_Load

AppDir = App.Path

FilePathName = AppDir + "\Settings.inf"
maximized = GetPrivateProfileString("settings", "maximized", "", FilePathName)
startupsound = GetPrivateProfileString("settings", "startupsound", "", FilePathName)
shutdownsound = GetPrivateProfileString("settings", "shutdownsound", "", FilePathName)
bandwidth = GetPrivateProfileString("settings", "bandwidth", "", FilePathName)

DoEvents
Me.Check1.Value = maximized
Me.Check2.Value = startupsound
Me.Check3.Value = shutdownsound
Me.Check4.Value = bandwidth

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
Private Function GetPrivateProfileString(ByVal szSection As String, ByVal szEntry As Variant, ByVal szDefault As String, ByVal szFileName As String) As String
   ' *** Get an entry in the inifile ***

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_GetPrivateProfileString

   Dim szTmp                     As String
   Dim nRet                      As Long

   If (IsNull(szEntry)) Then
      ' *** Get names of all entries in the named Section ***
      szTmp = String$(nBUFSIZEINIALL, 0)
      nRet = OSGetPrivateProfileString(szSection, 0&, szDefault, szTmp, nBUFSIZEINIALL, szFileName)
   Else
      ' *** Get the value of the named Entry ***
      szTmp = String$(nBUFSIZEINI, 0)
      nRet = OSGetPrivateProfileString(szSection, CStr(szEntry), szDefault, szTmp, nBUFSIZEINI, szFileName)
   End If
   GetPrivateProfileString = Left$(szTmp, nRet)

EXIT_GetPrivateProfileString:
   Exit Function

   ' #VBIDEUtilsERROR#
ERROR_GetPrivateProfileString:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in GetPrivateProfileString" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_GetPrivateProfileString
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_GetPrivateProfileString

End Function
Private Function GetProfileString(ByVal szSection As String, ByVal szEntry As Variant, ByVal szDefault As String) As String
   ' *** Get an entry in the WIN inifile ***

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_GetProfileString

   Dim szTmp                    As String
   Dim nRet                     As Long

   If (IsNull(szEntry)) Then
      ' *** Get names of all entries in the named Section ***
      szTmp = String$(nBUFSIZEINIALL, 0)
      nRet = OSGetProfileString(szSection, 0&, szDefault, szTmp, nBUFSIZEINIALL)
   Else
      ' *** Get the value of the named Entry ***
      szTmp = String$(nBUFSIZEINI, 0)
      nRet = OSGetProfileString(szSection, CStr(szEntry), szDefault, szTmp, nBUFSIZEINI)
   End If
   GetProfileString = Left$(szTmp, nRet)

EXIT_GetProfileString:
   Exit Function

   ' #VBIDEUtilsERROR#
ERROR_GetProfileString:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in GetProfileString" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_GetProfileString
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_GetProfileString

End Function
