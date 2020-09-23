VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check For Updates"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrTimeLeft 
      Interval        =   1000
      Left            =   0
      Top             =   1080
   End
   Begin VB.Timer tmrUpdateProgress 
      Interval        =   1
      Left            =   0
      Top             =   1440
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   120
      TabIndex        =   37
      Top             =   2760
      Width           =   1935
      Begin VB.CommandButton Command5 
         Caption         =   "&Stop"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Update"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "&File Download Progress"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2160
      TabIndex        =   24
      Top             =   2760
      Width           =   5175
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C00000&
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         ScaleHeight     =   225
         ScaleWidth      =   4905
         TabIndex        =   25
         Top             =   240
         Width           =   4935
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Total Size:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   600
         Width           =   750
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Recieved Size:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1800
         TabIndex        =   35
         Top             =   600
         Width           =   1050
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Speed:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3720
         TabIndex        =   34
         Top             =   600
         Width           =   510
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Time Remaining:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Width           =   1170
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Elapsed Time:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2640
         TabIndex        =   32
         Top             =   960
         Width           =   990
      End
      Begin VB.Label lblRemaining 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   31
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblElapsed 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   30
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblSpeed 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   29
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblRecieve 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   28
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblSize 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   27
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Visible         =   0   'False
         Width           =   4935
      End
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox TxtCurVersion 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox TxtUpdateDate 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   360
      Width           =   2535
   End
   Begin VB.TextBox TxtUpdateInfo 
      Height          =   1095
      Left            =   2160
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   1560
      Width           =   5175
   End
   Begin VB.TextBox TxtUpdateSize 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox TxtUpdateVersion 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check For Update"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6120
      TabIndex        =   15
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   6120
      TabIndex        =   14
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   6120
      TabIndex        =   13
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   6120
      TabIndex        =   12
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox TxtUpdateFileName 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   4800
      TabIndex        =   11
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   4800
      TabIndex        =   10
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   0
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   80
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "http://www.croftssoftware.com"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      MouseIcon       =   "FrmUpdate.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   39
      Top             =   4920
      Width           =   7215
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "If you have trouble using the new updater just click on this link to download the update file."
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   4560
      Width           =   7215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   840
      Picture         =   "FrmUpdate.frx":0614
      Top             =   2040
      Width           =   480
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Update Release Date"
      Height          =   255
      Left            =   4800
      TabIndex        =   21
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Current Version"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Update Info."
      Height          =   255
      Left            =   2160
      TabIndex        =   20
      Top             =   1320
      Width           =   5175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "File Size"
      Height          =   255
      Left            =   4800
      TabIndex        =   19
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Update Version"
      Height          =   255
      Left            =   2160
      TabIndex        =   18
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   4200
      Width           =   7215
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Update File Name"
      Height          =   255
      Left            =   2160
      TabIndex        =   16
      Top             =   720
      Width           =   2535
   End
End
Attribute VB_Name = "FrmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private FilePathName As String
Private Filename As String
Private FormName As String

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
Private NewVersion As String
Private OldVersion As String

Dim DATA As String
Dim Percent%
Dim BeginTransfer As Single
Dim BytesAlreadySent As Single
Dim BytesRemaining As Single
Dim Header As Variant
Dim Status As String
Dim TransferRate As Single

Dim Step1 As Boolean
Dim Step2 As Boolean

    Dim lTimer As Long
    Dim strRemote As String
    Dim strLocal As String
Dim NewVer As String
Dim Oldver As String
Dim URL As String
Dim AppDir As String
Dim YourVersion As String
Dim DOR As String
Dim FileSize As String
Dim WhatNew As String

Dim X As Long
Dim xx As Long
Dim PingTimes As Long
Dim Speed As Long
Dim IP As String
Dim KeepGoing As Long
Dim TotalNum As Long
Dim iReturn As Long, sLowByte As String, sHighByte As String
Dim sMsg As String, HostLen As Long
Dim HOSTENT As HOSTENT, PointerToPointer As Long, ListAddress As Long
Dim WSADATA As WSADATA, DotA As Long, DotAddr As String, ListAddr As Long
Dim MaxUDP As Long, MaxSockets As Long, i As Long
Dim Description As String
Dim ExitTheFor As Long
' Ping Variables
Dim bReturn As Boolean, hIP As Long
Dim szBuffer As String
Dim addr As Long
Dim RCode As String
Dim RespondingHost As String
' TRACERT Variables
Dim TraceRT As Boolean
Dim TTL As Integer
' WSock32 Constants
Const WS_VERSION_MAJOR = &H101 \ &H100 And &HFF&
Const WS_VERSION_MINOR = &H101 And &HFF&
Const MIN_SOCKETS_REQD = 0
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Function ConvertTime(ByVal TheTime As Single) As String

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_ConvertTime

    Dim NewTime                         As String
    Dim Sec                             As Single
    Dim Min                             As Single
    Dim H                               As Single
    If TheTime > 60 Then
        Sec = TheTime
        Min = Sec / 60
        Min = Int(Min)
        Sec = Sec - Min * 60
        H = Int(Min / 60)
        Min = Min - H * 60
        NewTime = H & ":" & Min & ":" & Sec
        If H < 0 Then H = 0
        If Min < 0 Then Min = 0
        If Sec < 0 Then Sec = 0
        NewTime = Format(NewTime, "HH:MM:SS")
        ConvertTime = NewTime
    End If
    If TheTime < 60 Then
        NewTime = "00:00:" & TheTime
        NewTime = Format(NewTime, "HH:MM:SS")
        ConvertTime = NewTime
    End If

EXIT_ConvertTime:
   Exit Function

   ' #VBIDEUtilsERROR#
ERROR_ConvertTime:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in ConvertTime" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_ConvertTime
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_ConvertTime

End Function

Public Function RunUpdate(UpdateURL As String)

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_RunUpdate

HyperJump UpdateURL

EXIT_RunUpdate:
   Exit Function

   ' #VBIDEUtilsERROR#
ERROR_RunUpdate:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in RunUpdate" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_RunUpdate
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_RunUpdate

End Function
Private Function HyperJump(ByVal URL As String) As Long

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_HyperJump

    HyperJump = ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)

EXIT_HyperJump:
   Exit Function

   ' #VBIDEUtilsERROR#
ERROR_HyperJump:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in HyperJump" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_HyperJump
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_HyperJump

End Function
Private Sub Command2_Click()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Command2_Click

Command2.Enabled = False
Step2 = True

Label3.Caption = "Status: Downloading Updates..."
If Text6.text = "Yes" Then
MsgBox "This update requires that the program not be running." & vbCrLf & "When the Update is downloaded and started this program will close.", vbExclamation
End If
Winsock.Close
FilePathName = App.Path & "\Updates\" & Text5.text
StartUpdate "http://" & Text1.text & "/" & Text2.text & Text5.text
DoEvents
lblStatus.Visible = False
Picture1.Visible = True
Winsock.connect strSvrURL, 80
DoEvents

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

Private Sub Command3_Click()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Command3_Click

Unload Me

EXIT_Command3_Click:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_Command3_Click:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in Command3_Click" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_Command3_Click
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_Command3_Click

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
Private Sub Command1_Click()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Command1_Click

Step1 = True

DoEvents
Call Check_Status
DoEvents
If Text7.text = "Online" Then
Label3.Caption = "Status: Connecting..."
DoEvents
NewVer = "none"
Oldver = "none"
YourVersion = TxtCurVersion.text


'Gets your Version
Oldver = YourVersion
Winsock.Close
FilePathName = App.Path & "\Updates\Update.inf"
StartUpdate "http://" & Text1.text & "/" & Text2.text & "Update.inf"
DoEvents
lblStatus.Visible = False
Picture1.Visible = True
Winsock.connect strSvrURL, 80
DoEvents

Else
Label3.Caption = "Status: Update Site Is Unavailable"
End If


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
Private Sub Command5_Click()
Command2.Enabled = True
If Winsock.State > 0 Then
Winsock.Close
Label3.Caption = "Status: Transfer Aborted!"
MsgBox "Transfer Aborted!", vbExclamation, "Aborted"
Reset
End If
End Sub

Private Sub Form_Load()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Form_Load

Dim intfile As Integer
Dim pass As String
Dim pass2 As String
TxtCurVersion.text = App.Major & "." & App.Minor & "." & App.Revision
intfile = FreeFile
  Open App.Path & "\Updates\UpdateSettings.ini" For Input As #intfile
  Input #intfile, pass
  Input #intfile, pass2
  Text1.text = pass
  Text2.text = pass2
  Close #intfile
  DoEvents
   
Dim mWSD As WSADataType
lV = WSAStartup(&H202, mWSD)
DoEvents
vbWSAStartup
vbWSACleanup
DoEvents

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
Public Sub Check_Status()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Check_Status

If gethostbyname(Text1.text) = 0 Then
Text7.text = "Offline"
Exit Sub
End If
    Speed = 0
    PingTimes = 0
    szBuffer = Space(Val("32"))
    DoEvents
    vbWSAStartup
    DoEvents
    DoEvents
    vbGetHostByName
    vbIcmpCreateFile
    DoEvents
    pIPo2.TTL = Trim$(255)
    '
    For Times = 1 To "1"
    If ExitTheFor = 1 Then ExitTheFor = 0: Exit For
    vbIcmpSendEcho
    DoEvents
    Next
    DoEvents
    vbIcmpCloseHandle
    vbWSACleanup
    On Error GoTo skipit
    'Speed = Speed / PingTimes
    Exit Sub
skipit:

EXIT_Check_Status:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_Check_Status:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in Check_Status" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_Check_Status
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_Check_Status

End Sub
Public Sub GetRCode()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_GetRCode

RCode = ""
    If pIPe.Status = 0 Then RCode = "Success"
    If pIPe.Status = 11001 Then RCode = "Buffer too Small"
    If pIPe.Status = 11002 Then RCode = "Destination Unreahable"
    If pIPe.Status = 11003 Then RCode = "Dest Host Not Reachable"
    If pIPe.Status = 11004 Then RCode = "Dest Protocol Not Reachable"
    If pIPe.Status = 11005 Then RCode = "Dest Port Not Reachable"
    If pIPe.Status = 11006 Then RCode = "No Resources Available"
    If pIPe.Status = 11007 Then RCode = "Bad Option"
    If pIPe.Status = 11008 Then RCode = "Hardware Error"
    If pIPe.Status = 11009 Then RCode = "Packet too Big"
    If pIPe.Status = 11010 Then RCode = "Reqested Timed Out"
    If pIPe.Status = 11011 Then RCode = "Bad Request"
    If pIPe.Status = 11012 Then RCode = "Bad Route"
    If pIPe.Status = 11014 Then RCode = "TTL Exprd Reassemb"
    If pIPe.Status = 11015 Then RCode = "Parameter Problem"
    If pIPe.Status = 11016 Then RCode = "Source Quench"
    If pIPe.Status = 11017 Then RCode = "Option too Big"
    If pIPe.Status = 11018 Then RCode = "Bad Destination"
    If pIPe.Status = 11019 Then RCode = "Address Deleted"
    If pIPe.Status = 11020 Then RCode = "Spec MTU Change"
    If pIPe.Status = 11021 Then RCode = "MTU Change"
    If pIPe.Status = 11022 Then RCode = "Unload"
    If pIPe.Status = 11050 Then RCode = "General Failure"

    DoEvents
DoEvents
        If RCode <> "" Then
        DoEvents
            If RCode = "Success" Then
                'Speed = Speed + Val(Trim$(CStr(pIPe2.RoundTripTime)))
                DoEvents
                Text7.text = "Online"
            Exit Sub
            End If
            DoEvents
            KeepGoing = 1
            Text7.text = RCode
            DoEvents
        Else
        DoEvents
            KeepGoing = 1
            Text7.text = RCode
            DoEvents
        End If

EXIT_GetRCode:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_GetRCode:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in GetRCode" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_GetRCode
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_GetRCode

    End Sub


Public Sub vbGetHostByName()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_vbGetHostByName

    Dim szString As String
    
    Host = Trim$(Text7.text)
    szString = String(64, &H0)
    Host = Host + Right$(szString, 64 - Len(Host))

    If gethostbyname(Host) = SOCKET_ERROR Then
        sMsg = "Winsock Error" & Str$(WSAGetLastError())
        Label3.Caption = "Status: " & sMsg
        ExitTheFor = 1
    Else
        PointerToPointer = gethostbyname(Host) ' Get the pointer to the address of the winsock hostent structure
        CopyMemory HOSTENT.h_name, ByVal _
        PointerToPointer, Len(HOSTENT) ' Copy Winsock structure to the VisualBasic structure
        ListAddress = HOSTENT.h_addr_list ' Get the ListAddress of the Address List
        CopyMemory ListAddr, ByVal ListAddress, 4 ' Copy Winsock structure To the VisualBasic structure
        CopyMemory IPLong2, ByVal ListAddr, 4 ' Get the first list entry from the Address List
        CopyMemory addr, ByVal ListAddr, 4
        IP = Trim$(CStr(Asc(IPLong2.Byte4)) + "." + CStr(Asc(IPLong2.Byte3)) _
        + "." + CStr(Asc(IPLong2.Byte2)) + "." + CStr(Asc(IPLong2.Byte1)))
    End If

EXIT_vbGetHostByName:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_vbGetHostByName:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in vbGetHostByName" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_vbGetHostByName
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_vbGetHostByName

End Sub


Public Sub vbGetHostName()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_vbGetHostName
    
    Host = String(64, &H0)
    


    If gethostname(Host, HostLen) = SOCKET_ERROR Then
        sMsg = "WSock32 Error" & Str$(WSAGetLastError())
        Label3.Caption = "Status: " & sMsg
        ExitTheFor = 1
    Else
        Host = Left$(Trim$(Host), Len(Trim$(Host)) - 1)
        Text1.text = Host
    End If

EXIT_vbGetHostName:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_vbGetHostName:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in vbGetHostName" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_vbGetHostName
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_vbGetHostName

End Sub


Public Sub vbIcmpSendEcho()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_vbIcmpSendEcho

    Dim NbrOfPkts As Integer
    For NbrOfPkts = 1 To Trim$(1)

        DoEvents
            bReturn = IcmpSendEcho(hIP, addr, szBuffer, Len(szBuffer), pIPo2, pIPe2, Len(pIPe2) + 8, 2700)
           DoEvents
            If bReturn Then
                If KeepGoing = 1 Then KeepGoing = 0: Exit For
                PingTimes = PingTimes + 1
                DoEvents
                RespondingHost = CStr(pIPe2.Address(0)) + "." + CStr(pIPe2.Address(1)) + "." + CStr(pIPe2.Address(2)) + "." + CStr(pIPe2.Address(3))
                GetRCode
            Else
                Text7.text = "Offline"
            End If
        Next NbrOfPkts

EXIT_vbIcmpSendEcho:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_vbIcmpSendEcho:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in vbIcmpSendEcho" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_vbIcmpSendEcho
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_vbIcmpSendEcho

    End Sub


Sub vbWSAStartup()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_vbWSAStartup

Dim wsAdata2 As WSADataType
    iReturn = WSAStartup(&H101, wsAdata2)


    If iReturn <> 0 Then ' If WSock32 error, then tell me about it
        Label3.Caption = "Status: WSock32.dll is Not responding!"
        ExitTheFor = 1
    End If


    If LoByte(wsAdata2.wversion) < WS_VERSION_MAJOR Or (LoByte(wsAdata2.wversion) = WS_VERSION_MAJOR And HiByte(wsAdata2.wversion) < WS_VERSION_MINOR) Then
        sHighByte = Trim$(Str$(HiByte(wsAdata2.wversion)))
        sLowByte = Trim$(Str$(LoByte(wsAdata2.wversion)))
        sMsg = "WinSock Version " & sLowByte & "." & sHighByte
        sMsg = sMsg & " is Not supported "
        Label3.Caption = "Status: " & sMsg
        ExitTheFor = 1
        End
    End If


    If wsAdata2.iMaxSockets < MIN_SOCKETS_REQD Then
        sMsg = "This application requires a minimum of "
        sMsg = sMsg & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
            Label3.Caption = "Status: " & sMsg
            ExitTheFor = 1
        End
    End If
    
    MaxSockets = wsAdata2.iMaxSockets


    If MaxSockets < 0 Then
        MaxSockets = 65536 + MaxSockets
    End If
    MaxUDP = wsAdata2.iMaxUdpDg


    If MaxUDP < 0 Then
        MaxUDP = 65536 + MaxUDP
    End If
    
    Description = ""


    'For i = 0 To WSADESCRIPTION_LEN
    '    If wsAdata2.szDescription(i) = 0 Then Exit For
    '    Description = Description + Chr$(wsAdata2.szDescription(i))
    'Next i
    Status = ""


    'For i = 0 To WSASYS_STATUS_LEN
    '    If wsAdata2.szSystemStatus(i) = 0 Then Exit For
    '    Status = Status + Chr$(wsAdata2.szSystemStatus(i))
    'Next i

EXIT_vbWSAStartup:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_vbWSAStartup:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in vbWSAStartup" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_vbWSAStartup
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_vbWSAStartup

End Sub


Public Function HiByte(ByVal wParam As Integer)

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_HiByte

    HiByte = wParam \ &H100 And &HFF&

EXIT_HiByte:
   Exit Function

   ' #VBIDEUtilsERROR#
ERROR_HiByte:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in HiByte" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_HiByte
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_HiByte

End Function


Public Function LoByte(ByVal wParam As Integer)

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_LoByte

    LoByte = wParam And &HFF&

EXIT_LoByte:
   Exit Function

   ' #VBIDEUtilsERROR#
ERROR_LoByte:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in LoByte" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_LoByte
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_LoByte

End Function


Public Sub vbWSACleanup()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_vbWSACleanup

    iReturn = WSACleanup()

EXIT_vbWSACleanup:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_vbWSACleanup:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in vbWSACleanup" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_vbWSACleanup
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_vbWSACleanup

End Sub


Public Sub vbIcmpCloseHandle()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_vbIcmpCloseHandle

    bReturn = IcmpCloseHandle(hIP)

EXIT_vbIcmpCloseHandle:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_vbIcmpCloseHandle:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in vbIcmpCloseHandle" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_vbIcmpCloseHandle
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_vbIcmpCloseHandle

End Sub


Public Sub vbIcmpCreateFile()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_vbIcmpCreateFile

    hIP = IcmpCreateFile()

EXIT_vbIcmpCreateFile:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_vbIcmpCreateFile:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in vbIcmpCreateFile" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_vbIcmpCreateFile
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_vbIcmpCreateFile

End Sub


Public Sub Reset()
DATA = ""
Percent = 0
BeginTransfer = 0
BytesAlreadySent = 1
BytesRemaining = 0
Status = ""
Header = ""
RESUMEFILE = False
UpdateProgress Picture1, 0
End Sub
Public Function StartUpdate(strURL As String)
BytesAlreadySent = 1
If strURL = "" Then Exit Function
URL = strURL
Dim pos%, Length%, NextPos%, LENGTH2%, POS2%, POS3%
    pos = InStr(strURL, "://") 'Record position of ://
    LENGTH2 = Len("://") 'Record the length of it
    Length = Len(strURL) 'Length of the entire url
        If InStr(strURL, "://") Then  ' check if they entered the http:// or ftp://
        strURL = Right(strURL, Length - LENGTH2 - pos + 1) ' remove http:// or ftp://
        End If
            If InStr(strURL, "/") Then 'looks for the first / mark going from left to right
            POS2 = InStr(strURL, "/") 'gets the position of the / mark
'-----------------GET THE FILENAME-------------
            Dim StrFile$: StrFile = strURL 'load the variables into each other
            Do Until InStr(StrFile, "/") = 0 'Do the loop until all is left is the filename
            LENGTH2 = Len(StrFile) 'get the length of the filename every time its passed over by the loop
            POS3 = InStr(StrFile, "/") 'find the / mark
            StrFile = Right(strURL, LENGTH2 - POS3) 'slash it down removing everything before the / mark including the / mark...
            Loop
            Filename = StrFile
'----------------END GET FILE NAME--------------
            strSvrURL = Left(strURL, POS2 - 1) 'removes everything after the / mark leaving just the server name as the end result
            End If
'-----------END TRIM THE URL FOR THE SERVER NAME-----------

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Command2.Enabled = True
If Winsock.State > 0 Then
Winsock.Close
Reset
End If
End Sub

Private Sub Label9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Call ShellExecute(hwnd, "Open", Label9.Caption, "", App.Path, 1)

End Sub

Private Sub Winsock_Close()
If Step2 = True Then
Label3.Caption = "Status: Download Complete"
RunUpdate App.Path & "\Updates\" & Text5.text
DoEvents
If Text6.text = "Yes" Then
End
End If
Step2 = False
End If

If Step1 = True Then
'State & Access 'Version.inf' file
FilePathName = App.Path & "\Updates\Update.inf"
NewVer = GetPrivateProfileString("Version", "Version", "", FilePathName)
NewVersion = NewVer
DOR = GetPrivateProfileString("Version", "DOR", "", FilePathName)
FileSize1 = GetPrivateProfileString("Version", "Filesize1", "", FilePathName)
FileSize2 = GetPrivateProfileString("Version", "Filesize2", "", FilePathName)
WhatNew = GetPrivateProfileString("Version", "Whatsnew", "", FilePathName)
Downloadsite = GetPrivateProfileString("Version", "DownloadSite", "", FilePathName)
DownloadPath = GetPrivateProfileString("Version", "DownloadPath", "", FilePathName)
DownloadFile = GetPrivateProfileString("Version", "DownloadFile", "", FilePathName)
CloseProgramBeforeUpdate = GetPrivateProfileString("Version", "CloseProgramBeforeUpdate", "", FilePathName)

'Compare for newer version
If Oldver >= NewVer Then
Command1.Enabled = False
TxtUpdateVersion.text = NewVer
TxtUpdateDate.text = DOR
TxtUpdateFileName = DownloadFile
TxtUpdateSize.text = FileSize1 & " " & FileSize2
TxtUpdateInfo.text = WhatNew
Text3.text = Downloadsite
Text4.text = DownloadPath
Text5.text = DownloadFile
Text6.text = CloseProgramBeforeUpdate
DoEvents
Label9.Caption = "http://" & Text1.text & "/" & Text2.text & Text5.text
Label3.Caption = "Status: You are Up to Date"
Else
Command1.Enabled = False
TxtUpdateVersion.text = NewVer
TxtUpdateDate.text = DOR
TxtUpdateFileName = DownloadFile
TxtUpdateSize.text = FileSize1 & " " & FileSize2
TxtUpdateInfo.text = WhatNew
Text3.text = Downloadsite
Text4.text = DownloadPath
Text5.text = DownloadFile
Text6.text = CloseProgramBeforeUpdate
Command2.Enabled = True
Command5.Enabled = True
DoEvents
Label9.Caption = "http://" & Text1.text & "/" & Text2.text & Text5.text
Label3.Caption = "Status: Update Available"
End If
Step1 = False
End If

End Sub

Private Sub Winsock_Connect()
Dim strCommand As String
 
 On Error Resume Next
 
 'Fixed the bug with unix servers thanks to Michael Pauletta
 
 If Not Unix Then
  strCommand = "GET " + URL + " HTTP/1.0" + vbCrLf 'tells server to GET the file if you just want the header info and not the data change "GET " to "HEAD
 Else
    strCommand = "GET " + "/" + Filename + " HTTP/1.0" + vbCrLf 'tells server to GET the file if you just want the header info and not the data change "GET "to "HEAD "
 End If
 
     strCommand = strCommand + "Accept: *.*, */*" + vbCrLf
 If RESUMEFILE = True Then strCommand = strCommand + "Range: bytes=" & FileLength & "-" & vbCrLf
    strCommand = strCommand + "User-Agent: Conquest" & vbCrLf
 
 
 If Not Unix Then
    strCommand = strCommand + "Referer: " & strSvrURL & vbCrLf
 Else
    strCommand = strCommand + "Host: " & strSvrURL & vbCrLf
 End If
 
 
    strCommand = strCommand + vbCrLf
    Winsock.SendData strCommand 'sends a header to the server instructing it what to do!
    BeginTransfer = Timer 'start timer for transfer rate

End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
Winsock.GetData DATA, vbString
If InStr(DATA, "Content-Type:") Then 'find out if this chunk has the header..you can change that to anything that the header contains
        
        If RESUMEFILE = True Then 'check to see if its gonna resume ok or not..This is actually the worst way to check this.
            If InStr(DATA, "HTTP/1.1 206 Partial Content") = 0 Then
            MsgBox "Server did not accept resuming.", vbCritical, "No Resuming Support"
            Exit Sub
            End If
        End If
        
    If InStr(DATA, "404 Not Found") > 0 Then
   'If Not Unix Then
   ' Unix = True
    'Winsock.connect strSvrURL, 80
   ' Exit Sub
   'End If
   Unix = False
   MsgBox "File not found on this server.", vbCritical, "File Not Found"
   Exit Sub
    End If
        
    Dim pos%, Length%, HEAD$
    pos = InStr(DATA, vbCrLf & vbCrLf) ' find out where the header and the data is split apart
    Length = Len(DATA) 'get the length of the data chunk
    HEAD = Left(DATA, pos - 1) 'Get the header from the chunk of data and ignore the data content
    DATA = Right(DATA, Length - pos - 3) 'Get the data from the first chunk that contains the header also
    Header = Header & HEAD 'Append the header to header text box

If RESUMEFILE = True Then
BytesAlreadySent = FileLength + 1
BytesRemaining = GETDATAHEAD(Header, "Content-Length:")
BytesRemaining = BytesRemaining + FileLength
Else
BytesRemaining = GETDATAHEAD(Header, "Content-Length:")
End If
txtHead = Header
End If

'-----------BEGIN WRITE CHUNK TO FILE CODE--------
        Open FilePathName For Binary Access Write As #1 'opens file for output
        Put #1, BytesAlreadySent, DATA 'writes data to the end of file
        BytesAlreadySent = Seek(1)
        Close #1 'close file for now until next data chunk is available
'--------------------------------------------------

'Lets explain this a bit..The variable BeginTransfer is given the starting value of the
'timer which in case you dont know is the amount of seconds til midnight but that has
'nothing to do with this. Anyways so its given the amount for the start time and then
'when this event below is fired for the first time the timer will be given the value again
'since your system clock was ticking along while the operation between the two of these
'events happened the number will be different.  The two values are subtracted and divided
'by the amount recieved and then by 1000 and put into a readable format
If RESUMEFILE = False Then
'This is pretty straightforward if you ever taken math before you can tell what im doing!
TransferRate = Format(Int(BytesAlreadySent / (Timer - BeginTransfer)) / 1024, "####.00")
Else
'If you dont subtract the difference you will get a really large and odd download speed hehe.
TransferRate = Format(Int((BytesAlreadySent - FileLength) / (Timer - BeginTransfer)) / 1024, "####.00")
End If
End Sub
Private Sub tmrUpdateProgress_Timer()
On Error Resume Next
If BytesAlreadySent > 0 And BytesRemaining > 0 Then
lblRecieve = File_ByteConversion(BytesAlreadySent - 1)
lblSize = File_ByteConversion(BytesRemaining)
Percent = Format((BytesAlreadySent / BytesRemaining) * 100, "00") 'calculates the percentage completed
UpdateProgress Picture1, Percent 'updates progress bar with new percentage rate
End If
End Sub
Private Sub tmrTimeLeft_Timer()
'On Error Resume Next
If BytesRemaining > 0 And BytesAlreadySent > 0 Then
If BytesRemaining <= BytesAlreadySent Then
lblSpeed = 0

lblElapsed = Format(Hr & ":" & Min & ":" & Sec, "HH:MM:SS")

Picture1.Visible = False
lblStatus.Visible = True
lblStatus.Caption = "Download Completed"
Reset
Else
    Sec = Sec + 1
    If Sec >= 60 Then
    Sec = 0
    Min = Min + 1
    ElseIf Min >= 60 Then
    Min = 0
    Hr = Hr + 1
    End If

lblElapsed = Format(Hr & ":" & Min & ":" & Sec, "HH:MM:SS")
'The reason I divide the difference of bytesalreadysent and bytesremaining is becuase they are in bytes right now.. I want it to be in KB so it can be Kbps and not bps
lblRemaining = ConvertTime(Int(((BytesRemaining - BytesAlreadySent) / 1024) / TransferRate))
lblSpeed = TransferRate & " KB"
End If

End If
End Sub

